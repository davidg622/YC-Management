import re
import traceback
from io import BytesIO
import zipfile

import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill, numbers
from openpyxl.utils import get_column_letter

# ============================================================
# Yellow Cluster: Budget Summary Generator
# - Preview both files BEFORE merge
# - Let user SELECT merge columns for Master and Award
# - Show match-rate diagnostics (keys overlap + unmatched samples)
# - Generate ONE ZIP containing ONE Excel file PER PI
#   (no organization sorting / no folders)
# - Sort each PI report by Allocated Budget* (descending)
# ============================================================

# -----------------------------
# PAGE CONFIG (must be first Streamlit call)
# -----------------------------
st.set_page_config(
    page_title="Yellow Cluster: Budget Summary Generator",
    page_icon="🐄",
    layout="wide",
)

# -----------------------------
# SETTINGS
# -----------------------------
SUMMARY_SHEET_NAME = "Summary"
HEADER_ROW_INDEX = 17  # 0-based index: Excel row 18 (i.e., delete first 17 rows)

# Master (Aggie) column identifiers we use for the final output
PI_COL_NAME = "Project Principal Investigator"
PROJECT_COL_NAME = "Project Number"
TASK_NAME_COL_NAME = "Task Name"
TASK_NUMBER_COL_NAME = "Task Number"
STATUS_COL_NAME = "Project Status"

ALLOC_BUDGET_NET_COL = "Allocated Budget*"
CURRENT_BAL_NET_COL = "Current Balance*"

# -----------------------------
# Helpers
# -----------------------------
def normalize_columns(cols):
    return [str(c).replace("\xa0", " ").strip() for c in cols]


def safe_str(x) -> str:
    if x is None:
        return ""
    if isinstance(x, float) and pd.isna(x):
        return ""
    return str(x).replace("\xa0", " ").strip()


def canon_key(x) -> str:
    """
    Canonicalize merge keys without destroying alphanumeric IDs.
    - Trim
    - Remove spaces/punctuation
    - Uppercase
    - If purely numeric-like, normalize 12345.0 -> 12345
    """
    s = safe_str(x)
    if not s:
        return ""

    # numeric-like?
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))
    except Exception:
        pass

    cleaned = re.sub(r"[^A-Za-z0-9]", "", s).upper()
    return cleaned


def find_column_by_exact_or_keywords(columns, target_name, keywords=None):
    columns = list(columns)
    if target_name in columns:
        return target_name

    if keywords:
        lowered = [c.lower() for c in columns]
        for col, low in zip(columns, lowered):
            if all(k.lower() in low for k in keywords):
                return col

    raise KeyError(
        f"Could not find a suitable column for '{target_name}'. Available columns: {columns}"
    )


def normalize_pi_last_first(pi_val: str) -> str:
    s = safe_str(pi_val)
    if not s:
        return ""
    if "," in s:
        return s
    parts = s.split()
    if len(parts) >= 2:
        return f"{parts[-1]}, {' '.join(parts[:-1])}"
    return s


def make_safe_filename_fragment(name: str) -> str:
    """
    Filesystem-safe fragment for filenames.
    """
    frag = safe_str(name)
    frag = re.sub(r'[\/:*?"<>|]+', "_", frag)
    frag = frag.strip().strip(".")
    return frag[:120] if frag else "PI"


def apply_currency_format(wb, ws_name, columns):
    ws = wb[ws_name]
    header_row = next(ws.iter_rows(min_row=1, max_row=1))
    pos = {}
    for cell in header_row:
        if cell.value in columns:
            pos[cell.value] = cell.column_letter

    for _, col_letter in pos.items():
        for cell in ws[col_letter]:
            if cell.row == 1:
                continue
            cell.number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE


def style_sheet(wb, ws_name, currency_cols, footnote_text, hide_indirect=True):
    ws = wb[ws_name]

    apply_currency_format(wb, ws_name, currency_cols)

    header_font = Font(bold=True, size=12)
    body_font = Font(size=11)

    header_row = next(ws.iter_rows(min_row=1, max_row=1))
    indirect_col_letter = None
    balance_col_letter = None

    for cell in header_row:
        cell.font = header_font
        if cell.value == CURRENT_BAL_NET_COL:
            balance_col_letter = cell.column_letter
            cell.fill = PatternFill(start_color="FFFAD7", end_color="FFFAD7", fill_type="solid")
        if cell.value == "Indirect Rate":
            indirect_col_letter = cell.column_letter

    # alternating row shading
    fill_green = PatternFill(start_color="FFE6F4EA", end_color="FFE6F4EA", fill_type="solid")
    for r in range(2, ws.max_row + 1):
        for cell in ws[r]:
            cell.font = body_font
            if r % 2 == 0:
                cell.fill = fill_green

    # color code current balance*
    if balance_col_letter:
        for cell in ws[balance_col_letter]:
            if cell.row == 1:
                continue
            color = None
            try:
                v = float(cell.value)
                if v < 0:
                    color = "8B0000"
                elif v > 0:
                    color = "004B00"
            except Exception:
                pass
            cell.font = Font(bold=True, size=11, color=color) if color else Font(bold=True, size=11)

    # widths
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(10, min(max_len + 2, 55))

    # hide indirect if requested
    if hide_indirect and indirect_col_letter:
        ws.column_dimensions[indirect_col_letter].hidden = True

    # footnote
    footer_row = ws.max_row + 2
    ws[f"A{footer_row}"] = footnote_text
    ws[f"A{footer_row}"].font = Font(italic=True, size=10)


def read_aggy_master(master_bytes: bytes) -> pd.DataFrame:
    """
    Reads Aggie Enterprise Database:
      - Sheet 'Summary'
      - Uses row 18 as header (0-based index 17)
    """
    df_raw = pd.read_excel(BytesIO(master_bytes), sheet_name=SUMMARY_SHEET_NAME, header=None)
    header = df_raw.iloc[HEADER_ROW_INDEX]
    df = df_raw.iloc[HEADER_ROW_INDEX + 1 :].copy()
    df.columns = header
    df = df.dropna(how="all")
    df.columns = normalize_columns(df.columns)
    return df


def read_award(award_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    xl = pd.ExcelFile(BytesIO(award_bytes))
    df = pd.read_excel(xl, sheet_name=sheet_name)
    df.columns = normalize_columns(df.columns)
    return df


def build_pi_zip(df_out: pd.DataFrame, pi_col: str, hide_indirect: bool, report_label: str) -> bytes:
    """
    Create one Excel file per PI and return a ZIP (bytes).
    Each PI's sheet is sorted by Allocated Budget* descending.
    """
    df_out = df_out.copy()
    df_out[pi_col] = df_out[pi_col].apply(normalize_pi_last_first)

    unique_pis = [p for p in df_out[pi_col].dropna().unique().tolist() if safe_str(p)]
    unique_pis_sorted = sorted(unique_pis, key=lambda s: safe_str(s).lower())

    zip_buf = BytesIO()
    used_names = set()

    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for pi in unique_pis_sorted:
            group = df_out[df_out[pi_col] == pi].copy()
            if group.empty:
                continue

            # --- SORT: largest Allocated Budget* at top ---
            if ALLOC_BUDGET_NET_COL in group.columns:
                group[ALLOC_BUDGET_NET_COL] = pd.to_numeric(group[ALLOC_BUDGET_NET_COL], errors="coerce")
                group = group.sort_values(by=ALLOC_BUDGET_NET_COL, ascending=False, na_position="last")

            # Footnote per PI (if one indirect rate)
            if "Indirect Rate" in group.columns:
                uniq = pd.Series(group["Indirect Rate"].dropna().unique())
                if len(uniq) == 1:
                    footnote = f"* Calculated minus the indirect costs (Indirect Rate = {float(uniq.iloc[0]):.2%})."
                else:
                    footnote = "* Calculated minus the indirect costs (Indirect Rate varies by project)."
            else:
                footnote = "* Calculated minus the indirect costs."

            currency_cols = [c for c in [ALLOC_BUDGET_NET_COL, CURRENT_BAL_NET_COL, "expenses"] if c in group.columns]

            xbuf = BytesIO()
            with pd.ExcelWriter(xbuf, engine="openpyxl") as writer:
                group.to_excel(writer, index=False, sheet_name="Budget Summary")
                wb = writer.book
                style_sheet(wb, "Budget Summary", currency_cols, footnote, hide_indirect=hide_indirect)
            xbuf.seek(0)

            safe_pi = make_safe_filename_fragment(pi)
            filename = f"{report_label} - {safe_pi}.xlsx"

            if filename in used_names:
                k = 2
                while True:
                    candidate = f"{report_label} - {safe_pi} ({k}).xlsx"
                    if candidate not in used_names:
                        filename = candidate
                        break
                    k += 1
            used_names.add(filename)

            zf.writestr(filename, xbuf.read())

    zip_buf.seek(0)
    return zip_buf.getvalue()


# -----------------------------
# UI
# -----------------------------
st.markdown(
    """
    <div style="padding: 1rem 1.25rem; border-radius: 12px; background: #01223d; color: white; margin-bottom: 1rem;">
      <div style="font-size: 1.35rem; font-weight: 700;">🐄 Yellow Cluster • Budget Report Generator</div>
      <div style="opacity: 0.85; margin-top: 0.25rem;">
        Preview both files, choose merge columns, validate match rate, then download one ZIP with one Excel file per PI.
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.expander("Debug options", expanded=False):
    show_trace = st.checkbox("Show full error trace", value=False)
    show_unmatched = st.checkbox("Show unmatched key samples", value=True)
    show_key_samples = st.checkbox("Show key samples from both files", value=True)

master_file = st.file_uploader("Upload Aggie Enterprise Database (Excel)", type=["xlsx"])
award_file = st.file_uploader("Upload Award Info (Excel)", type=["xlsx"])

date_pulled = st.text_input("Date Pulled (optional)", value="")
hide_indirect_in_output = st.checkbox("Hide 'Indirect Rate' column in PI files", value=True)

if master_file and award_file:
    try:
        master_bytes = master_file.getvalue()
        award_bytes = award_file.getvalue()

        # Read both
        df_master = read_aggy_master(master_bytes)

        xl_aw = pd.ExcelFile(BytesIO(award_bytes))
        award_sheet = st.selectbox("Award sheet to use", options=xl_aw.sheet_names, index=0)
        df_award = read_award(award_bytes, sheet_name=award_sheet)

        # Option: filter master to Active by default
        do_active_only = st.checkbox("Master: keep only Task Status == Active", value=True)
        status_col = find_column_by_exact_or_keywords(df_master.columns, STATUS_COL_NAME, keywords=["task", "status"])
        df_master_view = df_master[df_master[status_col] == "Active"].copy() if do_active_only else df_master.copy()

        st.markdown("### Preview (before merge)")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Aggie Enterprise (Master) preview**")
            st.dataframe(df_master_view.head(25), use_container_width=True)
        with c2:
            st.markdown("**Award Info preview**")
            st.dataframe(df_award.head(25), use_container_width=True)

        st.markdown("---")
        st.markdown("### Choose merge columns")

        default_master_merge = PROJECT_COL_NAME if PROJECT_COL_NAME in df_master_view.columns else df_master_view.columns[0]

        default_aw_merge = None
        for cand in ["Aggie Enterprise Project #", "AGGIE ENTERPRISE PROJECT #", "AGGIE ENTERPRISE PROJECT # "]:
            if cand in df_award.columns:
                default_aw_merge = cand
                break
        if default_aw_merge is None:
            default_aw_merge = df_award.columns[0]

        master_merge_col = st.selectbox(
            "Master merge column",
            options=list(df_master_view.columns),
            index=list(df_master_view.columns).index(default_master_merge) if default_master_merge in df_master_view.columns else 0,
        )
        award_merge_col = st.selectbox(
            "Award merge column",
            options=list(df_award.columns),
            index=list(df_award.columns).index(default_aw_merge) if default_aw_merge in df_award.columns else 0,
        )

        default_aw_rate = None
        for cand in ["INDIRECT RATE", "Indirect Rate", "Indirect rate"]:
            if cand in df_award.columns:
                default_aw_rate = cand
                break
        if default_aw_rate is None:
            indirect_candidates = [c for c in df_award.columns if "indirect" in c.lower()]
            default_aw_rate = indirect_candidates[0] if indirect_candidates else df_award.columns[-1]

        award_rate_col = st.selectbox(
            "Award indirect-rate column",
            options=list(df_award.columns),
            index=list(df_award.columns).index(default_aw_rate) if default_aw_rate in df_award.columns else 0,
        )

        # Merge preview metrics
        master_keys = df_master_view[master_merge_col].apply(canon_key)
        award_keys = df_award[award_merge_col].apply(canon_key)

        master_key_set = set(k for k in master_keys.unique() if k)
        award_key_set = set(k for k in award_keys.unique() if k)
        intersect = master_key_set.intersection(award_key_set)
        match_rate_unique = (len(intersect) / len(master_key_set)) if master_key_set else 0.0

        st.markdown("### Merge preview")
        st.write(f"**Unique master keys:** {len(master_key_set)}")
        st.write(f"**Unique award keys:** {len(award_key_set)}")
        st.write(f"**Key overlap (unique):** {len(intersect)}")
        st.write(f"**Approx. match rate (unique master keys found in award):** {match_rate_unique:.1%}")

        if show_key_samples:
            st.markdown("**Key samples (canonicalized)**")
            c3, c4 = st.columns(2)
            with c3:
                st.caption("Master key sample")
                st.code(", ".join(list(master_key_set)[:20]) if master_key_set else "(none)")
            with c4:
                st.caption("Award key sample")
                st.code(", ".join(list(award_key_set)[:20]) if award_key_set else "(none)")

        if show_unmatched:
            missing = sorted(list(master_key_set - award_key_set))[:40]
            if missing:
                st.warning("Some master keys were not found in award keys (sample):")
                st.code(", ".join(missing[:40]))

        st.markdown("---")
        st.markdown("### Generate Monthly Reports")

        if st.button("Generate ZIP (one Excel per PI)", type="primary"):
            df_work_full = df_master_view.copy()

            project_col = find_column_by_exact_or_keywords(df_work_full.columns, PROJECT_COL_NAME, keywords=["project", "number"])
            task_name_col = find_column_by_exact_or_keywords(df_work_full.columns, TASK_NAME_COL_NAME, keywords=["task", "name"])
            task_num_col = find_column_by_exact_or_keywords(df_work_full.columns, TASK_NUMBER_COL_NAME, keywords=["task", "number"])

            balance_candidates = [c for c in df_work_full.columns if str(c).startswith("Budget Balance")]
            if not balance_candidates:
                raise KeyError("Master is missing a column starting with 'Budget Balance'.")
            balance_col = balance_candidates[0]

            keep_cols = []
            for c in [
                PI_COL_NAME,
                project_col,
                "Project Name",
                "Project Manager",
                task_name_col,
                task_num_col,
                "Budget",
                "expenses",
                balance_col,
            ]:
                if c in df_work_full.columns and c not in keep_cols:
                    keep_cols.append(c)

            df_work = df_work_full[keep_cols].copy()

            df_work["_merge_key"] = df_work_full[master_merge_col].apply(canon_key)

            df_aw = df_award.copy()
            df_aw["_merge_key"] = df_aw[award_merge_col].apply(canon_key)

            df_aw_sub = df_aw[["_merge_key", award_rate_col]].copy()
            df_aw_sub = df_aw_sub.drop_duplicates(subset=["_merge_key"], keep="first")
            df_aw_sub = df_aw_sub.rename(columns={award_rate_col: "Indirect Rate"})

            df_merged = df_work.merge(df_aw_sub, on="_merge_key", how="left").drop(columns=["_merge_key"])

            matched_rows = df_merged["Indirect Rate"].notna().sum()
            total_rows = len(df_merged)
            st.info(f"Row-level merge match: {matched_rows}/{total_rows} ({(matched_rows/total_rows if total_rows else 0):.1%})")

            p = df_merged[project_col].apply(safe_str).str.replace(".0", "", regex=False)
            t = df_merged[task_num_col].apply(safe_str).str.replace(".0", "", regex=False)
            df_merged[project_col] = (p + "-" + t).str.strip("-")

            if "Project Name" in df_merged.columns:
                df_merged["Project Name"] = (
                    df_merged["Project Name"].apply(safe_str)
                    + " – "
                    + df_merged[task_name_col].apply(safe_str)
                )

            df_merged = df_merged.drop(columns=[task_name_col, task_num_col], errors="ignore")

            df_merged = df_merged.rename(columns={"Budget": "Allocated Budget", balance_col: "Current Balance"})

            df_merged["Indirect Rate"] = pd.to_numeric(df_merged["Indirect Rate"], errors="coerce").fillna(0.0)
            df_merged["Allocated Budget"] = pd.to_numeric(df_merged["Allocated Budget"], errors="coerce")
            df_merged["Current Balance"] = pd.to_numeric(df_merged["Current Balance"], errors="coerce")

            denom = 1.0 + df_merged["Indirect Rate"]
            df_merged[ALLOC_BUDGET_NET_COL] = df_merged["Allocated Budget"] / denom
            df_merged[CURRENT_BAL_NET_COL] = df_merged["Current Balance"] / denom

            df_merged = df_merged.drop(columns=["Allocated Budget", "Current Balance"], errors="ignore")

            date_label = safe_str(date_pulled)
            if date_label:
                df_merged["Date Pulled"] = date_label

            if PI_COL_NAME not in df_merged.columns:
                raise KeyError(f"PI column '{PI_COL_NAME}' not found in master. Columns: {list(df_merged.columns)}")
            df_merged[PI_COL_NAME] = df_merged[PI_COL_NAME].apply(normalize_pi_last_first)

            desired = [
                PI_COL_NAME,
                "Project Manager",
                "Date Pulled",
                project_col,
                "Project Name",
                ALLOC_BUDGET_NET_COL,
                CURRENT_BAL_NET_COL,
                "Indirect Rate",
                "expenses",
            ]
            desired = [c for c in desired if c in df_merged.columns]
            remaining = [c for c in df_merged.columns if c not in desired]
            df_out = df_merged[desired + remaining]

            report_label = "Budget Report"
            if date_label:
                report_label = f"{date_label} Budget Report"

            zip_bytes = build_pi_zip(
                df_out=df_out,
                pi_col=PI_COL_NAME,
                hide_indirect=hide_indirect_in_output,
                report_label=report_label,
            )

            st.success("ZIP generated!")
            st.download_button(
                "Download ZIP (PI files)",
                data=zip_bytes,
                file_name=f"{make_safe_filename_fragment(report_label)} - PI Files.zip",
                mime="application/zip",
            )

    except Exception as e:
        st.error(f"Error: {e}")
        if show_trace:
            st.code(traceback.format_exc())
else:
    st.info("Upload both files to preview and configure the merge.")
