import re
import traceback
import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile

from openpyxl.styles import numbers, PatternFill, Font
from openpyxl.utils import get_column_letter

# -----------------------------
# PAGE CONFIG (MUST BE FIRST Streamlit call)
# -----------------------------
st.set_page_config(
    page_title="Yellow Cluster: Monthly PI Budget Summary Generator",
    page_icon="🐄",
    layout="wide",
)

# -----------------------------
# SETTINGS – YOU CAN EDIT THESE
# -----------------------------
SUMMARY_SHEET_NAME = "Summary"
HEADER_ROW_INDEX = 17  # 0-based index: Excel row 18

PI_COL_NAME = "Project Principal Investigator"
PROJECT_COL_NAME = "Project Number"
TASK_NAME_COL_NAME = "Task Name"
TASK_NUMBER_COL_NAME = "Task Number"
STATUS_COL_NAME = "Task Status"
OWNING_ORG_COL_NAME = "Project Owning Organization"

BASE_OUTPUT_COLS = [
    "Task Name",
    "Task Number",
    "Project Number",
    "Project Name",
    "Project Manager",
    OWNING_ORG_COL_NAME,
    "Budget",
    "expenses",
]

AWARD_INFO_PROJECT_COL = "AGGIE ENTERPRISE PROJECT #"
AWARD_INFO_INDIRECT_COL = "INDIRECT RATE"

BALANCE_EX_INDIRECT_COL = "Current Balance*"


# -----------------------------
# Helper functions
# -----------------------------
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
        f"Could not find a suitable column for '{target_name}'. "
        f"Available columns: {columns}"
    )


def make_safe_sheet_name(name: str) -> str:
    sheet = str(name)
    for ch in r'\/:*?"<>[]':
        sheet = sheet.replace(ch, "_")
    return sheet[:31] if sheet else "Sheet"


def make_safe_filename_fragment(name: str) -> str:
    frag = str(name)
    for ch in r'\/:*?"<>|':
        frag = frag.replace(ch, "_")
    frag = frag.strip()
    return frag[:80] if frag else "PI"


def normalize_pi_name(pi: str) -> str:
    pi = str(pi).strip()
    if not pi:
        return "Unknown PI"
    if "," in pi:
        return pi

    parts = pi.split()
    if len(parts) >= 2:
        first = " ".join(parts[:-1])
        last = parts[-1]
        return f"{last}, {first}"
    return pi


def compute_org7(value) -> str:
    """
    Use the FIRST 7 DIGITS found in the value. This is robust to:
      - 1234567.0
      - '1234567 - Dept Name'
      - 'Org 1234567 Something'
    Fallback: first 7 characters if fewer than 7 digits exist.
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "UnknownOrg"

    s = str(value).strip()
    if not s:
        return "UnknownOrg"

    digits = re.findall(r"\d", s)
    if len(digits) >= 7:
        return "".join(digits[:7])

    # fallback: just take first 7 chars (trimmed)
    return s[:7] if len(s) >= 7 else "UnknownOrg"


def apply_currency_format(workbook, sheet_name, columns):
    ws = workbook[sheet_name]
    header_row = next(ws.iter_rows(min_row=1, max_row=1))

    header_to_col_letter = {}
    for cell in header_row:
        if cell.value in columns:
            header_to_col_letter[cell.value] = cell.column_letter

    for _, col_letter in header_to_col_letter.items():
        for cell in ws[col_letter]:
            if cell.row == 1:
                continue
            cell.number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE


def style_sheet(workbook, sheet_name, currency_cols, hide_indirect=True):
    ws = workbook[sheet_name]

    apply_currency_format(workbook, sheet_name, currency_cols)

    header_font = Font(bold=True, size=12)
    body_font = Font(size=11)

    header_row = next(ws.iter_rows(min_row=1, max_row=1))
    balance_col_letter = None
    indirect_col_letter = None

    for cell in header_row:
        cell.font = header_font
        if cell.value == BALANCE_EX_INDIRECT_COL:
            balance_col_letter = cell.column_letter
            cell.fill = PatternFill(start_color="FFFAD7", end_color="FFFAD7", fill_type="solid")
        if cell.value == "Indirect Rate":
            indirect_col_letter = cell.column_letter

    fill_green = PatternFill(start_color="FFE6F4EA", end_color="FFE6F4EA", fill_type="solid")
    for row_idx in range(2, ws.max_row + 1):
        for cell in ws[row_idx]:
            cell.font = body_font
            if row_idx % 2 == 0:
                cell.fill = fill_green

    if balance_col_letter is not None:
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
            except (TypeError, ValueError):
                pass

            if color:
                cell.font = Font(bold=True, size=11, color=color)
            else:
                cell.font = Font(bold=True, size=11)

    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value is not None:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(10, min(max_length + 2, 50))

    if indirect_col_letter is not None and hide_indirect:
        ws.column_dimensions[indirect_col_letter].hidden = True

    footer_row = ws.max_row + 2
    ws[f"A{footer_row}"] = "* Calculated minus the indirect costs."
    ws[f"A{footer_row}"].font = Font(italic=True, size=10)


def build_per_pi_org7_zip(df_merged, base_name, currency_cols):
    """
    ZIP structure:
      <Org7>/<Report Label> - <Last, First>.xlsx
    Each PI gets one file per Org7 (if PI spans multiple owning orgs).
    """
    zip_buf = BytesIO()
    used_paths = set()

    report_label = base_name.replace("_", " ").strip()

    required_helpers = {"_PI_stripped", "_Org7"}
    missing = required_helpers - set(df_merged.columns)
    if missing:
        raise ValueError(f"Missing helper columns needed for zipping: {missing}")

    # groupby without dropna kwarg for compatibility
    grouped = df_merged.groupby(["_Org7", "_PI_stripped"], sort=True)

    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for (org7, pi), group in grouped:
            if group.empty:
                continue

            org7 = str(org7) if org7 else "UnknownOrg"
            org7_safe = make_safe_filename_fragment(org7)

            pi_label = normalize_pi_name(pi)
            safe_pi = make_safe_filename_fragment(pi_label)

            filename = f"{report_label} - {safe_pi}.xlsx"
            zip_path = f"{org7_safe}/{filename}"

            if zip_path in used_paths:
                suffix = 2
                while True:
                    candidate_filename = f"{report_label} - {safe_pi} ({suffix}).xlsx"
                    candidate_path = f"{org7_safe}/{candidate_filename}"
                    if candidate_path not in used_paths:
                        zip_path = candidate_path
                        break
                    suffix += 1
            used_paths.add(zip_path)

            out = group.drop(
                columns=["_PI_stripped", "_Org7", OWNING_ORG_COL_NAME],
                errors="ignore",
            ).copy()

            pi_buf = BytesIO()
            with pd.ExcelWriter(pi_buf, engine="openpyxl") as writer:
                out.to_excel(writer, index=False, sheet_name="Budget Summary")
                wb = writer.book
                style_sheet(wb, "Budget Summary", currency_cols, hide_indirect=True)
            pi_buf.seek(0)

            zf.writestr(zip_path, pi_buf.read())

    zip_buf.seek(0)
    return zip_buf.getvalue()


@st.cache_data(show_spinner=False)
def process_workbooks_cached(master_bytes: bytes, award_bytes_list: list[bytes], date_pulled: str):
    """
    Cache-friendly wrapper: accept raw bytes (not UploadedFile objects).
    """
    return process_workbooks_bytes(master_bytes, award_bytes_list, date_pulled)


def process_workbooks_bytes(master_bytes: bytes, award_bytes_list: list[bytes], date_pulled=""):
    """
    Core processing logic.

    Returns:
        base_name,
        per_pi_zip_bytes,
        summary_info (dict)
    """
    base_name = "Budget_Report"
    prefix = (date_pulled or "").strip()
    if prefix:
        base_name = f"{prefix}_{base_name}"

    # Always read from fresh BytesIO to avoid pointer issues on rerun
    master_file = BytesIO(master_bytes)
    df_raw = pd.read_excel(master_file, sheet_name=SUMMARY_SHEET_NAME, header=None)

    header = df_raw.iloc[HEADER_ROW_INDEX]
    df = df_raw.iloc[HEADER_ROW_INDEX + 1 :].copy()
    df.columns = header
    df = df.dropna(how="all")
    df.columns = df.columns.astype(str).str.strip()

    # Identify key columns robustly
    pi_col = find_column_by_exact_or_keywords(df.columns, PI_COL_NAME, keywords=["principal", "investigator"])
    project_col = find_column_by_exact_or_keywords(df.columns, PROJECT_COL_NAME, keywords=["project", "number"])
    task_name_col = find_column_by_exact_or_keywords(df.columns, TASK_NAME_COL_NAME, keywords=["task", "name"])
    task_number_col = find_column_by_exact_or_keywords(df.columns, TASK_NUMBER_COL_NAME, keywords=["task", "number"])
    status_col = find_column_by_exact_or_keywords(df.columns, STATUS_COL_NAME, keywords=["task", "status"])
    owning_org_col = find_column_by_exact_or_keywords(df.columns, OWNING_ORG_COL_NAME, keywords=["owning", "org"])

    balance_col_candidates = [c for c in df.columns if str(c).startswith("Budget Balance")]
    if not balance_col_candidates:
        raise KeyError("Could not find a column whose name starts with 'Budget Balance'.")
    balance_col = balance_col_candidates[0]

    # Active rows only
    df_active = df[df[status_col] == "Active"].copy()
    if df_active.empty:
        raise ValueError("No rows found with Task Status == 'Active'.")

    # Sort
    df_active[project_col] = pd.to_numeric(df_active[project_col], errors="coerce")
    df_active[task_number_col] = pd.to_numeric(df_active[task_number_col], errors="coerce")
    df_active = df_active.sort_values(by=[pi_col, project_col, task_number_col], na_position="last")

    df_active = df_active.dropna(axis=1, how="all")

    # Select needed cols
    needed_cols = [pi_col]
    for col_name in BASE_OUTPUT_COLS:
        if col_name == "Project Number":
            actual = project_col
        elif col_name == "Task Name":
            actual = task_name_col
        elif col_name == "Task Number":
            actual = task_number_col
        elif col_name == OWNING_ORG_COL_NAME:
            actual = owning_org_col
        else:
            actual = col_name

        if actual in df_active.columns:
            needed_cols.append(actual)

    if balance_col not in needed_cols:
        needed_cols.append(balance_col)

    df_active = df_active[needed_cols].copy()

    # Rename budget balance column to canonical
    rename_map = {balance_col: "Budget Balance"}
    df_active = df_active.rename(columns=rename_map)

    # Preserve raw project number for Award merge
    df_active["_ProjectNumberRaw"] = df_active[project_col].astype(str).str.strip()

    # Combine Project Number + Task Number
    df_active[project_col] = (
        df_active[project_col].astype("Int64").astype(str).replace("<NA>", "").str.strip()
        + "-"
        + df_active[task_number_col].astype("Int64").astype(str).replace("<NA>", "").str.strip()
    ).str.strip("-")

    # Combine Project Name + Task Name
    if "Project Name" in df_active.columns and task_name_col in df_active.columns:
        df_active["Project Name"] = (
            df_active["Project Name"].astype(str).str.strip()
            + " – "
            + df_active[task_name_col].astype(str).str.strip()
        )

    # Drop task fields from output
    df_active = df_active.drop(columns=[task_name_col, task_number_col], errors="ignore")

    # Rename budget columns
    df_active = df_active.rename(
        columns={
            "Budget": "Allocated Budget",
            "Budget Balance": "Current Balance",
        }
    )

    # Read award info files from bytes
    award_dfs = []
    for b in award_bytes_list:
        df_aw = pd.read_excel(BytesIO(b))
        df_aw.columns = df_aw.columns.astype(str).str.strip()
        award_dfs.append(df_aw)
    award_df = pd.concat(award_dfs, ignore_index=True)

    if AWARD_INFO_PROJECT_COL not in award_df.columns:
        raise KeyError(f"Expected column '{AWARD_INFO_PROJECT_COL}' not found in Award Info file(s).")
    if AWARD_INFO_INDIRECT_COL not in award_df.columns:
        raise KeyError(f"Expected column '{AWARD_INFO_INDIRECT_COL}' not found in Award Info file(s).")

    df_active["_proj_key"] = df_active["_ProjectNumberRaw"].astype(str).str.strip()
    award_df["_proj_key"] = award_df[AWARD_INFO_PROJECT_COL].astype(str).str.strip()
    award_df = award_df.drop_duplicates(subset=["_proj_key"], keep="first")

    df_merged = df_active.merge(
        award_df[["_proj_key", AWARD_INFO_INDIRECT_COL]],
        on="_proj_key",
        how="left",
    ).drop(columns=["_proj_key", "_ProjectNumberRaw"])

    df_merged = df_merged.rename(columns={AWARD_INFO_INDIRECT_COL: "Indirect Rate"})

    if "Indirect Rate" not in df_merged.columns:
        raise ValueError("Missing 'Indirect Rate' after merge; cannot compute net values.")

    df_merged["Indirect Rate"] = pd.to_numeric(df_merged["Indirect Rate"], errors="coerce").fillna(0.0)

    df_merged["Allocated Budget"] = pd.to_numeric(df_merged["Allocated Budget"], errors="coerce")
    df_merged["Current Balance"] = pd.to_numeric(df_merged["Current Balance"], errors="coerce")

    denom = 1.0 + df_merged["Indirect Rate"]
    df_merged["Allocated Budget*"] = df_merged["Allocated Budget"] / denom
    df_merged["Current Balance*"] = df_merged["Current Balance"] / denom

    df_merged = df_merged.drop(columns=["Allocated Budget", "Current Balance"], errors="ignore")

    if pi_col in df_merged.columns and pi_col != "Principal Investigator":
        df_merged = df_merged.rename(columns={pi_col: "Principal Investigator"})
        pi_col = "Principal Investigator"

    df_merged["_PI_stripped"] = df_merged[pi_col].astype(str).apply(normalize_pi_name)
    df_merged[pi_col] = df_merged["_PI_stripped"]

    # Ensure owning org column has canonical name in df_merged
    if owning_org_col != OWNING_ORG_COL_NAME and owning_org_col in df_merged.columns:
        df_merged = df_merged.rename(columns={owning_org_col: OWNING_ORG_COL_NAME})

    if OWNING_ORG_COL_NAME not in df_merged.columns:
        raise ValueError(f"'{OWNING_ORG_COL_NAME}' not found; cannot build Org7 folders.")

    df_merged["_Org7"] = df_merged[OWNING_ORG_COL_NAME].apply(compute_org7)

    if prefix:
        df_merged["Date Pulled"] = prefix

    # Column order (keep owning org internally; it is dropped from output files)
    desired_order = []
    for name in [
        pi_col,
        "Project Manager",
        "Date Pulled",
        project_col,
        "Project Name",
        "Allocated Budget*",
        "Current Balance*",
        "Indirect Rate",
        OWNING_ORG_COL_NAME,
    ]:
        if name in df_merged.columns:
            desired_order.append(name)
    remaining = [c for c in df_merged.columns if c not in desired_order]
    df_merged = df_merged[desired_order + remaining]

    currency_cols = ["Allocated Budget*", "Current Balance*", "expenses"]

    per_pi_zip_bytes = build_per_pi_org7_zip(df_merged, base_name, currency_cols)

    summary_info = {
        "base_name": base_name,
        "date_pulled": prefix or "(not specified)",
        "num_rows": int(len(df_merged)),
        "num_pis": int(df_merged["_PI_stripped"].nunique(dropna=True)),
        "num_org7": int(df_merged["_Org7"].nunique(dropna=True)),
        "org7_examples": [str(x) for x in df_merged["_Org7"].dropna().unique()[:8]],
    }

    return base_name, per_pi_zip_bytes, summary_info


# -----------------------------
# UI helpers
# -----------------------------
def ucd_banner():
    st.markdown(
        """
        <div style="
            background: linear-gradient(90deg, #002855, #01223d);
            padding: 1.2rem 1.6rem;
            border-radius: 12px;
            margin-bottom: 1.5rem;
            display: flex;
            align-items: center;
            box-shadow: 0 6px 18px rgba(0, 0, 0, 0.18);
        ">
            <div style="font-size: 2.4rem; margin-right: 1rem;">🐄</div>
            <div style="flex: 1;">
                <div style="color: #FFBF00; font-weight: 700; letter-spacing: 0.16em;
                            text-transform: uppercase; font-size: 0.8rem;">
                    UC Davis • Yellow Cluster
                </div>
                <div style="color: #FFFFFF; font-size: 1.5rem; font-weight: 600; margin-top: 0.15rem;">
                    PI Budget Summary Generator
                </div>
                <div style="color: #d7e3f3; font-size: 0.9rem; margin-top: 0.25rem;">
                    Generate PI-specific budget summaries for sharing with faculty.
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def main():
    ucd_banner()
    st.markdown("_Written by David Railton Garrett_")

    with st.expander("Debug options", expanded=False):
        show_trace = st.checkbox("Show full error trace", value=False)

    st.markdown(
        """
        Upload your **Aggie Enterprise Database** and one or more **Award Info** workbooks.

        Output is a **ZIP** organized by owning org (first 7 digits from **Project Owning Organization**):
        `Org7/<Report> - <PI>.xlsx`
        """
    )

    master_file = st.file_uploader(
        "Upload Aggie Enterprise Database Excel file",
        type=["xlsx"],
        key="master_file",
    )

    award_files = st.file_uploader(
        "Upload one or more Award Info Excel file(s)",
        type=["xlsx"],
        key="award_files",
        accept_multiple_files=True,
    )

    date_pulled = st.text_input(
        "Date Pulled (optional, e.g., '2025-01-31' or 'Jan 2025')",
        value="",
    )

    if master_file and award_files:
        if st.button("Run processing", type="primary"):
            try:
                master_bytes = master_file.getvalue()
                award_bytes_list = [f.getvalue() for f in award_files]

                with st.spinner("Processing files..."):
                    base_name, per_pi_zip_bytes, summary = process_workbooks_cached(
                        master_bytes, award_bytes_list, date_pulled
                    )

                st.success("Processing complete!")

                st.markdown("### Summary")
                st.write(f"- Base filename: **{summary['base_name']}**")
                st.write(f"- Date Pulled: **{summary['date_pulled']}**")
                st.write(f"- Total rows (active, merged): **{summary['num_rows']}**")
                st.write(f"- Unique PIs: **{summary['num_pis']}**")
                st.write(f"- Org folders (Org7): **{summary['num_org7']}**")
                if summary["org7_examples"]:
                    st.write("Examples of Org7 folders:")
                    st.write(", ".join(summary["org7_examples"]))

                st.download_button(
                    "Download PI Budget Summaries (ZIP)",
                    data=per_pi_zip_bytes,
                    file_name=f"{base_name}_PI_files_by_OwningOrg.zip",
                    mime="application/zip",
                )

            except Exception as e:
                st.error(f"Error: {e}")
                if show_trace:
                    st.code(traceback.format_exc())
    else:
        st.info("Please upload both the Aggie Enterprise Database and at least one Award Info file to begin.")


if __name__ == "__main__":
    main()
