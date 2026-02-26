import re
from io import BytesIO
import pandas as pd
import streamlit as st


# ----------------------------
# Helpers
# ----------------------------
def norm_text(x) -> str:
    if pd.isna(x):
        return ""
    return re.sub(r"\s+", " ", str(x)).strip().lower()


def clean_series(s: pd.Series) -> pd.Series:
    """
    Convert NaN/None to empty string, strip whitespace, and prevent literal 'nan' text.
    """
    s = s.where(~pd.isna(s), "")
    s = s.astype(str).str.strip()
    return s.replace({"nan": "", "NaN": ""})


def kit_label(raw: str) -> str:
    v = norm_text(raw)
    if v == "":
        return ""
    if "instructor" in v:
        return "Instructor Kit"
    if "essential" in v:
        return "Essential Kit"
    return str(raw).strip().title()


def safe_sheet_name(name: str, used: set) -> str:
    """
    Excel sheet name rules: max 31 chars, no : \ / ? * [ ]
    Must be unique in workbook.
    """
    name = re.sub(r"[:\\/?*\[\]]", "-", name)
    name = re.sub(r"\s+", " ", name).strip()

    base = name[:31] if len(name) > 31 else name
    if base == "":
        base = "Sheet"

    if base not in used:
        used.add(base)
        return base

    i = 2
    while True:
        suffix = f"_{i}"
        cut = 31 - len(suffix)
        candidate = (base[:cut] if cut > 0 else base) + suffix
        if candidate not in used:
            used.add(candidate)
            return candidate
        i += 1


def guess_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """
    Return the first df column whose normalized name matches any candidate.
    """
    norm_map = {norm_text(c): c for c in df.columns}
    for cand in candidates:
        if cand in norm_map:
            return norm_map[cand]
    return None


def parse_lesson_tokens(value: object) -> list[str]:
    """
    Turn Lesson # cell into a list of lesson tokens.
    Examples:
      "1, 4, 7"  -> ["1","4","7"]
      "3,4"      -> ["3","4"]
      ""/NaN     -> [""]
      "All"      -> ["__ALL__"]
    """
    if pd.isna(value):
        return [""]
    s = str(value).strip()
    if s == "" or s.lower() in ["nan", "none"]:
        return [""]
    if s.strip().lower() == "all":
        return ["__ALL__"]

    # split by comma
    parts = [p.strip() for p in s.split(",")]
    tokens = [p for p in parts if p != ""]
    return tokens if tokens else [""]


# ----------------------------
# Core transformation
# ----------------------------
def build_output_excel(
    df: pd.DataFrame,
    col_class: str,
    col_lesson: str,
    col_lesson_num: str,
    col_item: str,
    col_per_section: str,
    col_size: str | None,
    col_notes: str | None,
    col_kit_src: str | None,
    col_class_type: str | None,
    include_kit_column: bool,
    put_kit_under_lesson_num: bool,
) -> tuple[bytes, int]:

    data = df.copy()

    # Clean core fields (prevents 'nan' showing up anywhere)
    data["_class_clean"] = clean_series(data[col_class])
    data["_lesson_clean"] = clean_series(data[col_lesson])
    data["_item_clean"] = clean_series(data[col_item])

    # Lesson # raw string (keep blanks blank)
    lesson_num_clean = data[col_lesson_num].where(~pd.isna(data[col_lesson_num]), "")
    data["_lessonnum_raw"] = lesson_num_clean.astype(str).str.strip().replace({"nan": "", "NaN": ""})

    # Per section total: keep blanks blank (don't turn into 'nan')
    data["_per_section"] = data[col_per_section].where(~pd.isna(data[col_per_section]), "")

    # Optional fields
    data["_size_clean"] = clean_series(data[col_size]) if col_size else ""
    data["_notes_clean"] = clean_series(data[col_notes]) if col_notes else ""

    # Kit derived from a source column (if provided)
    data["_kit"] = data[col_kit_src].apply(kit_label) if col_kit_src else ""

    # Optional class type
    data["_class_type_clean"] = clean_series(data[col_class_type]) if col_class_type else ""

    # ----------------------------
    # EXPAND LIST LESSON NUMBERS
    # If Lesson # is "1, 4, 7" then duplicate the row into 3 rows with lesson numbers 1 / 4 / 7.
    # If Lesson # is "All", duplicate into every lesson number that exists for the same Class Type + Class + Lesson Name.
    # ----------------------------
    base_group = ["_class_type_clean", "_class_clean", "_lesson_clean"]

    # First pass: explode comma lists (and keep "__ALL__" marker)
    data["_lesson_tokens"] = data["_lessonnum_raw"].apply(parse_lesson_tokens)
    data = data.explode("_lesson_tokens", ignore_index=True)
    data["_lessonnum_clean"] = data["_lesson_tokens"].astype(str).str.strip().replace({"nan": "", "NaN": ""})

    # Handle "All"
    all_mask = data["_lessonnum_clean"] == "__ALL__"
    if all_mask.any():
        # lessons available in each group (exclude blanks and "__ALL__")
        lessons_by_group = (
            data.loc[~all_mask, base_group + ["_lessonnum_clean"]]
            .copy()
        )
        lessons_by_group = lessons_by_group[lessons_by_group["_lessonnum_clean"] != ""]
        lessons_by_group = lessons_by_group.drop_duplicates()
        lesson_sets = (
            lessons_by_group.groupby(base_group)["_lessonnum_clean"]
            .apply(lambda s: sorted(set([str(x).strip() for x in s if str(x).strip() != ""])))
            .to_dict()
        )

        all_rows = data.loc[all_mask].copy()
        non_all = data.loc[~all_mask].copy()

        expanded_all = []
        for _, r in all_rows.iterrows():
            key = (r["_class_type_clean"], r["_class_clean"], r["_lesson_clean"])
            lessons = lesson_sets.get(key, [])
            if lessons:
                for ln in lessons:
                    rr = r.copy()
                    rr["_lessonnum_clean"] = ln
                    expanded_all.append(rr)
            else:
                # no known lessons in this group; keep it blank (still not dropped)
                rr = r.copy()
                rr["_lessonnum_clean"] = ""
                expanded_all.append(rr)

        data = pd.concat([non_all, pd.DataFrame(expanded_all)], ignore_index=True)

    # ----------------------------
    # DROP ONLY 100% BLANK ROWS (your rule)
    # ----------------------------
    size_blank = (data["_size_clean"] == "") if isinstance(data["_size_clean"], pd.Series) else True
    notes_blank = (data["_notes_clean"] == "") if isinstance(data["_notes_clean"], pd.Series) else True

    per_blank = data["_per_section"].copy()
    per_blank = per_blank.where(~pd.isna(per_blank), "")
    per_blank = per_blank.astype(str).str.strip().replace({"nan": "", "NaN": ""})
    per_is_blank = (per_blank == "")

    fully_blank = (
        (data["_class_clean"] == "") &
        (data["_lesson_clean"] == "") &
        (data["_lessonnum_clean"] == "") &
        (data["_item_clean"] == "") &
        per_is_blank &
        size_blank &
        notes_blank &
        (data["_kit"] == "") &
        (data["_class_type_clean"] == "")
    )
    data = data.loc[~fully_blank].copy()

    # ----------------------------
    # Grouping logic (Option 2)
    # - Group by Class Type + Class Name + Lesson Name
    # - Use first non-blank Lesson # as the group's "rep" lesson number for grouping only
    # - Keep the displayed Lesson # exactly as-is (blank stays blank)
    # ----------------------------
    def pick_rep(series: pd.Series) -> str:
        for v in series:
            v = "" if pd.isna(v) else str(v).strip()
            if v:
                return v
        return ""

    rep_map = (
        data.groupby(base_group, dropna=False)["_lessonnum_clean"]
        .apply(pick_rep)
        .rename("_lessonnum_rep")
        .reset_index()
    )
    data = data.merge(rep_map, how="left", on=base_group)

    data["_lessonnum_group"] = data["_lessonnum_clean"]
    missing_mask = data["_lessonnum_group"] == ""
    data.loc[missing_mask, "_lessonnum_group"] = data.loc[missing_mask, "_lessonnum_rep"]

    # Display Lesson # (optionally with kit under it)
    display_lesson = []
    for ln, kit in zip(data["_lessonnum_clean"], data["_kit"]):
        ln = "" if pd.isna(ln) else str(ln).strip()
        kit = "" if pd.isna(kit) else str(kit).strip()
        if not put_kit_under_lesson_num:
            display_lesson.append(ln)
        else:
            if ln and kit:
                display_lesson.append(f"{ln}\n{kit}")
            elif ln:
                display_lesson.append(ln)
            elif kit:
                display_lesson.append(kit)
            else:
                display_lesson.append("")
    data["Lesson #"] = display_lesson

    # Output table
    out = pd.DataFrame({
        "Packed": "",
        "Received": "",
        "Class Type": data["_class_type_clean"],
        "Class Name": data["_class_clean"],
        "Lesson Name": data["_lesson_clean"],
        "Lesson #": data["Lesson #"],
        "Item Description": data["_item_clean"],
        "Per Section total": data["_per_section"],
        "Item Size": data["_size_clean"] if isinstance(data["_size_clean"], pd.Series) else "",
        "Notes": data["_notes_clean"] if isinstance(data["_notes_clean"], pd.Series) else "",
    })

    if include_kit_column:
        out["Kit"] = data["_kit"]

    # Helper for grouping into tabs
    out["_lessonnum_group"] = data["_lessonnum_group"]

    # Rows that still have no Class/Lesson/Item info (but weren't fully blank) go to Unassigned
    is_unassigned = (
        (out["Class Name"] == "") &
        (out["Lesson Name"] == "") &
        (out["_lessonnum_group"] == "") &
        (out["Item Description"] == "") &
        (out["Per Section total"].astype(str).str.strip().replace({"nan": "", "NaN": ""}) == "")
    )

    output = BytesIO()
    used = set()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        normal = out[~is_unassigned].copy()
        normal_sorted = normal.sort_values(by=["Class Type", "Class Name", "Lesson Name", "_lessonnum_group", "Item Description"])

        group_keys = ["Class Type", "Class Name", "Lesson Name", "_lessonnum_group"]

        for (ctype, cname, lname, lnum_group), g in normal_sorted.groupby(group_keys, dropna=False):
            cname_s = "" if pd.isna(cname) else str(cname).strip()
            lname_s = "" if pd.isna(lname) else str(lname).strip()
            lnum_s = "" if pd.isna(lnum_group) else str(lnum_group).strip()

            parts = []
            if cname_s:
                parts.append(cname_s)
            if lname_s:
                parts.append(lname_s)
            if lnum_s:
                parts.append(f"Lesson {lnum_s}")

            sheet_title = " - ".join(parts) if parts else "Unassigned"
            sheet = safe_sheet_name(sheet_title, used)

            g = g.drop(columns=["_lessonnum_group"])
            g.to_excel(writer, index=False, sheet_name=sheet)

            ws = writer.sheets[sheet]
            ws.freeze_panes = "A2"

            header = [cell.value for cell in ws[1]]
            if "Lesson #" in header:
                col_idx = header.index("Lesson #") + 1
                for r in range(2, ws.max_row + 1):
                    cell = ws.cell(row=r, column=col_idx)
                    cell.alignment = cell.alignment.copy(wrapText=True)

            widths = {
                "Packed": 10, "Received": 10, "Class Type": 18, "Class Name": 22, "Lesson Name": 28,
                "Lesson #": 16, "Item Description": 32, "Per Section total": 16, "Item Size": 14, "Notes": 18, "Kit": 14,
            }
            for i, col_name in enumerate(header, start=1):
                ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = widths.get(col_name, 16)

        unassigned = out[is_unassigned].copy()
        if len(unassigned) > 0:
            sheet = safe_sheet_name("Unassigned Rows", used)
            unassigned = unassigned.drop(columns=["_lessonnum_group"])
            unassigned.to_excel(writer, index=False, sheet_name=sheet)
            ws = writer.sheets[sheet]
            ws.freeze_panes = "A2"

    return output.getvalue(), len(used)


# ----------------------------
# Streamlit UI (Master-only)
# ----------------------------
st.set_page_config(page_title="Packing List Generator", layout="centered")
st.title("Packing List Generator")
st.write("Upload your bulk order Excel file and download one Excel file with **one tab per lesson**.")

uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

if uploaded is None:
    st.info("Upload an Excel file to begin.")
    st.stop()

xls = pd.ExcelFile(uploaded)

# Always use Master
if "Master" not in xls.sheet_names:
    st.error("This file must contain a sheet named 'Master'.")
    st.stop()

df = pd.read_excel(xls, sheet_name="Master")
st.caption(f"Loaded **Master** — {len(df):,} rows, {len(df.columns)} columns")

all_cols = ["(None)"] + list(df.columns)

guess = {
    "Class Name": guess_column(df, ["class name", "class"]),
    "Lesson Name": guess_column(df, ["lesson name"]),
    "Lesson #": guess_column(df, ["lesson #", "lesson number", "lesson num", "lesson number "]),
    "Item Description": guess_column(df, ["item description", "item"]),
    "Per Section total": guess_column(df, ["per section total", "per section"]),
    "Item Size": guess_column(df, ["item size", "size"]),
    "Notes": guess_column(df, ["notes", "note"]),
    "Essential Items": guess_column(df, ["essential items", "essential"]),
    "Class Type": guess_column(df, ["class type", "class type (name)"]),
}

with st.expander("Column Matching (only use if needed)", expanded=False):
    st.write("If anything is wrong, choose the correct column from each dropdown.")

    def pick(label, guessed, required=False):
        idx = all_cols.index(guessed) if guessed in all_cols else 0
        return st.selectbox(label + (" *" if required else ""), all_cols, index=idx)

    col_class = pick("Class Name", guess["Class Name"], required=True)
    col_lesson = pick("Lesson Name", guess["Lesson Name"], required=True)
    col_lesson_num = pick("Lesson # / Lesson Number", guess["Lesson #"], required=True)
    col_item = pick("Item Description", guess["Item Description"], required=True)
    col_per_section = pick("Per Section total", guess["Per Section total"], required=True)

    col_size = pick("Item Size (optional)", guess["Item Size"])
    col_notes = pick("Notes (optional)", guess["Notes"])
    col_kit_src = pick("Essential Items (optional)", guess["Essential Items"])
    col_class_type = pick("Class Type (optional)", guess["Class Type"])

if "col_class" not in locals():
    col_class = guess["Class Name"] or "(None)"
    col_lesson = guess["Lesson Name"] or "(None)"
    col_lesson_num = guess["Lesson #"] or "(None)"
    col_item = guess["Item Description"] or "(None)"
    col_per_section = guess["Per Section total"] or "(None)"
    col_size = guess["Item Size"] or "(None)"
    col_notes = guess["Notes"] or "(None)"
    col_kit_src = guess["Essential Items"] or "(None)"
    col_class_type = guess["Class Type"] or "(None)"

required = {
    "Class Name": col_class,
    "Lesson Name": col_lesson,
    "Lesson #": col_lesson_num,
    "Item Description": col_item,
    "Per Section total": col_per_section,
}
missing = [k for k, v in required.items() if v == "(None)"]
if missing:
    st.error("Missing required columns: " + ", ".join(missing) + ". Open **Column Matching** and select them.")
    st.stop()

include_kit_column = st.checkbox("Also include a separate 'Kit' column (optional)", value=False)
put_kit_under_lesson_num = st.checkbox("Put kit label under 'Lesson #' (recommended)", value=True)

st.divider()

if st.checkbox("Preview first 20 rows", value=False):
    st.dataframe(df.head(20), use_container_width=True)

if st.button("Generate packing list Excel", type="primary"):
    xlsx_bytes, tabs = build_output_excel(
        df=df,
        col_class=col_class,
        col_lesson=col_lesson,
        col_lesson_num=col_lesson_num,
        col_item=col_item,
        col_per_section=col_per_section,
        col_size=None if col_size == "(None)" else col_size,
        col_notes=None if col_notes == "(None)" else col_notes,
        col_kit_src=None if col_kit_src == "(None)" else col_kit_src,
        col_class_type=None if col_class_type == "(None)" else col_class_type,
        include_kit_column=include_kit_column,
        put_kit_under_lesson_num=put_kit_under_lesson_num,
    )

    st.success(f"Done. Created {tabs} tabs.")
    st.download_button(
        "Download packing lists (Excel)",
        data=xlsx_bytes,
        file_name="packing_lists_by_lesson.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )