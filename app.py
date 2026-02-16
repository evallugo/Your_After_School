import io
import re
import pandas as pd
import streamlit as st

import os
APP_PASSWORD = os.getenv("APP_PASSWORD")
if APP_PASSWORD:
    pw = st.text_input("Password", type="password")
    if pw != APP_PASSWORD:
        st.stop()

st.set_page_config(page_title="Packing List Generator", layout="centered")
st.title("Packing List Generator")
st.caption("Upload your bulk purchasing Excel file, confirm the columns, then download packing lists.")

uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

# ---------- helpers ----------
def safe_sheet_name(name: str) -> str:
    name = re.sub(r"[:\\/?*\[\]]", "-", str(name)).strip()
    return (name if name else "Sheet")[:31]

def normalize(s: str) -> str:
    return re.sub(r"\s+", " ", str(s)).strip().lower()

def guess_col(columns, candidates):
    """
    Return best-guess column based on normalized substring matches.
    candidates: list of phrases to look for (normalized)
    """
    cols_norm = {c: normalize(c) for c in columns}
    for phrase in candidates:
        for c, cn in cols_norm.items():
            if phrase in cn:
                return c
    return None

def make_output_excel(df, class_col, lesson_col, item_col, qty_col, size_col=None, uom_col=None):
    df = df.dropna(subset=[class_col, lesson_col, item_col]).copy()
    df[qty_col] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0)

    group_cols = [class_col, lesson_col, item_col]
    agg = {qty_col: "sum"}
    if size_col and size_col in df.columns:
        agg[size_col] = "first"
    if uom_col and uom_col in df.columns:
        agg[uom_col] = "first"

    df2 = df.groupby(group_cols, as_index=False).agg(agg)

    output = io.BytesIO()
    index_rows = []

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # INDEX tab
        for (class_name, lesson_name), g in df2.groupby([class_col, lesson_col]):
            sheet = safe_sheet_name(f"{class_name} - {lesson_name}")

            out_cols = [item_col, qty_col]
            if size_col and size_col in g.columns: out_cols.append(size_col)
            if uom_col and uom_col in g.columns: out_cols.append(uom_col)

            out = g[out_cols].sort_values(by=item_col).rename(columns={
                item_col: "Item",
                qty_col: "Quantity",
                size_col: "Size" if size_col else "Size",
                uom_col: "Unit/Notes" if uom_col else "Unit/Notes",
            })

            out.to_excel(writer, sheet_name=sheet, index=False)
            index_rows.append({"Class": class_name, "Lesson": lesson_name, "Sheet": sheet, "Items": len(out)})

        pd.DataFrame(index_rows).sort_values(["Class", "Lesson"]).to_excel(writer, sheet_name="INDEX", index=False)

    return output.getvalue()

# ---------- app ----------
if uploaded:
    file_bytes = uploaded.getvalue()

    try:
        xl = pd.ExcelFile(io.BytesIO(file_bytes))
        sheet_names = xl.sheet_names
    except Exception as e:
        st.error(f"Could not read this as an Excel file. Details: {e}")
        st.stop()

    # heuristic: prefer sheets with "master" / "purchase" / "bulk" in the name
    preferred = None
    for key in ["master", "purch", "bulk", "order", "list"]:
        for s in sheet_names:
            if key in normalize(s):
                preferred = s
                break
        if preferred:
            break

    sheet = st.selectbox(
        "Which sheet should we use?",
        sheet_names,
        index=sheet_names.index(preferred) if preferred in sheet_names else 0
    )

    try:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet)
    except Exception as e:
        st.error(f"Could not read sheet '{sheet}'. Details: {e}")
        st.stop()

    # drop completely empty columns (common in exports)
    df = df.dropna(axis=1, how="all")

    if df.empty or len(df.columns) == 0:
        st.error("This sheet looks empty or has no usable columns.")
        st.stop()

    st.subheader("Step 1: Confirm column mapping")
    st.write("Pick which columns correspond to Class, Lesson, Item, and Quantity. Optional fields help add detail.")

    cols = list(df.columns)

    # smart guesses (works even if headers vary slightly)
    default_class = guess_col(cols, ["class", "course", "program"])
    default_lesson = guess_col(cols, ["lesson", "module", "unit", "activity"])
    default_item = guess_col(cols, ["item description", "item", "product", "material", "supply"])
    # quantity can vary a lot: "per section total", "needed", "qty", "quantity"
    default_qty = guess_col(cols, ["per section total", "needed", "quantity", "qty", "total"])

    default_size = guess_col(cols, ["size"])
    default_uom = guess_col(cols, ["unit of measure", "uom", "units", "unit"])

    col1, col2 = st.columns(2)
    with col1:
        class_col = st.selectbox("Class column (required)", cols, index=cols.index(default_class) if default_class in cols else 0)
        lesson_col = st.selectbox("Lesson column (required)", cols, index=cols.index(default_lesson) if default_lesson in cols else 0)
        item_col = st.selectbox("Item column (required)", cols, index=cols.index(default_item) if default_item in cols else 0)
        qty_col = st.selectbox("Quantity column (required)", cols, index=cols.index(default_qty) if default_qty in cols else 0)
    with col2:
        size_col = st.selectbox("Size column (optional)", ["(none)"] + cols,
                                index=(["(none)"] + cols).index(default_size) if default_size in cols else 0)
        uom_col = st.selectbox("Unit/Notes column (optional)", ["(none)"] + cols,
                               index=(["(none)"] + cols).index(default_uom) if default_uom in cols else 0)

    # validate distinct required selections
    required = [class_col, lesson_col, item_col, qty_col]
    if len(set(required)) < 4:
        st.error("Your required column selections must be four different columns.")
        st.stop()

    # optional → None
    size_col = None if size_col == "(none)" else size_col
    uom_col = None if uom_col == "(none)" else uom_col

    st.subheader("Step 2: Preview")
    preview = df[[c for c in [class_col, lesson_col, item_col, qty_col, size_col, uom_col] if c is not None]].head(25)
    st.dataframe(preview, use_container_width=True)

    st.subheader("Step 3: Generate")
    if st.button("Generate Packing Lists"):
        try:
            result = make_output_excel(
                df=df,
                class_col=class_col,
                lesson_col=lesson_col,
                item_col=item_col,
                qty_col=qty_col,
                size_col=size_col,
                uom_col=uom_col
            )
            st.success("Done! Your packing lists are ready.")
            st.download_button(
                "Download Packing_Lists.xlsx",
                data=result,
                file_name="Packing_Lists.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"Something went wrong while generating output: {e}")
