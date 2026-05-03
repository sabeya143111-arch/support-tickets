import io
from typing import List, Dict

import pdfplumber
import pandas as pd
import streamlit as st


# -------------------- APP CONFIG & TITLE --------------------

st.set_page_config(page_title="Supplier spec reader", page_icon="📏", layout="wide")
st.title("📏 Supplier spec reader – size chart extractor")

st.write(
    """
    This internal tool reads a supplier's specification PDF and extracts the size chart
    for each model (style). The output is shown in clear English so you can **manually
    compare** with the MF Template Matrix.

    **Steps:**
    1. Upload the supplier PO spec PDF.
    2. Select a model from the list on the left.
    3. Use the English size chart to check against Template Matrix Detail.
    """
)


# -------------------- CONFIG: Chinese -> English names --------------------

CN_TO_EN = {
    "衣长": "Body length",
    "肩宽": "Shoulder width",
    "胸围": "Chest",
    "袖长": "Sleeve length",
    "袖口宽": "Sleeve opening width",
    "下摆": "Bottom hem width",
}


# -------------------- Simple data classes --------------------

class Measurement:
    def __init__(self, name: str, values: Dict[str, float]):
        self.name = name
        self.values = values  # size -> value


class SupplierStyle:
    def __init__(self, style_code: str, product_name: str, color: str,
                 sizes: List[str], measurements: List[Measurement]):
        self.style_code = style_code
        self.product_name = product_name
        self.color = color
        self.sizes = sizes
        self.measurements = measurements


# -------------------- PDF parsing helpers --------------------

def extract_numbers(part: str) -> List[float]:
    """Extract numeric values from a string."""
    nums = []
    part = part.replace("：", ":").replace(",", " ")
    for tok in part.split():
        try:
            nums.append(float(tok))
        except ValueError:
            continue
    return nums


def parse_supplier_pdf_multi(file_bytes: bytes) -> List[SupplierStyle]:
    """
    Parse supplier spec PDF and return a list of SupplierStyle objects.
    This is tuned for your current PDF format (TFxxxx-x-S, S M L XL, 衣长/肩宽/胸围/etc.).
    If supplier changes template, tweak this function.
    """
    styles: List[SupplierStyle] = []

    # 1) Extract all text from all pages
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        text_all = ""
        for page in pdf.pages:
            text_all += (page.extract_text() or "") + "\n"

    # 2) Split into non-empty lines
    lines = [ln.strip() for ln in text_all.splitlines() if ln.strip()]

    current_style = None
    buffer_measurements: Dict[str, Measurement] = {}
    sizes: List[str] = []

    def flush_style():
        nonlocal current_style, buffer_measurements, sizes, styles
        if current_style and buffer_measurements and sizes:
            styles.append(
                SupplierStyle(
                    style_code=current_style["code"],
                    product_name=current_style.get("product", ""),
                    color=current_style.get("color", ""),
                    sizes=sizes.copy(),
                    measurements=list(buffer_measurements.values()),
                )
            )
        current_style = None
        buffer_measurements = {}
        sizes = []

    i = 0
    while i < len(lines):
        line = lines[i]

        # Detect new style, e.g. "TF0019-1-S"
        if line.startswith("TF") and "-S" in line:
            flush_style()
            code = line.split("-S")[0].strip()
            current_style = {"code": code, "product": "", "color": ""}

            # Next line normally has "MEN T-SHIRT <color> ..."
            if i + 1 < len(lines):
                info = lines[i + 1]
                if "MEN T-SHIRT" in info:
                    current_style["product"] = "MEN T-SHIRT"
                    after = info.split("MEN T-SHIRT", 1)[1].strip()
                    if after:
                        current_style["color"] = after.split()[0]
                else:
                    current_style["product"] = info
            i += 1
            continue

        # Detect sizes line: allow flexible spaces
        normalized = " ".join(line.split())
        if normalized.startswith("S M L XL"):
            sizes = ["S", "M", "L", "XL"]
            i += 1
            continue

        # Detect measurement rows using Chinese names
        for cn, en in CN_TO_EN.items():
            if cn in line:
                part = line.split(cn, 1)[-1]
                nums = extract_numbers(part)
                if sizes and len(nums) >= len(sizes):
                    vals = {sizes[idx]: nums[idx] for idx in range(len(sizes))}
                    buffer_measurements[en] = Measurement(name=en, values=vals)
                break

        i += 1

    flush_style()
    return styles


# -------------------- Streamlit UI --------------------

st.header("1️⃣ Upload supplier spec PDF")

pdf_file = st.file_uploader("Select PDF file", type=["pdf"])

if not pdf_file:
    st.info("Upload the supplier specification PDF to start.")
    st.stop()

# Parse PDF
try:
    styles = parse_supplier_pdf_multi(pdf_file.read())
except Exception as e:
    st.error(f"Error while reading PDF: {e}")
    st.stop()

if not styles:
    st.warning("No models / measurements found in this PDF. Check format or parser.")
    st.stop()

st.success(f"Found **{len(styles)}** model(s) with size charts in this PDF.")

# Sidebar: list of models
st.sidebar.header("Models in this PDF")
style_labels = [f"{s.style_code} – {s.product_name} – {s.color}" for s in styles]
selected_idx = st.sidebar.selectbox(
    "Select model",
    options=range(len(styles)),
    format_func=lambda i: style_labels[i],
)

style = styles[selected_idx]

# Main layout: top info + text + table
st.header("2️⃣ Model details")

col1, col2, col3 = st.columns(3)
with col1:
    st.metric("Model code", style.style_code)
with col2:
    st.metric("Product", style.product_name or "N/A")
with col3:
    st.metric("Color", style.color or "N/A")

st.markdown(f"**Sizes in spec:** {', '.join(style.sizes)}")

# Build English text block
lines = [
    f"Model: {style.style_code}",
    f"Product: {style.product_name}",
    f"Color: {style.color}",
    f"Sizes: {', '.join(style.sizes)}",
    "",
]
for m in style.measurements:
    lines.append(f"{m.name}:")
    for size in style.sizes:
        if size in m.values:
            lines.append(f"  {size} = {m.values[size]} cm")
    lines.append("")
full_text = "\n".join(lines)

st.subheader("Readable size chart (copy for MF check)")
st.text_area("Model details", full_text, height=320)

# Table view
st.subheader("Size chart table")
rows = []
for m in style.measurements:
    row = {"Measurement": m.name}
    for size in style.sizes:
        row[size] = m.values.get(size)
    rows.append(row)
df = pd.DataFrame(rows)
st.dataframe(df, use_container_width=True)

# Simple stats
st.header("3️⃣ Quick statistics")
col_a, col_b = st.columns(2)
with col_a:
    st.write(f"- Measurements for this model: **{len(style.measurements)}**")
    st.write(f"- Sizes per row: **{len(style.sizes)}**")
with col_b:
    st.write("- Use this chart to manually compare with Template Matrix Detail.")
    st.write("- If supplier PDF format changes, parser function can be updated.")
