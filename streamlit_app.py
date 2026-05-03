"""
Supplier PO Size Chart Extractor
Reads Chinese supplier Purchase Order PDFs (BOY LONG PANTS / MEN styles)
and displays clean English size charts for manual comparison.

Supports two PDF structures found in real supplier POs:
  1. Main PO table  – PMY/TF style rows with size columns in a table
  2. Appended size sheet – 大童尺寸表 / 大童长裤尺寸表 tables on a separate page

No external LLM calls. Uses only: streamlit, pdfplumber, pandas.
"""

import io
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

import pandas as pd
import pdfplumber
import streamlit as st

# ─────────────────────────────────────────────────────────────────────────────
# Data classes
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class Measurement:
    name: str               # English label
    values: Dict[str, str]  # size_label -> value string


@dataclass
class SupplierStyle:
    style_code: str
    product_name: str
    color: str
    fabric: str
    sizes: List[str]
    measurements: List[Measurement]
    raw_qty: str = ""
    raw_price: str = ""


# ─────────────────────────────────────────────────────────────────────────────
# Chinese → English measurement mapping
# ─────────────────────────────────────────────────────────────────────────────

MEAS_MAP = {
    "二分之一腰围": "1/2 Waist",
    "二分之一脚口": "1/2 Leg Opening",
    "坐围（直裆1/3处弧度）": "Seat (arc at 1/3 crotch)",
    "坐围": "Seat circumference",
    "裤长": "Pant length",
    "衣长": "Body length",
    "肩宽": "Shoulder width",
    "肩": "Shoulder width",
    "胸围": "Chest",
    "胸": "Chest",
    "袖长": "Sleeve length",
    "袖口宽": "Sleeve opening width",
    "袖口": "Sleeve opening width",
    "下摆": "Bottom hem width",
}

def chinese_to_english(label: str) -> str:
    label = label.strip().rstrip("：:").strip()
    for zh, en in MEAS_MAP.items():
        if zh in label:
            return en
    return label  # fallback: return as-is


def extract_numbers(text: str) -> List[str]:
    """Extract numeric tokens (int or float) from a string."""
    return re.findall(r"\d+(?:\.\d+)?", text)


# ─────────────────────────────────────────────────────────────────────────────
# Parser for the appended size-chart pages (附页合同)
# These pages contain structured tables like:
#   码数  | 24  | 25  | 26  | 27  | 28  | 单位
#   二分之一腰围 | 29 | 30.5 | 32 | 33.5 | 35 | CM
# ─────────────────────────────────────────────────────────────────────────────

def parse_size_sheet_tables(pages) -> List[SupplierStyle]:
    """
    Extract size charts from the appended 附页合同 page(s).
    Returns one SupplierStyle per distinct size table found.
    """
    styles: List[SupplierStyle] = []

    for page in pages:
        tables = page.extract_tables()
        if not tables:
            continue

        for table in tables:
            if not table or len(table) < 2:
                continue

            # First row should be the header: 码数 | size1 | size2 | ... | 单位
            header_row = table[0]
            if not header_row:
                continue

            # Normalize header cells
            header = [str(c).strip() if c else "" for c in header_row]

            # Detect if this looks like a size table
            has_madas = any("码数" in h for h in header)
            if not has_madas:
                continue

            # Size labels are between 码数 and 单位
            size_start = next((i for i, h in enumerate(header) if "码数" in h), None)
            size_end = next((i for i, h in enumerate(header) if "单位" in h), None)

            if size_start is None:
                continue

            if size_end is None:
                size_end = len(header)

            sizes = [h for h in header[size_start + 1:size_end] if h]

            if not sizes:
                continue

            # Collect measurements from remaining rows
            measurements: List[Measurement] = []
            for row in table[1:]:
                if not row or not row[0]:
                    continue
                label_raw = str(row[0]).strip()
                if not label_raw or label_raw == "单位":
                    continue

                en_label = chinese_to_english(label_raw)
                vals_raw = [str(row[i]).strip() if i < len(row) and row[i] else "" 
                            for i in range(size_start + 1, size_start + 1 + len(sizes))]

                values = {sizes[j]: vals_raw[j] if j < len(vals_raw) else "" 
                          for j in range(len(sizes))}
                measurements.append(Measurement(name=en_label, values=values))

            if measurements:
                # Infer product group from page text
                page_text = page.extract_text() or ""
                if "PN007" in page_text or "PN008" in page_text or "PN009" in page_text or "PN010" in page_text:
                    group = "PN007-010 Size Chart"
                elif "PN002" in page_text or "PN003" in page_text or "PN004" in page_text:
                    group = "PN002-006 Size Chart"
                else:
                    group = f"Size Chart (sizes: {', '.join(sizes)})"

                # Detect fabric from page text
                fabric_match = re.search(r"面料[：:]\s*(\S+)", page_text)
                fabric = fabric_match.group(1) if fabric_match else ""

                styles.append(SupplierStyle(
                    style_code=group,
                    product_name="BOY LONG PANTS",
                    color="(multiple)",
                    fabric=fabric,
                    sizes=sizes,
                    measurements=measurements,
                ))

    return styles


# ─────────────────────────────────────────────────────────────────────────────
# Parser for the main PO table pages
# Extracts individual PMY/TF style rows from pdfplumber tables
# ─────────────────────────────────────────────────────────────────────────────

def clean_cell(v) -> str:
    return str(v).strip() if v else ""


def parse_po_table(pages) -> List[SupplierStyle]:
    """
    Parse main PO table rows.
    Each row group looks like:
      STYLE NO | IMAGE | DESCRIPTION | COLOR | QTY | PRICE | AMOUNT | FABRIC | SIZE | ...
    Style rows share the same DESCRIPTION / FABRIC / SIZE cell (merged cells).
    """
    styles: List[SupplierStyle] = []

    # Collect all table rows across pages
    all_rows: List[List[str]] = []
    for page in pages:
        tables = page.extract_tables()
        for table in tables:
            if not table:
                continue
            # Skip tables that don't look like the PO table
            # (size sheet tables handled separately)
            flat_header = " ".join(str(c) for c in (table[0] or []))
            if "码数" in flat_header:
                continue
            for row in table:
                cleaned = [clean_cell(c) for c in row]
                all_rows.append(cleaned)

    # Find header row
    header_idx = None
    for i, row in enumerate(all_rows):
        joined = " ".join(row)
        if "STYLE" in joined and "DESCRIPTION" in joined:
            header_idx = i
            break

    if header_idx is None:
        return []

    header = all_rows[header_idx]
    # Map column names to indices
    def col_idx(keywords):
        for j, h in enumerate(header):
            if any(k.upper() in h.upper() for k in keywords):
                return j
        return None

    col_style   = col_idx(["STYLE"])
    col_desc    = col_idx(["DESCRIPTION"])
    col_color   = col_idx(["COLOR"])
    col_qty     = col_idx(["QTY", "T'QTY"])
    col_price   = col_idx(["PRICE"])
    col_fabric  = col_idx(["FABRIC"])
    col_size    = col_idx(["SIZE"])

    if col_style is None or col_color is None:
        return []

    # Variables carried forward (for merged-cell rows)
    current_desc = ""
    current_fabric = ""
    current_sizes_raw = ""

    style_pattern = re.compile(r"^(PMY|TF|SW)\d+", re.IGNORECASE)

    for row in all_rows[header_idx + 1:]:
        if not any(row):
            continue

        style_val = row[col_style] if col_style < len(row) else ""
        color_val = row[col_color] if col_color < len(row) else ""
        qty_val   = row[col_qty]   if col_qty   is not None and col_qty < len(row) else ""
        price_val = row[col_price] if col_price is not None and col_price < len(row) else ""

        # Update carried-forward merged cells
        if col_desc is not None and col_desc < len(row) and row[col_desc]:
            current_desc = row[col_desc]
        if col_fabric is not None and col_fabric < len(row) and row[col_fabric]:
            current_fabric = row[col_fabric]
        if col_size is not None and col_size < len(row) and row[col_size]:
            current_sizes_raw = row[col_size]

        # Only process rows that have a real style code
        if not style_val or not style_pattern.match(style_val):
            continue
        if not color_val:
            continue

        # Parse sizes from the SIZE cell
        # Typical: "24（7-8） 25（9-10） 26（11-12）\n27（13-14） 28（15-16）\n1 1 1 1 1=5PCS/DOZEN"
        sizes = parse_sizes_from_cell(current_sizes_raw)

        styles.append(SupplierStyle(
            style_code=style_val,
            product_name=current_desc or "BOY LONG PANTS",
            color=color_val,
            fabric=current_fabric,
            sizes=sizes,
            measurements=[],  # measurement data is in the size sheet
            raw_qty=qty_val,
            raw_price=price_val,
        ))

    return styles


def parse_sizes_from_cell(raw: str) -> List[str]:
    """
    Turn a SIZE cell like:
    "24（7-8） 25（9-10） 26（11-12）\n27（13-14） 28（15-16）\n1 1 1 1 1=5PCS/DOZEN"
    into ["24(7-8)", "25(9-10)", "26(11-12)", "27(13-14)", "28(15-16)"]
    OR "7-8 9-10 11-12 13-14 15-16" → ["7-8", "9-10", ...]
    """
    if not raw:
        return []

    # Try to find waist size + age range tokens like "24（7-8）"
    tokens = re.findall(r"\d{2}[（(]\d+-\d+[）)]", raw)
    if tokens:
        # Normalize full-width brackets
        return [t.replace("（", "(").replace("）", ")") for t in tokens]

    # Try age-range-only tokens like "7-8 9-10 ..."
    tokens = re.findall(r"\d+-\d+", raw)
    if tokens:
        return tokens

    # Fallback: plain numbers on the first line
    first_line = raw.split("\n")[0]
    nums = re.findall(r"\d+", first_line)
    return nums if nums else []


# ─────────────────────────────────────────────────────────────────────────────
# Main parse entry point
# ─────────────────────────────────────────────────────────────────────────────

def parse_supplier_pdf_multi(file_bytes: bytes) -> Tuple[List[SupplierStyle], List[SupplierStyle]]:
    """
    Returns (po_styles, size_chart_styles).
    po_styles        – individual color variants from the PO table
    size_chart_styles – size measurement tables from the appended page
    """
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        pages = pdf.pages
        po_styles = parse_po_table(pages)
        size_styles = parse_size_sheet_tables(pages)
    return po_styles, size_styles


def get_raw_text(file_bytes: bytes, max_chars: int = 6000) -> str:
    text_parts = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for i, page in enumerate(pdf.pages):
            t = page.extract_text() or ""
            text_parts.append(f"=== PAGE {i+1} ===\n{t}")
    full = "\n\n".join(text_parts)
    return full[:max_chars]


# ─────────────────────────────────────────────────────────────────────────────
# Streamlit UI
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="PO Size Chart Extractor",
    page_icon="📏",
    layout="wide",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');

html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
}
h1, h2, h3 {
    font-family: 'IBM Plex Mono', monospace !important;
    letter-spacing: -0.03em;
}
.metric-box {
    background: #0f1117;
    border: 1px solid #2a2d3e;
    border-radius: 6px;
    padding: 12px 16px;
    text-align: center;
}
.metric-label {
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.12em;
    color: #888;
    margin-bottom: 4px;
}
.metric-value {
    font-size: 18px;
    font-weight: 600;
    font-family: 'IBM Plex Mono', monospace;
    color: #e8e8e8;
}
.size-badge {
    display: inline-block;
    background: #1a1d2e;
    border: 1px solid #3d4166;
    border-radius: 4px;
    padding: 3px 10px;
    margin: 3px;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 13px;
    color: #7c9eff;
}
.section-header {
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.14em;
    color: #555;
    border-bottom: 1px solid #222;
    padding-bottom: 6px;
    margin-bottom: 12px;
}
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("# 📏 Supplier PO — Size Chart Extractor")
st.markdown(
    "Upload a supplier Purchase Order PDF to extract style information and size measurements. "
    "Use the extracted charts to compare against your **MF Template Matrix Detail** screen."
)
st.divider()

# ── File uploader ─────────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "Drop your supplier PO PDF here",
    type=["pdf"],
    help="PDF must have embedded text (not a scanned image). Supports BOY LONG PANTS and MEN T-SHIRT style POs.",
)

if not uploaded:
    st.info("👆 Upload a PDF to get started. The extractor will find all style rows and measurement tables automatically.")
    st.stop()

file_bytes = uploaded.read()

with st.spinner("Parsing PDF…"):
    po_styles, size_styles = parse_supplier_pdf_multi(file_bytes)

total_po = len(po_styles)
total_size = len(size_styles)

if total_po == 0 and total_size == 0:
    st.warning(
        "⚠️ No styles or measurement tables found in this PDF. "
        "Open the **Debug** expander below to inspect raw text and adjust the parser."
    )
else:
    st.success(
        f"✅ Found **{total_po}** style variant(s) in the PO table "
        f"and **{total_size}** measurement table(s) in the size sheet."
    )

# ── Sidebar navigation ────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🗂 Navigation")

    view_mode = st.radio("View", ["📋 Size Measurement Tables", "🧾 PO Style Variants"], index=0)

    if view_mode == "📋 Size Measurement Tables":
        if size_styles:
            labels = [s.style_code for s in size_styles]
            sel_idx = st.selectbox("Select size chart", range(len(labels)), format_func=lambda i: labels[i])
            selected = size_styles[sel_idx]
        else:
            selected = None
            st.warning("No size tables extracted.")
    else:
        if po_styles:
            labels = [f"{s.style_code} – {s.color}" for s in po_styles]
            sel_idx = st.selectbox("Select style", range(len(labels)), format_func=lambda i: labels[i])
            selected = po_styles[sel_idx]
        else:
            selected = None
            st.warning("No PO style rows extracted.")

    st.divider()
    st.markdown(
        "<div style='font-size:11px;color:#555;'>Use this chart to manually compare<br>with Template Matrix Detail.</div>",
        unsafe_allow_html=True,
    )

# ── Main content ──────────────────────────────────────────────────────────────

if selected is None:
    st.info("Select an item from the sidebar.")
else:
    s = selected

    # Metrics row
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f'<div class="metric-box"><div class="metric-label">Style / Group</div><div class="metric-value">{s.style_code}</div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="metric-box"><div class="metric-label">Product</div><div class="metric-value">{s.product_name}</div></div>', unsafe_allow_html=True)
    with c3:
        st.markdown(f'<div class="metric-box"><div class="metric-label">Color</div><div class="metric-value">{s.color}</div></div>', unsafe_allow_html=True)
    with c4:
        fabric_disp = s.fabric if s.fabric else "—"
        st.markdown(f'<div class="metric-box"><div class="metric-label">Fabric</div><div class="metric-value">{fabric_disp}</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Sizes
    st.markdown('<div class="section-header">Sizes in spec</div>', unsafe_allow_html=True)
    if s.sizes:
        badges = "".join(f'<span class="size-badge">{sz}</span>' for sz in s.sizes)
        st.markdown(badges, unsafe_allow_html=True)
    else:
        st.markdown("*No sizes extracted*")

    st.markdown("<br>", unsafe_allow_html=True)

    # Measurement table + text block
    if s.measurements:
        col_left, col_right = st.columns([1, 1], gap="large")

        with col_left:
            st.markdown('<div class="section-header">Size chart table</div>', unsafe_allow_html=True)

            rows = []
            for m in s.measurements:
                row = {"Measurement": m.name}
                for sz in s.sizes:
                    row[sz] = m.values.get(sz, "")
                rows.append(row)

            df = pd.DataFrame(rows)
            st.dataframe(df, use_container_width=True, hide_index=True)

            # Quick stats
            with st.expander("📊 Quick statistics"):
                st.markdown(f"- **Measurements:** {len(s.measurements)}")
                st.markdown(f"- **Sizes:** {len(s.sizes)}")
                st.markdown(
                    "- 💡 Use this chart to manually compare with **Template Matrix Detail**."
                )

        with col_right:
            st.markdown('<div class="section-header">Readable text block</div>', unsafe_allow_html=True)

            lines = [
                f"Style / Group : {s.style_code}",
                f"Product       : {s.product_name}",
                f"Color         : {s.color}",
                f"Fabric        : {s.fabric or '—'}",
                f"Sizes         : {', '.join(s.sizes)}",
                "",
            ]
            for m in s.measurements:
                lines.append(f"{m.name}:")
                for sz in s.sizes:
                    val = m.values.get(sz, "")
                    lines.append(f"  {sz} = {val} cm")
                lines.append("")

            st.text_area(
                label="Copy-paste block",
                value="\n".join(lines),
                height=420,
                label_visibility="collapsed",
            )

    elif view_mode == "🧾 PO Style Variants":
        # PO rows don't carry measurements themselves; link to size sheets
        st.info(
            "ℹ️ This PO style row does not contain individual measurement data. "
            "Switch to **📋 Size Measurement Tables** in the sidebar to see the appended size charts, "
            "then cross-reference by fabric code or size range."
        )
        st.markdown("**Order details extracted from PO table:**")
        detail_data = {
            "Field": ["Style code", "Product", "Color", "Fabric", "Sizes", "Quantity", "Unit price"],
            "Value": [
                s.style_code,
                s.product_name,
                s.color,
                s.fabric or "—",
                ", ".join(s.sizes) if s.sizes else "—",
                s.raw_qty or "—",
                s.raw_price or "—",
            ],
        }
        st.table(pd.DataFrame(detail_data))
    else:
        st.warning("No measurement data extracted for this entry.")

# ── Debug expander ────────────────────────────────────────────────────────────
st.divider()
with st.expander("🔍 Debug: raw PDF text (first 6 000 chars)"):
    raw = get_raw_text(file_bytes, max_chars=6000)
    st.text_area("Raw extracted text", value=raw, height=400, label_visibility="collapsed")
