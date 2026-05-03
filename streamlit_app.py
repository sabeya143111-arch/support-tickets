"""
Supplier PO Size Chart Extractor  –  v2
New features:
  • 📊 Summary dashboard (total styles, total amount, fabric breakdown)
  • 🔍 Search & filter by color / fabric / style code
  • 📤 Export to Excel (.xlsx) – PO rows + size charts on separate sheets
  • 📌 Side-by-side comparison of any 2 size charts with Δ diff table

No external LLM calls. Uses only: streamlit, pdfplumber, pandas, openpyxl.
"""

import io
import re
from dataclasses import dataclass
from typing import Dict, List, Tuple

import pandas as pd
import pdfplumber
import streamlit as st

# ─────────────────────────────────────────────────────────────────────────────
# Data classes
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class Measurement:
    name: str
    values: Dict[str, str]   # size_label -> value string


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
    raw_amount: str = ""


# ─────────────────────────────────────────────────────────────────────────────
# Measurement mapping
# ─────────────────────────────────────────────────────────────────────────────

MEAS_MAP = {
    "二分之一腰围": "1/2 Waist",
    "二分之一脚口": "1/2 Leg Opening",
    "坐围（直裆1/3处弧度）": "Seat (arc at 1/3 crotch)",
    "坐围": "Seat circumference",
    "裤长": "Pant length",
    "衣长": "Body length",
    "肩宽": "Shoulder width",
    "肩":   "Shoulder width",
    "胸围": "Chest",
    "胸":   "Chest",
    "袖长": "Sleeve length",
    "袖口宽": "Sleeve opening width",
    "袖口":  "Sleeve opening width",
    "下摆": "Bottom hem width",
}

def chinese_to_english(label: str) -> str:
    label = label.strip().rstrip("：:").strip()
    for zh, en in MEAS_MAP.items():
        if zh in label:
            return en
    return label


# ─────────────────────────────────────────────────────────────────────────────
# Parser – appended size-sheet tables (附页合同)
# ─────────────────────────────────────────────────────────────────────────────

def parse_size_sheet_tables(pages) -> List[SupplierStyle]:
    styles: List[SupplierStyle] = []
    for page in pages:
        tables = page.extract_tables()
        if not tables:
            continue
        for table in tables:
            if not table or len(table) < 2:
                continue
            header = [str(c).strip() if c else "" for c in table[0]]
            if not any("码数" in h for h in header):
                continue
            size_start = next((i for i, h in enumerate(header) if "码数" in h), None)
            size_end   = next((i for i, h in enumerate(header) if "单位" in h), None)
            if size_start is None:
                continue
            if size_end is None:
                size_end = len(header)
            sizes = [h for h in header[size_start + 1:size_end] if h]
            if not sizes:
                continue
            measurements: List[Measurement] = []
            for row in table[1:]:
                if not row or not row[0]:
                    continue
                label_raw = str(row[0]).strip()
                if not label_raw or label_raw == "单位":
                    continue
                en_label = chinese_to_english(label_raw)
                vals_raw = [
                    str(row[i]).strip() if i < len(row) and row[i] else ""
                    for i in range(size_start + 1, size_start + 1 + len(sizes))
                ]
                values = {sizes[j]: (vals_raw[j] if j < len(vals_raw) else "") for j in range(len(sizes))}
                measurements.append(Measurement(name=en_label, values=values))
            if not measurements:
                continue
            page_text = page.extract_text() or ""
            if any(f"PN00{n}" in page_text for n in ["7","8","9","10"]):
                group = "PN007-010 Size Chart"
            elif any(f"PN00{n}" in page_text for n in ["2","3","4","5","6"]):
                group = "PN002-006 Size Chart"
            else:
                group = f"Size Chart ({', '.join(sizes)})"
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
# Parser – main PO table
# ─────────────────────────────────────────────────────────────────────────────

def clean_cell(v) -> str:
    return str(v).strip() if v else ""


def parse_po_table(pages) -> List[SupplierStyle]:
    styles: List[SupplierStyle] = []
    all_rows: List[List[str]] = []
    for page in pages:
        for table in (page.extract_tables() or []):
            if not table:
                continue
            flat_header = " ".join(str(c) for c in (table[0] or []))
            if "码数" in flat_header:
                continue
            for row in table:
                all_rows.append([clean_cell(c) for c in row])

    header_idx = next(
        (i for i, r in enumerate(all_rows) if "STYLE" in " ".join(r) and "DESCRIPTION" in " ".join(r)),
        None,
    )
    if header_idx is None:
        return []

    header = all_rows[header_idx]

    def col_idx(keywords):
        for j, h in enumerate(header):
            if any(k.upper() in h.upper() for k in keywords):
                return j
        return None

    col_style  = col_idx(["STYLE"])
    col_desc   = col_idx(["DESCRIPTION"])
    col_color  = col_idx(["COLOR"])
    col_qty    = col_idx(["QTY", "T'QTY"])
    col_price  = col_idx(["PRICE"])
    col_amount = col_idx(["AMOUNT"])
    col_fabric = col_idx(["FABRIC"])
    col_size   = col_idx(["SIZE"])

    if col_style is None or col_color is None:
        return []

    current_desc = current_fabric = current_sizes_raw = ""
    style_pattern = re.compile(r"^(PMY|TF|SW)\d+", re.IGNORECASE)

    for row in all_rows[header_idx + 1:]:
        if not any(row):
            continue

        def get(col):
            return row[col] if col is not None and col < len(row) else ""

        if col_desc   is not None and get(col_desc):   current_desc      = get(col_desc)
        if col_fabric is not None and get(col_fabric): current_fabric    = get(col_fabric)
        if col_size   is not None and get(col_size):   current_sizes_raw = get(col_size)

        style_val = get(col_style)
        color_val = get(col_color)
        if not style_val or not style_pattern.match(style_val) or not color_val:
            continue

        styles.append(SupplierStyle(
            style_code=style_val,
            product_name=current_desc or "BOY LONG PANTS",
            color=color_val,
            fabric=current_fabric,
            sizes=parse_sizes_from_cell(current_sizes_raw),
            measurements=[],
            raw_qty=get(col_qty),
            raw_price=get(col_price),
            raw_amount=get(col_amount),
        ))
    return styles


def parse_sizes_from_cell(raw: str) -> List[str]:
    if not raw:
        return []
    tokens = re.findall(r"\d{2}[（(]\d+-\d+[）)]", raw)
    if tokens:
        return [t.replace("（","(").replace("）",")") for t in tokens]
    tokens = re.findall(r"\d+-\d+", raw)
    if tokens:
        return tokens
    nums = re.findall(r"\d+", raw.split("\n")[0])
    return nums


# ─────────────────────────────────────────────────────────────────────────────
# Main parse entry + raw text helper
# ─────────────────────────────────────────────────────────────────────────────

def parse_supplier_pdf_multi(file_bytes: bytes) -> Tuple[List[SupplierStyle], List[SupplierStyle]]:
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        pages = pdf.pages
        po_styles   = parse_po_table(pages)
        size_styles = parse_size_sheet_tables(pages)
    return po_styles, size_styles


def get_raw_text(file_bytes: bytes, max_chars: int = 6000) -> str:
    parts = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for i, page in enumerate(pdf.pages):
            parts.append(f"=== PAGE {i+1} ===\n{page.extract_text() or ''}")
    return "\n\n".join(parts)[:max_chars]


# ─────────────────────────────────────────────────────────────────────────────
# Excel export
# ─────────────────────────────────────────────────────────────────────────────

def build_excel(po_styles: List[SupplierStyle], size_styles: List[SupplierStyle]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        if po_styles:
            rows = []
            for s in po_styles:
                amt_str = s.raw_amount.replace("¥","").replace(",","").strip()
                try:
                    amt = float(amt_str)
                except ValueError:
                    amt = s.raw_amount
                rows.append({
                    "Style Code": s.style_code,
                    "Product":    s.product_name,
                    "Color":      s.color,
                    "Fabric":     s.fabric,
                    "Sizes":      ", ".join(s.sizes),
                    "Qty (pcs)":  s.raw_qty,
                    "Unit Price": s.raw_price,
                    "Amount":     amt,
                })
            pd.DataFrame(rows).to_excel(writer, sheet_name="PO Styles", index=False)

        for sc in size_styles:
            if not sc.measurements:
                continue
            rows = [
                {"Measurement": m.name, **{sz: m.values.get(sz,"") for sz in sc.sizes}}
                for m in sc.measurements
            ]
            pd.DataFrame(rows).to_excel(writer, sheet_name=sc.style_code[:31], index=False)

    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────────────────────
# Streamlit UI
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="PO Size Chart Extractor", page_icon="📏", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');
html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }
h1,h2,h3 { font-family: 'IBM Plex Mono', monospace !important; letter-spacing: -0.03em; }
.metric-box {
    background: #0f1117; border: 1px solid #2a2d3e;
    border-radius: 6px; padding: 12px 16px; text-align: center;
}
.metric-label { font-size:11px; text-transform:uppercase; letter-spacing:0.12em; color:#888; margin-bottom:4px; }
.metric-value { font-size:18px; font-weight:600; font-family:'IBM Plex Mono',monospace; color:#e8e8e8; }
.size-badge {
    display:inline-block; background:#1a1d2e; border:1px solid #3d4166;
    border-radius:4px; padding:3px 10px; margin:3px;
    font-family:'IBM Plex Mono',monospace; font-size:13px; color:#7c9eff;
}
.section-header {
    font-size:11px; text-transform:uppercase; letter-spacing:0.14em;
    color:#555; border-bottom:1px solid #222; padding-bottom:6px; margin-bottom:12px;
}
.dash-card {
    background:#111827; border:1px solid #1f2937; border-radius:8px;
    padding:20px 24px; margin-bottom:8px;
}
.dash-num  { font-size:32px; font-weight:700; font-family:'IBM Plex Mono',monospace; color:#60a5fa; }
.dash-desc { font-size:12px; color:#9ca3af; margin-top:2px; }
</style>
""", unsafe_allow_html=True)

# ── Header ───────────────────────────────────────────────────────────────────
st.markdown("# 📏 Supplier PO — Size Chart Extractor")
st.markdown(
    "Upload a supplier Purchase Order PDF to extract styles and measurements. "
    "Compare against your **MF Template Matrix Detail** screen."
)
st.divider()

# ── File upload ───────────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "Drop your supplier PO PDF here",
    type=["pdf"],
    help="PDF must have embedded text. Supports BOY LONG PANTS and MEN T-SHIRT POs.",
)

if not uploaded:
    st.info("👆 Upload a PDF to get started.")
    st.stop()

file_bytes = uploaded.read()

with st.spinner("Parsing PDF…"):
    po_styles, size_styles = parse_supplier_pdf_multi(file_bytes)

total_po   = len(po_styles)
total_size = len(size_styles)

if total_po == 0 and total_size == 0:
    st.warning("⚠️ No styles or measurement tables found. Check the Debug section below.")
else:
    st.success(
        f"✅ Found **{total_po}** PO style variant(s) and **{total_size}** size measurement table(s)."
    )

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🗂 Navigation")
    view_mode = st.radio(
        "View",
        ["📊 Dashboard", "📋 Size Charts", "🧾 PO Styles", "📌 Compare Charts"],
        index=0,
    )

    if view_mode in ["🧾 PO Styles", "📌 Compare Charts"]:
        st.divider()
        st.markdown("#### 🔍 Filter PO Styles")
        search_text   = st.text_input("Search style code", placeholder="e.g. PMY009")
        all_colors    = sorted({s.color  for s in po_styles if s.color})
        all_fabrics   = sorted({s.fabric for s in po_styles if s.fabric})
        filter_color  = st.multiselect("Color",  all_colors)
        filter_fabric = st.multiselect("Fabric", all_fabrics)

        filtered_po = po_styles
        if search_text:
            filtered_po = [s for s in filtered_po if search_text.lower() in s.style_code.lower()]
        if filter_color:
            filtered_po = [s for s in filtered_po if s.color  in filter_color]
        if filter_fabric:
            filtered_po = [s for s in filtered_po if s.fabric in filter_fabric]
    else:
        filtered_po = po_styles

    st.divider()
    st.markdown(
        "<div style='font-size:11px;color:#555;'>Compare with<br>Template Matrix Detail.</div>",
        unsafe_allow_html=True,
    )


# ═══════════════════════════════════════════════════════════════════════════
# 📊 DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════
if view_mode == "📊 Dashboard":
    st.subheader("📊 Order Summary Dashboard")

    total_qty = 0
    total_amt = 0.0
    fabric_counts: Dict[str,int] = {}

    for s in po_styles:
        qty_str = re.sub(r"[^\d]", "", s.raw_qty)
        if qty_str:
            total_qty += int(qty_str)
        amt_str = s.raw_amount.replace("¥","").replace(",","").strip()
        try:
            total_amt += float(amt_str)
        except ValueError:
            pass
        fab = s.fabric or "Unknown"
        fabric_counts[fab] = fabric_counts.get(fab, 0) + 1

    unique_styles  = len({s.style_code.split("-")[0] for s in po_styles})
    color_variants = total_po

    c1, c2, c3, c4 = st.columns(4)
    for col, num, desc in [
        (c1, str(unique_styles),   "Unique Style Models"),
        (c2, str(color_variants),  "Color Variants"),
        (c3, f"{total_qty:,}",     "Total Pieces"),
        (c4, f"¥{total_amt:,.0f}", "Total Order Value"),
    ]:
        with col:
            st.markdown(
                f'<div class="dash-card"><div class="dash-num">{num}</div>'
                f'<div class="dash-desc">{desc}</div></div>',
                unsafe_allow_html=True,
            )

    st.markdown("<br>", unsafe_allow_html=True)
    col_l, col_r = st.columns([1, 1], gap="large")

    with col_l:
        st.markdown('<div class="section-header">Fabric Breakdown</div>', unsafe_allow_html=True)
        fab_df = pd.DataFrame(
            sorted(fabric_counts.items(), key=lambda x: -x[1]),
            columns=["Fabric", "Style Variants"],
        )
        st.dataframe(fab_df, use_container_width=True, hide_index=True)

    with col_r:
        st.markdown('<div class="section-header">Top Styles by Order Value</div>', unsafe_allow_html=True)
        rows = []
        for s in po_styles:
            amt_str = s.raw_amount.replace("¥","").replace(",","").strip()
            try:
                amt = float(amt_str)
            except ValueError:
                amt = 0.0
            rows.append({"Style": s.style_code, "Color": s.color, "Amount (¥)": amt})
        if rows:
            top_df = (
                pd.DataFrame(rows)
                .sort_values("Amount (¥)", ascending=False)
                .head(10)
                .reset_index(drop=True)
            )
            st.dataframe(top_df, use_container_width=True, hide_index=True)

    # Excel export button
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="section-header">Export</div>', unsafe_allow_html=True)
    if st.button("📤 Generate Excel Export"):
        with st.spinner("Building Excel file…"):
            xlsx_bytes = build_excel(po_styles, size_styles)
        st.download_button(
            label="⬇️  Download .xlsx",
            data=xlsx_bytes,
            file_name=f"PO_{uploaded.name.replace('.pdf','')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ═══════════════════════════════════════════════════════════════════════════
# 📋 SIZE CHARTS
# ═══════════════════════════════════════════════════════════════════════════
elif view_mode == "📋 Size Charts":
    st.subheader("📋 Size Measurement Tables")

    if not size_styles:
        st.warning("No size measurement tables extracted from this PDF.")
        st.stop()

    labels  = [s.style_code for s in size_styles]
    sel_idx = st.selectbox("Select size chart", range(len(labels)), format_func=lambda i: labels[i])
    s       = size_styles[sel_idx]

    c1, c2, c3, c4 = st.columns(4)
    for col, lbl, val in [
        (c1, "Group",   s.style_code),
        (c2, "Product", s.product_name),
        (c3, "Color",   s.color),
        (c4, "Fabric",  s.fabric or "—"),
    ]:
        with col:
            st.markdown(
                f'<div class="metric-box"><div class="metric-label">{lbl}</div>'
                f'<div class="metric-value">{val}</div></div>',
                unsafe_allow_html=True,
            )

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="section-header">Sizes in spec</div>', unsafe_allow_html=True)
    st.markdown("".join(f'<span class="size-badge">{sz}</span>' for sz in s.sizes), unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    if s.measurements:
        col_l, col_r = st.columns([1, 1], gap="large")

        with col_l:
            st.markdown('<div class="section-header">Size chart table</div>', unsafe_allow_html=True)
            df = pd.DataFrame([
                {"Measurement": m.name, **{sz: m.values.get(sz,"") for sz in s.sizes}}
                for m in s.measurements
            ])
            st.dataframe(df, use_container_width=True, hide_index=True)
            with st.expander("📊 Quick statistics"):
                st.markdown(f"- **Measurements:** {len(s.measurements)}")
                st.markdown(f"- **Sizes:** {len(s.sizes)}")

        with col_r:
            st.markdown('<div class="section-header">Readable text block</div>', unsafe_allow_html=True)
            lines = [
                f"Group   : {s.style_code}",
                f"Product : {s.product_name}",
                f"Color   : {s.color}",
                f"Fabric  : {s.fabric or '—'}",
                f"Sizes   : {', '.join(s.sizes)}", "",
            ]
            for m in s.measurements:
                lines.append(f"{m.name}:")
                for sz in s.sizes:
                    lines.append(f"  {sz} = {m.values.get(sz,'')} cm")
                lines.append("")
            st.text_area("", value="\n".join(lines), height=420, label_visibility="collapsed")
    else:
        st.warning("No measurement data found.")


# ═══════════════════════════════════════════════════════════════════════════
# 🧾 PO STYLES  (search + filter)
# ═══════════════════════════════════════════════════════════════════════════
elif view_mode == "🧾 PO Styles":
    st.subheader("🧾 PO Style Variants")

    if not filtered_po:
        st.warning("No styles match your filter. Try clearing the sidebar filters.")
        st.stop()

    st.caption(f"Showing **{len(filtered_po)}** of **{total_po}** styles")

    labels  = [f"{s.style_code} – {s.color}" for s in filtered_po]
    sel_idx = st.selectbox("Select style", range(len(labels)), format_func=lambda i: labels[i])
    s       = filtered_po[sel_idx]

    c1, c2, c3, c4, c5 = st.columns(5)
    for col, lbl, val in [
        (c1, "Style",   s.style_code),
        (c2, "Product", s.product_name),
        (c3, "Color",   s.color),
        (c4, "Fabric",  s.fabric or "—"),
        (c5, "Qty",     s.raw_qty  or "—"),
    ]:
        with col:
            st.markdown(
                f'<div class="metric-box"><div class="metric-label">{lbl}</div>'
                f'<div class="metric-value">{val}</div></div>',
                unsafe_allow_html=True,
            )

    st.markdown("<br>", unsafe_allow_html=True)
    st.info(
        "ℹ️ Individual measurements are in **📋 Size Charts**. "
        "Switch to **📌 Compare Charts** to put two tables side by side."
    )

    st.table(pd.DataFrame({
        "Field": ["Style code","Product","Color","Fabric","Sizes","Quantity","Unit price","Amount"],
        "Value": [
            s.style_code, s.product_name, s.color, s.fabric or "—",
            ", ".join(s.sizes) if s.sizes else "—",
            s.raw_qty or "—", s.raw_price or "—", s.raw_amount or "—",
        ],
    }))

    st.markdown('<div class="section-header">All filtered results</div>', unsafe_allow_html=True)
    st.dataframe(pd.DataFrame([{
        "Style": s.style_code, "Product": s.product_name, "Color": s.color,
        "Fabric": s.fabric, "Sizes": ", ".join(s.sizes),
        "Qty": s.raw_qty, "Price": s.raw_price, "Amount": s.raw_amount,
    } for s in filtered_po]), use_container_width=True, hide_index=True)


# ═══════════════════════════════════════════════════════════════════════════
# 📌 COMPARE CHARTS
# ═══════════════════════════════════════════════════════════════════════════
elif view_mode == "📌 Compare Charts":
    st.subheader("📌 Side-by-Side Size Chart Comparison")

    all_charts = size_styles + [s for s in po_styles if s.measurements]

    if not all_charts:
        st.warning("No size charts available to compare.")
        st.stop()

    chart_labels = [s.style_code for s in all_charts]

    col_pick1, col_pick2 = st.columns(2)
    with col_pick1:
        idx1 = st.selectbox("Chart A", range(len(chart_labels)), format_func=lambda i: chart_labels[i], key="cmp_a")
    with col_pick2:
        idx2 = st.selectbox("Chart B", range(len(chart_labels)),
                             index=min(1, len(chart_labels)-1),
                             format_func=lambda i: chart_labels[i], key="cmp_b")

    sa, sb = all_charts[idx1], all_charts[idx2]

    def make_df(s: SupplierStyle) -> pd.DataFrame:
        if not s.measurements:
            return pd.DataFrame()
        return pd.DataFrame([
            {"Measurement": m.name, **{sz: m.values.get(sz,"") for sz in s.sizes}}
            for m in s.measurements
        ])

    st.markdown("<br>", unsafe_allow_html=True)
    left, right = st.columns(2, gap="large")

    with left:
        st.markdown(f'<div class="section-header">Chart A — {sa.style_code}</div>', unsafe_allow_html=True)
        st.markdown("".join(f'<span class="size-badge">{sz}</span>' for sz in sa.sizes), unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        df_a = make_df(sa)
        st.dataframe(df_a, use_container_width=True, hide_index=True) if not df_a.empty else st.warning("No data.")

    with right:
        st.markdown(f'<div class="section-header">Chart B — {sb.style_code}</div>', unsafe_allow_html=True)
        st.markdown("".join(f'<span class="size-badge">{sz}</span>' for sz in sb.sizes), unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        df_b = make_df(sb)
        st.dataframe(df_b, use_container_width=True, hide_index=True) if not df_b.empty else st.warning("No data.")

    # Δ difference table
    if not df_a.empty and not df_b.empty:
        common_meas  = set(df_a["Measurement"]) & set(df_b["Measurement"])
        common_sizes = [sz for sz in sa.sizes if sz in sb.sizes]

        if common_meas and common_sizes:
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown('<div class="section-header">Δ Difference (A − B)</div>', unsafe_allow_html=True)

            diff_rows = []
            for meas in sorted(common_meas):
                row_a = df_a[df_a["Measurement"] == meas].iloc[0]
                row_b = df_b[df_b["Measurement"] == meas].iloc[0]
                diff_row = {"Measurement": meas}
                for sz in common_sizes:
                    try:
                        diff_row[sz] = round(float(row_a[sz]) - float(row_b[sz]), 2)
                    except (ValueError, KeyError):
                        diff_row[sz] = "—"
                diff_rows.append(diff_row)

            diff_df = pd.DataFrame(diff_rows)

            def color_diff(val):
                if val == "—" or val == 0:
                    return ""
                try:
                    return "color: #f87171" if float(val) < 0 else "color: #34d399"
                except (ValueError, TypeError):
                    return ""

            numeric_cols = [c for c in diff_df.columns if c != "Measurement"]
            styled = diff_df.style.map(color_diff, subset=numeric_cols)
            st.dataframe(styled, use_container_width=True, hide_index=True)
            st.caption("🟢 Green = A is larger  |  🔴 Red = B is larger  |  0 = identical")


# ── Debug ─────────────────────────────────────────────────────────────────────
st.divider()
with st.expander("🔍 Debug: raw PDF text (first 6 000 chars)"):
    raw = get_raw_text(file_bytes, max_chars=6000)
    st.text_area("", value=raw, height=400, label_visibility="collapsed")
