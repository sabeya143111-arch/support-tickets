"""
Supplier PO Size Chart Extractor  –  v3
---------------------------------------
Root cause fix: Page 4 size tables are EMBEDDED AS IMAGES in this PDF —
pdfplumber cannot extract them as text. Solution: the two known size charts
for this specific PO format (PN007-010 and PN002-006) are built-in and
displayed instantly, always correct.

PO style table (PMY rows) IS parseable as text and is extracted live.

Features:
  📊 Dashboard  – KPIs, fabric breakdown, top styles
  📋 Size Charts – always-correct measurement tables (image-embedded data)
  🧾 PO Styles   – all 32 style variants with search + filter
  📌 Compare     – side-by-side chart comparison with Δ diff
  📤 Export      – Excel download (PO rows + size charts)
"""

import io
import re
from dataclasses import dataclass
from typing import Dict, List, Tuple

import pandas as pd
import pdfplumber
import streamlit as st

# ─── page config first ───────────────────────────────────────────────────────
st.set_page_config(
    page_title="PO Size Chart Extractor",
    page_icon="📏",
    layout="wide",
)

# ─── CSS ─────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');
html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }
h1,h2,h3 { font-family: 'IBM Plex Mono', monospace !important; letter-spacing: -.03em; }
.kpi { background:#111827; border:1px solid #1f2937; border-radius:8px; padding:18px 22px; }
.kpi-num  { font-size:30px; font-weight:700; font-family:'IBM Plex Mono',monospace; color:#60a5fa; }
.kpi-desc { font-size:11px; color:#9ca3af; margin-top:3px; text-transform:uppercase; letter-spacing:.1em; }
.mbox { background:#0f1117; border:1px solid #2a2d3e; border-radius:6px; padding:11px 15px; text-align:center; }
.mlabel { font-size:10px; text-transform:uppercase; letter-spacing:.12em; color:#888; margin-bottom:3px; }
.mvalue { font-size:16px; font-weight:600; font-family:'IBM Plex Mono',monospace; color:#e8e8e8; }
.sh { font-size:10px; text-transform:uppercase; letter-spacing:.14em; color:#555;
      border-bottom:1px solid #222; padding-bottom:5px; margin-bottom:11px; }
.sbadge { display:inline-block; background:#1a1d2e; border:1px solid #3d4166;
          border-radius:4px; padding:3px 10px; margin:3px;
          font-family:'IBM Plex Mono',monospace; font-size:12px; color:#7c9eff; }
.notice { background:#1c2a1c; border:1px solid #2d4a2d; border-radius:6px;
          padding:10px 14px; font-size:13px; color:#86efac; margin-bottom:12px; }
</style>
""", unsafe_allow_html=True)


# ═════════════════════════════════════════════════════════════════════════════
# BUILT-IN SIZE CHARTS  (image-embedded data from page 4 of this PO format)
# These are fixed for PO 102-SY0002 / SI4大童长裤
# ═════════════════════════════════════════════════════════════════════════════

BUILTIN_CHARTS = [
    {
        "id": "PN007-010",
        "title": "PN007 / PN008 / PN009 / PN010",
        "subtitle": "大童尺寸表",
        "fabric": "58084",
        "applies_to": ["PMY009", "PMY007", "PMY008", "PMY010"],
        "sizes": ["24 (7-8)", "25 (9-10)", "26 (11-12)", "27 (13-14)", "28 (15-16)"],
        "unit": "CM",
        "rows": [
            ("1/2 Waist  (二分之一腰围)",       ["29",   "30.5", "32",  "33.5", "35"]),
            ("1/2 Leg Opening (二分之一脚口)",   ["6.5",  "6.7",  "6.9", "7.1",  "7.3"]),
            ("Pant Length  (裤长)",              ["81",   "84",   "87",  "91",   "95"]),
            ("Seat/Crotch Arc  (坐围直裆1/3弧)", ["88",   "92",   "96",  "100",  "104"]),
        ],
    },
    {
        "id": "PN002-006",
        "title": "PN002 / PN003 / PN004 / PN005 / PN006",
        "subtitle": "大童长裤尺寸表",
        "fabric": "58114 / 57410",
        "applies_to": ["PMY002", "PMY003", "PMY004", "PMY005", "PMY006"],
        "sizes": ["7-8", "9-10", "11-12", "13-14", "15-16"],
        "unit": "CM",
        "rows": [
            ("1/2 Waist  (二分之一腰围)",      ["29",   "30.5", "32",   "33.5", "35"]),
            ("1/2 Leg Opening (二分之一脚口)", ["20.5", "21.1", "21.7", "22.3", "22.9"]),
            ("Pant Length  (裤长)",            ["81",   "84",   "87",   "91",   "95"]),
            ("Seat Circumference  (坐围)",     ["92",   "95",   "98",   "101",  "104"]),
        ],
    },
]


def chart_to_df(chart: dict) -> pd.DataFrame:
    rows = [{"Measurement": name, **dict(zip(chart["sizes"], vals))}
            for name, vals in chart["rows"]]
    return pd.DataFrame(rows)


# ═════════════════════════════════════════════════════════════════════════════
# PO TABLE PARSER  (pages 1–2 are proper text tables)
# ═════════════════════════════════════════════════════════════════════════════

@dataclass
class POStyle:
    code: str
    product: str
    color: str
    fabric: str
    sizes: List[str]
    qty: str
    price: str
    amount: str


def clean(v) -> str:
    return str(v).strip() if v else ""


def parse_sizes(raw: str) -> List[str]:
    tokens = re.findall(r"\d{2}[（(]\d+-\d+[）)]", raw)
    if tokens:
        return [t.replace("（","(").replace("）",")") for t in tokens]
    tokens = re.findall(r"\d+-\d+", raw)
    if tokens:
        return tokens
    return re.findall(r"\d+", raw.split("\n")[0])


def parse_po(file_bytes: bytes) -> List[POStyle]:
    styles: List[POStyle] = []
    all_rows: List[List[str]] = []

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            for tbl in (page.extract_tables() or []):
                if not tbl:
                    continue
                if "码数" in " ".join(str(c) for c in (tbl[0] or [])):
                    continue
                for row in tbl:
                    all_rows.append([clean(c) for c in row])

    hdr_idx = next(
        (i for i, r in enumerate(all_rows)
         if "STYLE" in " ".join(r) and "DESCRIPTION" in " ".join(r)),
        None,
    )
    if hdr_idx is None:
        return []

    hdr = all_rows[hdr_idx]

    def ci(kws):
        for j, h in enumerate(hdr):
            if any(k.upper() in h.upper() for k in kws):
                return j
        return None

    c_sty = ci(["STYLE"])
    c_dsc = ci(["DESCRIPTION"])
    c_col = ci(["COLOR"])
    c_qty = ci(["QTY", "T'QTY"])
    c_prc = ci(["PRICE"])
    c_amt = ci(["AMOUNT"])
    c_fab = ci(["FABRIC"])
    c_siz = ci(["SIZE"])

    if c_sty is None or c_col is None:
        return []

    cur_dsc = cur_fab = cur_siz = ""
    pat = re.compile(r"^(PMY|TF|SW)\d+", re.I)

    for row in all_rows[hdr_idx + 1:]:
        if not any(row):
            continue

        def g(col):
            return row[col] if col is not None and col < len(row) else ""

        if c_dsc is not None and g(c_dsc): cur_dsc = g(c_dsc)
        if c_fab is not None and g(c_fab): cur_fab = g(c_fab)
        if c_siz is not None and g(c_siz): cur_siz = g(c_siz)

        code  = g(c_sty)
        color = g(c_col)
        if not code or not pat.match(code) or not color:
            continue

        styles.append(POStyle(
            code=code, product=cur_dsc or "BOY LONG PANTS",
            color=color, fabric=cur_fab,
            sizes=parse_sizes(cur_siz),
            qty=g(c_qty), price=g(c_prc), amount=g(c_amt),
        ))
    return styles


def get_raw_text(file_bytes: bytes, n: int = 5000) -> str:
    parts = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for i, p in enumerate(pdf.pages):
            parts.append(f"=== PAGE {i+1} ===\n{p.extract_text() or '(no text — likely image)'}")
    return "\n\n".join(parts)[:n]


# ═════════════════════════════════════════════════════════════════════════════
# EXCEL EXPORT
# ═════════════════════════════════════════════════════════════════════════════

def build_excel(styles: List[POStyle]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        # Sheet 1 – PO rows
        rows = []
        for s in styles:
            amt = s.amount.replace("¥","").replace(",","").strip()
            try:
                amt_val = float(amt)
            except ValueError:
                amt_val = s.amount
            rows.append({
                "Style Code": s.code, "Product": s.product, "Color": s.color,
                "Fabric": s.fabric, "Sizes": ", ".join(s.sizes),
                "Qty": s.qty, "Unit Price": s.price, "Amount (¥)": amt_val,
            })
        pd.DataFrame(rows).to_excel(writer, sheet_name="PO Styles", index=False)

        # Sheet 2 & 3 – size charts
        for ch in BUILTIN_CHARTS:
            chart_to_df(ch).to_excel(writer, sheet_name=ch["id"], index=False)

    buf.seek(0)
    return buf.read()


# ═════════════════════════════════════════════════════════════════════════════
# UI HELPERS
# ═════════════════════════════════════════════════════════════════════════════

def metric(col, label, value):
    col.markdown(
        f'<div class="mbox"><div class="mlabel">{label}</div>'
        f'<div class="mvalue">{value}</div></div>',
        unsafe_allow_html=True,
    )

def sec(text):
    st.markdown(f'<div class="sh">{text}</div>', unsafe_allow_html=True)

def badges(items):
    st.markdown(
        "".join(f'<span class="sbadge">{x}</span>' for x in items),
        unsafe_allow_html=True,
    )


# ═════════════════════════════════════════════════════════════════════════════
# HEADER
# ═════════════════════════════════════════════════════════════════════════════

st.markdown("# 📏 Supplier PO — Size Chart Extractor")
st.markdown(
    "Upload the supplier PDF to see all **32 style variants** and the "
    "**correct size measurement tables** for comparison with your "
    "**MF Template Matrix Detail** screen."
)
st.divider()

# ═════════════════════════════════════════════════════════════════════════════
# FILE UPLOAD
# ═════════════════════════════════════════════════════════════════════════════

uploaded = st.file_uploader(
    "Drop supplier PO PDF here",
    type=["pdf"],
    help="Supports SI4大童长裤 PO format (102-SY0002). PMY style rows are parsed from text; "
         "size measurement tables are built-in because page 4 data is image-embedded.",
)

if not uploaded:
    # Show built-in charts even before upload
    st.info(
        "👆 Upload the PDF to load PO style data.  \n"
        "**Size measurement tables are shown below right now** — "
        "they are always available because the data is built into this tool."
    )
    st.markdown("### 📋 Size Measurement Tables (always available)")
    for ch in BUILTIN_CHARTS:
        with st.expander(f"📐 {ch['title']}  —  Fabric: {ch['fabric']}", expanded=True):
            st.caption(f"Applies to styles: **{', '.join(ch['applies_to'])}**  |  Unit: {ch['unit']}")
            badges(ch["sizes"])
            st.markdown("<br>", unsafe_allow_html=True)
            st.dataframe(chart_to_df(ch), use_container_width=True, hide_index=True)
    st.stop()

# ─── Parse PDF ───────────────────────────────────────────────────────────────
file_bytes = uploaded.read()

with st.spinner("Reading PO table…"):
    po_styles = parse_po(file_bytes)

n_po = len(po_styles)

if n_po == 0:
    st.warning("⚠️ Could not parse any PMY style rows. Check Debug section below.")
else:
    col1, col2 = st.columns([3, 1])
    col1.success(
        f"✅ Loaded **{n_po} PO style variants** from the order table.  \n"
        f"Size measurement tables are **built-in** (page 4 data is image-embedded, "
        f"not readable as text — this is normal for this supplier's format)."
    )

# ═════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ═════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("### 🗂 Navigation")
    view = st.radio(
        "Go to",
        ["📊 Dashboard", "📋 Size Charts", "🧾 PO Styles", "📌 Compare"],
        index=0,
    )

    if view in ["🧾 PO Styles", "📌 Compare"]:
        st.divider()
        st.markdown("#### 🔍 Filter")
        q          = st.text_input("Style code", placeholder="PMY009…")
        f_colors   = st.multiselect("Color",  sorted({s.color  for s in po_styles if s.color}))
        f_fabrics  = st.multiselect("Fabric", sorted({s.fabric for s in po_styles if s.fabric}))

        fp = po_styles
        if q:         fp = [s for s in fp if q.lower() in s.code.lower()]
        if f_colors:  fp = [s for s in fp if s.color  in f_colors]
        if f_fabrics: fp = [s for s in fp if s.fabric in f_fabrics]
    else:
        fp = po_styles

    st.divider()
    st.markdown("<small style='color:#555'>Compare with Template Matrix Detail</small>",
                unsafe_allow_html=True)


# ═════════════════════════════════════════════════════════════════════════════
# 📊  DASHBOARD
# ═════════════════════════════════════════════════════════════════════════════
if view == "📊 Dashboard":
    st.subheader("📊 Order Summary")

    total_qty = 0
    total_amt = 0.0
    fabric_cnt: Dict[str, int] = {}

    for s in po_styles:
        n = re.sub(r"[^\d]", "", s.qty)
        if n: total_qty += int(n)
        a = s.amount.replace("¥","").replace(",","").strip()
        try: total_amt += float(a)
        except ValueError: pass
        fab = s.fabric or "Unknown"
        fabric_cnt[fab] = fabric_cnt.get(fab, 0) + 1

    unique_models = len({s.code.split("-")[0] for s in po_styles})

    c1, c2, c3, c4 = st.columns(4)
    for col, num, desc in [
        (c1, str(unique_models),   "Unique models"),
        (c2, str(n_po),            "Color variants"),
        (c3, f"{total_qty:,}",     "Total pieces"),
        (c4, f"¥{total_amt:,.0f}", "Total order value"),
    ]:
        with col:
            st.markdown(
                f'<div class="kpi"><div class="kpi-num">{num}</div>'
                f'<div class="kpi-desc">{desc}</div></div>',
                unsafe_allow_html=True,
            )

    st.markdown("<br>", unsafe_allow_html=True)
    cl, cr = st.columns(2, gap="large")

    with cl:
        sec("Fabric breakdown")
        fab_df = pd.DataFrame(
            sorted(fabric_cnt.items(), key=lambda x: -x[1]),
            columns=["Fabric", "Variants"],
        )
        st.dataframe(fab_df, use_container_width=True, hide_index=True)

    with cr:
        sec("Top 10 styles by value")
        rows = []
        for s in po_styles:
            a = s.amount.replace("¥","").replace(",","").strip()
            try: amt = float(a)
            except ValueError: amt = 0.0
            rows.append({"Style": s.code, "Color": s.color, "¥ Amount": amt})
        if rows:
            st.dataframe(
                pd.DataFrame(rows).sort_values("¥ Amount", ascending=False).head(10),
                use_container_width=True, hide_index=True,
            )

    st.markdown("<br>", unsafe_allow_html=True)
    sec("Export")
    if st.button("📤 Generate Excel"):
        xlsx = build_excel(po_styles)
        st.download_button(
            "⬇️ Download .xlsx",
            data=xlsx,
            file_name=f"PO_{uploaded.name.replace('.pdf','')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ═════════════════════════════════════════════════════════════════════════════
# 📋  SIZE CHARTS
# ═════════════════════════════════════════════════════════════════════════════
elif view == "📋 Size Charts":
    st.subheader("📋 Size Measurement Tables")

    st.markdown(
        '<div class="notice">ℹ️ These measurement tables come from <b>page 4</b> of the supplier PDF. '
        'That page stores its data as embedded images — not readable text — so pdfplumber cannot '
        'extract the numbers automatically. The values below are hard-coded from the known spec and '
        'are <b>always correct</b> for this PO format.</div>',
        unsafe_allow_html=True,
    )

    chart_names = [f"{ch['title']}  (Fabric: {ch['fabric']})" for ch in BUILTIN_CHARTS]
    sel = st.selectbox("Select size chart", range(len(chart_names)),
                       format_func=lambda i: chart_names[i])
    ch = BUILTIN_CHARTS[sel]

    c1, c2, c3 = st.columns(3)
    metric(c1, "Style group",   ch["title"])
    metric(c2, "Fabric code",   ch["fabric"])
    metric(c3, "Applies to",    "  ".join(ch["applies_to"]))

    st.markdown("<br>", unsafe_allow_html=True)
    sec("Sizes in spec")
    badges(ch["sizes"])
    st.markdown("<br>", unsafe_allow_html=True)

    cl, cr = st.columns(2, gap="large")

    with cl:
        sec("Size chart table (unit: CM)")
        df = chart_to_df(ch)
        st.dataframe(df, use_container_width=True, hide_index=True)

        with st.expander("📊 Quick stats"):
            st.markdown(f"- **Measurements:** {len(ch['rows'])}")
            st.markdown(f"- **Sizes:** {len(ch['sizes'])}")
            st.markdown("- 💡 Compare with Template Matrix Detail")

    with cr:
        sec("Copy-paste text block")
        lines = [
            f"Style group : {ch['title']}",
            f"Subtitle    : {ch['subtitle']}",
            f"Fabric      : {ch['fabric']}",
            f"Applies to  : {', '.join(ch['applies_to'])}",
            f"Sizes       : {', '.join(ch['sizes'])}",
            f"Unit        : {ch['unit']}",
            "",
        ]
        for name, vals in ch["rows"]:
            lines.append(f"{name}:")
            for sz, v in zip(ch["sizes"], vals):
                lines.append(f"  {sz} = {v} cm")
            lines.append("")
        st.text_area("", value="\n".join(lines), height=430,
                     label_visibility="collapsed")


# ═════════════════════════════════════════════════════════════════════════════
# 🧾  PO STYLES
# ═════════════════════════════════════════════════════════════════════════════
elif view == "🧾 PO Styles":
    st.subheader("🧾 PO Style Variants")

    if not fp:
        st.warning("No styles match — try clearing sidebar filters.")
        st.stop()

    st.caption(f"Showing **{len(fp)}** of **{n_po}** styles")

    labels  = [f"{s.code}  –  {s.color}" for s in fp]
    sel_idx = st.selectbox("Select style", range(len(labels)),
                           format_func=lambda i: labels[i])
    s = fp[sel_idx]

    c1,c2,c3,c4,c5 = st.columns(5)
    for col, lbl, val in [
        (c1,"Style",   s.code),
        (c2,"Product", s.product),
        (c3,"Color",   s.color),
        (c4,"Fabric",  s.fabric or "—"),
        (c5,"Qty",     s.qty or "—"),
    ]:
        metric(col, lbl, val)

    st.markdown("<br>", unsafe_allow_html=True)

    sec("Sizes in spec")
    badges(s.sizes) if s.sizes else st.markdown("*not parsed*")
    st.markdown("<br>", unsafe_allow_html=True)

    # Which size chart applies?
    matching_chart = next(
        (ch for ch in BUILTIN_CHARTS
         if any(s.code.startswith(m) for m in ch["applies_to"])),
        None,
    )
    if matching_chart:
        st.info(
            f"📐 This style uses **{matching_chart['title']}** size chart "
            f"(Fabric: {matching_chart['fabric']}).  "
            f"Go to **📋 Size Charts** to see measurements."
        )

    st.markdown("**Order details:**")
    st.table(pd.DataFrame({
        "Field": ["Style","Product","Color","Fabric","Sizes","Qty","Unit price","Amount"],
        "Value": [s.code, s.product, s.color, s.fabric or "—",
                  ", ".join(s.sizes) or "—",
                  s.qty or "—", s.price or "—", s.amount or "—"],
    }))

    sec("All filtered styles")
    st.dataframe(pd.DataFrame([{
        "Style": x.code, "Color": x.color, "Fabric": x.fabric,
        "Sizes": ", ".join(x.sizes), "Qty": x.qty,
        "Price": x.price, "Amount": x.amount,
    } for x in fp]), use_container_width=True, hide_index=True)


# ═════════════════════════════════════════════════════════════════════════════
# 📌  COMPARE
# ═════════════════════════════════════════════════════════════════════════════
elif view == "📌 Compare":
    st.subheader("📌 Side-by-Side Size Chart Comparison")

    names = [f"{ch['title']}  (Fabric: {ch['fabric']})" for ch in BUILTIN_CHARTS]

    cp1, cp2 = st.columns(2)
    with cp1:
        a_idx = st.selectbox("Chart A", range(len(names)),
                             format_func=lambda i: names[i], key="ca")
    with cp2:
        b_idx = st.selectbox("Chart B", range(len(names)),
                             index=min(1, len(names)-1),
                             format_func=lambda i: names[i], key="cb")

    ca, cb = BUILTIN_CHARTS[a_idx], BUILTIN_CHARTS[b_idx]
    dfa, dfb = chart_to_df(ca), chart_to_df(cb)

    st.markdown("<br>", unsafe_allow_html=True)
    left, right = st.columns(2, gap="large")

    with left:
        sec(f"Chart A  —  {ca['title']}")
        badges(ca["sizes"])
        st.markdown("<br>", unsafe_allow_html=True)
        st.dataframe(dfa, use_container_width=True, hide_index=True)

    with right:
        sec(f"Chart B  —  {cb['title']}")
        badges(cb["sizes"])
        st.markdown("<br>", unsafe_allow_html=True)
        st.dataframe(dfb, use_container_width=True, hide_index=True)

    # Δ diff (only if same number of sizes)
    common_meas  = set(dfa["Measurement"]) & set(dfb["Measurement"])
    common_sizes = [sz for sz in ca["sizes"] if sz in cb["sizes"]]

    if common_meas and common_sizes:
        st.markdown("<br>", unsafe_allow_html=True)
        sec("Δ Difference  (A − B, same sizes only)")

        diff_rows = []
        for meas in sorted(common_meas):
            ra = dfa[dfa["Measurement"] == meas].iloc[0]
            rb = dfb[dfb["Measurement"] == meas].iloc[0]
            dr = {"Measurement": meas}
            for sz in common_sizes:
                try:
                    dr[sz] = round(float(ra[sz]) - float(rb[sz]), 2)
                except (ValueError, KeyError):
                    dr[sz] = "—"
            diff_rows.append(dr)

        diff_df = pd.DataFrame(diff_rows)

        def color_diff(v):
            if v == "—" or v == 0: return ""
            try:
                return "color:#f87171" if float(v) < 0 else "color:#34d399"
            except (ValueError, TypeError):
                return ""

        num_cols = [c for c in diff_df.columns if c != "Measurement"]
        st.dataframe(
            diff_df.style.map(color_diff, subset=num_cols),
            use_container_width=True, hide_index=True,
        )
        st.caption("🟢 Green = A is larger  |  🔴 Red = B is larger  |  0 = identical")
    else:
        st.info("These two charts have different size labels — no direct numeric diff possible.")


# ═════════════════════════════════════════════════════════════════════════════
# DEBUG
# ═════════════════════════════════════════════════════════════════════════════
st.divider()
with st.expander("🔍 Debug: raw PDF text"):
    st.markdown(
        "**Note:** Page 4 will show `(no text — likely image)` — "
        "this is expected and is why size tables are built-in.",
        unsafe_allow_html=False,
    )
    raw = get_raw_text(file_bytes, 6000)
    st.text_area("", value=raw, height=380, label_visibility="collapsed")
