import streamlit as st
import sys
from pathlib import Path
import tempfile, shutil

sys.path.insert(0, str(Path(__file__).parent))
from build_excel import build as build_excel_file
from parser_pdf import parse_factory_pdf

st.set_page_config(
    page_title="Hyginox Quote Converter",
    page_icon="🏗️",
    layout="centered"
)

# ── Styling ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { background-color: #f8f9fa; }
    .title  { text-align:center; color:#8B6914; font-size:2.8em; font-weight:900; margin-bottom:0; }
    .sub    { text-align:center; color:#555; font-size:1.05em; margin-top:0; }
    .badge  { text-align:center; color:#888; font-size:0.85em; margin-bottom:20px; }
    .result { background:#d4edda; border-radius:8px; padding:14px; margin-top:16px; }
    .stButton>button {
        background-color:#2E7D32; color:white; font-weight:700;
        border:none; border-radius:6px; padding:10px 28px;
        font-size:1.05em; width:100%;
    }
    .stButton>button:hover { background-color:#1b5e20; }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown('<p class="title">HYGINOX</p>', unsafe_allow_html=True)
st.markdown('<p class="sub">Quote Converter — Factory PDF → Hyginox Excel</p>', unsafe_allow_html=True)
st.markdown('<p class="badge">Solana Eco Homes format &nbsp;•&nbsp; 10% markup &nbsp;•&nbsp; Editable Excel download</p>', unsafe_allow_html=True)
st.divider()

# ── Upload ────────────────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "📄 Upload Hygiene Kitchens Factory PDF",
    type=["pdf"],
    help="Upload the factory quote PDF — text or scanned"
)

if uploaded:
    st.success(f"✅ Loaded: **{uploaded.name}**")

    if st.button("Convert → Download Excel"):
        with st.spinner("Reading PDF and building Excel..."):
            try:
                # Save upload to temp file
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                    tmp.write(uploaded.read())
                    tmp_path = tmp.name

                # Parse + build
                import re
                quote_no, date_str, client, project, items = parse_factory_pdf(tmp_path)
                logo_path = Path(__file__).parent / "hyginox_logo.png"
                out_path  = tmp_path.replace(".pdf", "_HYGINOX.xlsx")

                out_file, grand = build_excel_file(
                    items, quote_no, date_str, client, project,
                    logo_path=str(logo_path)
                )

                num_part  = re.sub(r'[^\d]', '', quote_no)[:8]
                sol_quote = f'SEH{num_part}'

                # Read output for download
                with open(out_file, "rb") as f:
                    excel_bytes = f.read()

                # Results
                st.markdown(f"""
                <div class="result">
                    <b>✅ Conversion complete!</b><br>
                    Quote: <b>{sol_quote}</b> &nbsp;|&nbsp;
                    Date: <b>{date_str}</b> &nbsp;|&nbsp;
                    Items: <b>{len(items)}</b> &nbsp;|&nbsp;
                    Grand Total: <b>₹{grand:,.2f}</b>
                </div>
                """, unsafe_allow_html=True)

                st.download_button(
                    label="📥 Download Excel",
                    data=excel_bytes,
                    file_name=f"{Path(uploaded.name).stem}_HYGINOX.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"❌ Error: {str(e)}")

st.divider()
st.caption("Solana Eco Homes · HYGINOX Stainless Steel Modular Kitchen")
