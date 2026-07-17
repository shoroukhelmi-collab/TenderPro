"""TenderPro Streamlit application."""

from __future__ import annotations

import streamlit as st

from modules.comparison_engine import enrich_comparison
from modules.dashboard import render_dashboard
from modules.excel_reader import create_sample_files, read_all_excels, save_uploaded_files
from modules.export import export_comparison

st.set_page_config(page_title="TenderPro", page_icon="📊", layout="wide")

st.markdown(
    """
    <style>
    .main .block-container {padding-top: 2rem;}
    .hero {
        border-radius: 24px;
        padding: 32px;
        background: linear-gradient(135deg, #0f172a 0%, #1d4ed8 58%, #38bdf8 100%);
        color: white;
        margin-bottom: 24px;
        box-shadow: 0 18px 45px rgba(15, 23, 42, 0.18);
    }
    .hero h1 {font-size: 3.2rem; margin: 0; letter-spacing: -0.06em;}
    .hero p {font-size: 1.1rem; opacity: 0.9; margin-top: 10px; max-width: 900px;}
    .step-card {
        border: 1px solid #e2e8f0;
        border-radius: 18px;
        padding: 18px;
        background: #ffffff;
        box-shadow: 0 8px 22px rgba(15, 23, 42, 0.06);
        min-height: 118px;
    }
    .step-card h3 {margin-top: 0; color: #0f172a;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="hero">
      <h1>TenderPro</h1>
      <p>Commercial bid comparison for EPC/MEP procurement teams. Upload vendor BOQs, normalize inconsistent Excel formats, and identify the lowest offers, missing prices, vendor ranking, and variance.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

steps = st.columns(3)
steps[0].markdown("<div class='step-card'><h3>1. New Project</h3><p>Start a comparison workspace for your tender package.</p></div>", unsafe_allow_html=True)
steps[1].markdown("<div class='step-card'><h3>2. Upload Vendor Files</h3><p>Add multiple Excel BOQs from suppliers with different column names.</p></div>", unsafe_allow_html=True)
steps[2].markdown("<div class='step-card'><h3>3. Generate Comparison</h3><p>Produce ranking, variance, and exportable side-by-side analysis.</p></div>", unsafe_allow_html=True)

st.divider()

with st.sidebar:
    st.header("Project Controls")
    project_name = st.text_input("New Project", value="MEP Tender Comparison")
    uploaded_files = st.file_uploader(
        "Upload Vendor Files",
        type=["xlsx"],
        accept_multiple_files=True,
        help="Upload one or more Excel BOQs. If none are available, TenderPro creates Enext and Valtria samples.",
    )
    if uploaded_files:
        saved = save_uploaded_files(uploaded_files)
        st.success(f"Saved {len(saved)} uploaded file(s).")
    if st.button("Create Sample Enext & Valtria Files"):
        sample_paths = create_sample_files()
        st.success(f"Sample files ready: {', '.join(path.name for path in sample_paths)}")

st.subheader(project_name)

if st.button("Generate Comparison", type="primary"):
    master = read_all_excels()
    comparison = enrich_comparison(master)
    st.session_state["comparison"] = comparison
    st.success("Comparison generated successfully.")

if "comparison" not in st.session_state:
    create_sample_files()
    st.info("No comparison generated yet. Sample Enext and Valtria files are available in uploads/. Click Generate Comparison to begin.")
else:
    comparison = st.session_state["comparison"]
    render_dashboard(comparison)
    export_path = export_comparison(comparison)
    with open(export_path, "rb") as workbook:
        st.download_button(
            "Download Excel Comparison",
            data=workbook,
            file_name=export_path.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
