import streamlit as st
import pandas as pd
import io
import time
from invoice_app import process_pdf_stream

# ==========================================
# PAGE CONFIG & CSS
# ==========================================
st.set_page_config(
    page_title="Invoice Parser AI",
    page_icon="⚡",
    layout="centered",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    /* Dark vibrant theme with Glassmorphism */
    .stApp {
        background: linear-gradient(135deg, #0f172a 0%, #1e1b4b 100%);
        color: #f8fafc;
        font-family: 'Inter', sans-serif;
    }
    
    .css-18e3th9 {
        padding-top: 2rem;
    }

    h1 {
        color: #e2e8f0;
        text-align: center;
        font-weight: 800;
        font-size: 3rem;
        background: -webkit-linear-gradient(45deg, #38bdf8, #818cf8);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.5rem;
    }

    .subtitle {
        text-align: center;
        color: #94a3b8;
        font-size: 1.1rem;
        margin-bottom: 2rem;
    }

    /* Glassmorphism Card for uploader */
    .stFileUploader {
        background: rgba(255, 255, 255, 0.03);
        border-radius: 16px;
        padding: 2rem;
        border: 1px solid rgba(255, 255, 255, 0.05);
        box-shadow: 0 4px 30px rgba(0, 0, 0, 0.1);
        backdrop-filter: blur(5px);
        -webkit-backdrop-filter: blur(5px);
        transition: transform 0.3s ease;
    }
    
    .stFileUploader:hover {
        transform: translateY(-2px);
    }

    /* Primary Button Styling */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #6366f1 0%, #a855f7 100%);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.75rem 2rem;
        font-weight: 600;
        font-size: 1.1rem;
        transition: all 0.3s ease;
        width: 100%;
        margin-top: 1rem;
    }
    
    .stDownloadButton > button:hover {
        opacity: 0.9;
        transform: scale(1.02);
        box-shadow: 0 10px 20px rgba(168, 85, 247, 0.3);
    }
    
    /* Hide Streamlit Branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# ==========================================
# MAIN APP
# ==========================================
st.markdown("<h1>⚡ Invoice Parser AI</h1>", unsafe_allow_html=True)
st.markdown("<p class='subtitle'>Upload your broadcast invoice PDFs to extract structured Excel data instantly.</p>", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "Drag and drop PDF invoices here", 
    accept_multiple_files=True,
    help="Supports Zee, Star, Sun TV, Sony, Eenadu, Polimer, and more."
)

if uploaded_files:
    if st.button("Process Invoices", type="primary", use_container_width=True):
        all_data = []
        
        # Progress Bar & Status
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        total_files = len(uploaded_files)
        
        for i, file in enumerate(uploaded_files):
            if not file.name.lower().endswith('.pdf'):
                st.warning(f"Skipping {file.name}: Only PDF files are supported.")
                continue

            status_text.text(f"Processing: {file.name} ({i+1}/{total_files})")
            
            # Using process_pdf_stream from refactored invoice_app
            try:
                rows = process_pdf_stream(file, filename=file.name)
                all_data.extend(rows)
            except Exception as e:
                st.error(f"Error parsing {file.name}: {e}")
                
            progress_bar.progress((i + 1) / total_files)
            
        time.sleep(0.5)
        status_text.text("Extraction Complete!")
        progress_bar.empty()
        
        if all_data:
            st.success(f"Successfully extracted {len(all_data)} total rows from {total_files} file(s).")
            
            column_order = [
                "Broadcaster Name",
                "Agency Name",
                "Advertiser Name",
                "Channel Name",
                "Billing Period",
                "PO Number",
                "Invoice Number",
                "Invoice Date",
                "Brand",
                "Spot Copy (Caption)",
                "Program",
                "TP",
                "Time Range/Sales Unit",
                "Date",
                "Day",
                "Air Time",
                "LEN (Duration Sec)",
                "Rate (INR)",
                "Calculated Amount (INR)",
            ]
            
            df = pd.DataFrame(all_data)
            for col in column_order:
                if col not in df.columns:
                    df[col] = ""
            df = df[column_order]
            
            # Preview the data
            st.markdown("### Data Preview")
            st.dataframe(df.head(10), use_container_width=True)
            
            # Excel Generation in memory
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Parsed_Invoices')
            excel_data = output.getvalue()
            
            st.download_button(
                label="📥 Download Excel Sheet",
                data=excel_data,
                file_name="extracted_invoices.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No data could be extracted from the uploaded PDFs. Format might be unsupported.")
