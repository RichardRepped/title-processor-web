import streamlit as st
import tempfile
import os
import sys
import io
from pathlib import Path
from contextlib import redirect_stdout

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import process_compound_titles as pct

st.set_page_config(
    page_title="Compound Title Processor",
    page_icon="🎵",
    layout="centered"
)

st.markdown("""
<style>
    .stButton > button {
        background-color: #1F4E79;
        color: white;
        font-size: 16px;
        font-weight: bold;
        padding: 12px 32px;
        border-radius: 6px;
        border: none;
        width: 100%;
    }
    .stButton > button:hover { background-color: #2E75B6; }
    .result-box {
        background-color: #E2EFDA;
        border-left: 4px solid #375623;
        padding: 16px;
        border-radius: 4px;
        margin: 16px 0;
    }
    .info-box {
        background-color: #D6E4F0;
        border-left: 4px solid #1F4E79;
        padding: 16px;
        border-radius: 4px;
        margin: 16px 0;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("## 🎵 Compound Title Processor")
st.markdown("**VMG Repertoire Audit — Automated Compound Title Harmonisation**")
st.divider()

st.markdown("""
<div class="info-box">
<strong>How to use:</strong><br>
1. Complete Data → Refresh All in your project workbook and save it<br>
2. Upload the saved workbook below<br>
3. Click <strong>Process Compound Titles</strong><br>
4. Download the processed file and review any flagged rows in the REVIEW column (the last column on the IMPORT sheet)
</div>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "Upload your workbook (.xlsm or .xlsx)",
    type=["xlsm", "xlsx"],
    help="Upload the project workbook after refreshing and saving"
)

if uploaded_file is not None:
    st.success(f"✓ Uploaded: **{uploaded_file.name}**")

    if st.button("⚙️  Process Compound Titles"):

        with st.spinner("Processing — please wait, this takes about 30–60 seconds..."):
            try:
                with tempfile.TemporaryDirectory() as tmpdir:
                    # Save uploaded file to temp dir
                    input_path  = os.path.join(tmpdir, uploaded_file.name)
                    output_name = Path(uploaded_file.name).stem + "_processed.xlsm"
                    output_path = os.path.join(tmpdir, output_name)

                    with open(input_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())

                    # Capture the summary output from process()
                    summary_buf = io.StringIO()
                    with redirect_stdout(summary_buf):
                        pct.process(input_path, output_path)
                    summary = summary_buf.getvalue()

                    # Read output file for download
                    with open(output_path, "rb") as f:
                        output_bytes = f.read()

                # Parse key numbers from summary output
                flagged = 0
                auto    = 0
                written = 0
                total   = 0
                nulls   = 0
                for line in summary.splitlines():
                    try:
                        if "Flagged for review" in line:
                            flagged = int(line.split(":")[1].strip())
                        elif "Auto-resolved" in line:
                            auto = int(line.split(":")[1].strip())
                        elif "Null remaining" in line:
                            nulls = int(line.split(":")[1].strip())
                        elif "Compound titles written" in line:
                            written = int(line.split(":")[1].split("of")[0].strip())
                            total = int(line.split("of")[1].split("rows")[0].strip())
                    except (ValueError, IndexError):
                        pass

                st.markdown(f"""
<div class="result-box">
<strong>✓ Processing complete</strong><br><br>
• <strong>{written:,}</strong> of <strong>{total:,}</strong> rows now have a compound title<br>
• <strong>{auto:,}</strong> rows auto-resolved<br>
• <strong>{flagged:,}</strong> rows flagged for review in the REVIEW column<br>
• Rows still without a title: <strong>{nulls}</strong>
</div>
""", unsafe_allow_html=True)

                if flagged > 0:
                    st.info(
                        f"**{flagged} rows need your attention.** "
                        "In the downloaded file, filter the REVIEW column (the last column on the IMPORT sheet) to see which rows "
                        "need a decision. "
                        "🟠 Orange = different artist credits across DSPs. "
                        "🔵 Blue = Tidal-only, no matching sibling found."
                    )

                st.download_button(
                    label="⬇️  Download processed workbook",
                    data=output_bytes,
                    file_name=output_name,
                    mime="application/vnd.ms-excel.sheet.macroEnabled.12"
                )

            except Exception as e:
                st.error(f"**Processing failed:** {str(e)}")
                st.markdown("""
**Please check:**
- The file is a valid VMG Repertoire Audit workbook
- It contains an IMPORT sheet
- Data has been refreshed and the file saved before uploading
""")

st.divider()
st.caption("VMG Repertoire Audit · Compound Title Processor · Internal use only")
