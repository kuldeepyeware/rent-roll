import streamlit as st
import os
import tempfile
import subprocess
import json
import zipfile
from pathlib import Path
import shutil

# Basic page configuration
st.set_page_config(
    page_title="Rent Roll Standardizer",
    page_icon="🏢",
    layout="wide"
)

st.title("Rent Roll Standardizer")
st.markdown("""
Upload a Rent Roll spreadsheet (`.xlsx`, `.xls`, `.csv`) and get a standardized JSON and Excel file.
""")

# Main content area
uploaded_file = st.file_uploader("Upload Rent Roll", type=['csv', 'xlsx', 'xls'])

if st.button("Standardize", type="primary"):
    if not uploaded_file:
        st.warning("Please upload a file first.")
    else:
        with st.spinner("Standardizing Rent Roll... This may take a minute."):
            # Create a temporary directory to work in
            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir_path = Path(tmpdir)
                
                # Save the uploaded file
                input_file_path = tmpdir_path / uploaded_file.name
                with open(input_file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                
                output_dir_path = tmpdir_path / "output"
                
                # Set up environment variables for the subprocess
                env = os.environ.copy()
                
                # Automatically pull from st.secrets if running on Streamlit Community Cloud
                if "OPENROUTER_API_KEY" in st.secrets:
                    env["OPENROUTER_API_KEY"] = st.secrets["OPENROUTER_API_KEY"]
                if "OPENROUTER_MODEL" in st.secrets:
                    env["OPENROUTER_MODEL"] = st.secrets["OPENROUTER_MODEL"]
                
                # Run the standardizer script as a subprocess
                script_path = Path(__file__).parent / "rent_roll_standardizer.py"
                
                import sys
                try:
                    result = subprocess.run(
                        [
                            sys.executable, 
                            str(script_path), 
                            "--input", str(input_file_path),
                            "--output-dir", str(output_dir_path)
                        ],
                        env=env,
                        capture_output=True,
                        text=True,
                        check=True
                    )
                    
                    st.success("Standardization Complete!")
                    
                    # Find the generated output folder
                    if output_dir_path.exists():
                        output_folders = list(output_dir_path.glob("*"))
                        if output_folders:
                            run_folder = output_folders[0]
                            
                            # Look for standardized_rent_roll.xlsx
                            excel_file = run_folder / "standardized_rent_roll.xlsx"
                            
                            if excel_file.exists():
                                with open(excel_file, "rb") as f:
                                    st.session_state['excel_data'] = f.read()
                                
                            # Create a zip of the entire run folder for debug artifacts
                            zip_path = tmpdir_path / "artifacts.zip"
                            shutil.make_archive(str(zip_path.with_suffix('')), 'zip', run_folder)
                            
                            with open(zip_path, "rb") as f:
                                st.session_state['zip_data'] = f.read()
                                
                except subprocess.CalledProcessError as e:
                    st.error("An error occurred during standardization.")
                    with st.expander("Error Logs"):
                        st.text(e.stderr)

if 'excel_data' in st.session_state and 'zip_data' in st.session_state:
    st.markdown("---")
    st.subheader("Download Results")
    col1, col2 = st.columns(2)
    
    col1.download_button(
        label="📥 Download Excel",
        data=st.session_state['excel_data'],
        file_name="standardized_rent_roll.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    col2.download_button(
        label="📦 Download all artifacts (ZIP)",
        data=st.session_state['zip_data'],
        file_name="standardizer_artifacts.zip",
        mime="application/zip"
    )
