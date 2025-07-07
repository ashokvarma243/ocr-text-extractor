import streamlit as st
import tempfile
import os
import zipfile
from datetime import datetime
from streamlit_ocr_engine import StreamlitOCREngine

# Page configuration
st.set_page_config(
    page_title="OCR Text Extractor",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better appearance
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .success-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
    }
    .error-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        color: #721c24;
    }
</style>
""", unsafe_allow_html=True)

def main():
    # Header
    st.markdown('<h1 class="main-header">🔍 OCR Text Extractor</h1>', unsafe_allow_html=True)
    st.markdown("**Extract text from images and PDFs to Excel format while preserving layout**")
    
    # Sidebar for settings
    with st.sidebar:
        st.header("⚙️ Settings")
        confidence = st.slider("Confidence Threshold", 10, 90, 25, 
                              help="Higher values = more accurate text, Lower values = more text detected")
        row_height = st.slider("Row Height Threshold", 10, 50, 20, 
                              help="Adjust for different document layouts")
        
        st.header("📋 Instructions")
        st.markdown("""
        1. **Upload files** using the file uploader
        2. **Adjust settings** if needed (defaults work well)
        3. **Click Process** to extract text
        4. **Download** the Excel files with extracted text
        
        **Supported formats:**
        - Images: PNG, JPG, JPEG, TIFF, BMP
        - Documents: PDF
        """)
        
        st.header("ℹ️ About")
        st.markdown("""
        This tool uses advanced OCR technology to extract text from images and PDFs 
        while preserving the original layout and formatting.
        
        **Features:**
        - Layout preservation
        - Batch processing
        - Excel output format
        - Secure processing (files not stored)
        """)
    
    # Main content area
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📁 File Upload")
        uploaded_files = st.file_uploader(
            "Choose files to process",
            type=['png', 'jpg', 'jpeg', 'pdf', 'tiff', 'bmp'],
            accept_multiple_files=True,
            help="Select one or more images or PDF files"
        )
        
        if uploaded_files:
            st.success(f"✅ {len(uploaded_files)} file(s) selected")
            
            # Show file details
            with st.expander("📄 File Details"):
                for file in uploaded_files:
                    file_size = len(file.read()) / 1024 / 1024  # MB
                    file.seek(0)  # Reset file pointer
                    st.write(f"**{file.name}** - {file_size:.2f} MB")
    
    with col2:
        st.header("🚀 Processing")
        
        if uploaded_files:
            if st.button("🔄 Process Files", type="primary", use_container_width=True):
                process_files(uploaded_files, confidence, row_height)
        else:
            st.info("👆 Upload files to enable processing")

def process_files(uploaded_files, confidence, row_height):
    """Process uploaded files and provide download links"""
    
    # Initialize OCR engine
    ocr_engine = StreamlitOCREngine()
    ocr_engine.confidence_threshold = confidence
    ocr_engine.row_height_threshold = row_height
    
    # Progress tracking
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    processed_files = []
    total_files = len(uploaded_files)
    
    # Create temporary directory for results
    with tempfile.TemporaryDirectory() as temp_dir:
        
        for idx, uploaded_file in enumerate(uploaded_files):
            # Update progress
            progress = (idx + 1) / total_files
            progress_bar.progress(progress)
            status_text.text(f"Processing {uploaded_file.name}... ({idx + 1}/{total_files})")
            
            try:
                if uploaded_file.type == "application/pdf":
                    # Process PDF
                    excel_files, message = ocr_engine.process_pdf(uploaded_file, uploaded_file.name)
                    
                    if excel_files:
                        for excel_path, page_info in excel_files:
                            # Copy to temp directory with proper name
                            result_name = f"{uploaded_file.name}_{page_info}_extracted.xlsx"
                            result_path = os.path.join(temp_dir, result_name)
                            os.rename(excel_path, result_path)
                            processed_files.append((result_path, result_name))
                        
                        st.success(f"✅ {uploaded_file.name}: {message}")
                    else:
                        st.error(f"❌ {uploaded_file.name}: {message}")
                
                else:
                    # Process image
                    excel_path, message = ocr_engine.process_image(uploaded_file, uploaded_file.name)
                    
                    if excel_path:
                        # Copy to temp directory with proper name
                        result_name = f"{os.path.splitext(uploaded_file.name)[0]}_extracted.xlsx"
                        result_path = os.path.join(temp_dir, result_name)
                        os.rename(excel_path, result_path)
                        processed_files.append((result_path, result_name))
                        
                        st.success(f"✅ {uploaded_file.name}: {message}")
                    else:
                        st.error(f"❌ {uploaded_file.name}: {message}")
            
            except Exception as e:
                st.error(f"❌ Error processing {uploaded_file.name}: {str(e)}")
        
        # Complete processing
        progress_bar.progress(1.0)
        status_text.text("Processing complete!")
        
        # Provide download options
        if processed_files:
            st.header("📥 Download Results")
            
            if len(processed_files) == 1:
                # Single file download
                file_path, file_name = processed_files[0]
                with open(file_path, "rb") as file:
                    st.download_button(
                        label=f"📄 Download {file_name}",
                        data=file.read(),
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                # Multiple files - create ZIP
                zip_path = os.path.join(temp_dir, "ocr_results.zip")
                with zipfile.ZipFile(zip_path, 'w') as zipf:
                    for file_path, file_name in processed_files:
                        zipf.write(file_path, file_name)
                
                with open(zip_path, "rb") as zip_file:
                    st.download_button(
                        label=f"📦 Download All Results ({len(processed_files)} files)",
                        data=zip_file.read(),
                        file_name=f"ocr_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                        mime="application/zip"
                    )
            
            # Show individual file download options
            with st.expander("📄 Individual File Downloads"):
                for file_path, file_name in processed_files:
                    with open(file_path, "rb") as file:
                        st.download_button(
                            label=f"Download {file_name}",
                            data=file.read(),
                            file_name=file_name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"download_{file_name}"
                        )
        else:
            st.warning("⚠️ No files were successfully processed. Please check your input files and try again.")

if __name__ == "__main__":
    main()
