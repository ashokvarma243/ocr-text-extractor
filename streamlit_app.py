import streamlit as st
import tempfile
import os
import zipfile
import subprocess
import sys
from datetime import datetime
from PIL import Image
import pytesseract

# Setup for Streamlit Cloud deployment
def setup_tesseract_cloud():
    """Configure Tesseract for Streamlit Cloud deployment"""
    try:
        # Check if tesseract is available
        result = subprocess.run(['tesseract', '--version'], 
                              capture_output=True, text=True, check=True)
        print(f"Tesseract found: {result.stdout.split()[1]}")
        
        # Set environment variables for Streamlit Cloud
        os.environ['TESSDATA_PREFIX'] = '/usr/share/tesseract-ocr/4.00/tessdata'
        
        # Configure pytesseract
        if os.path.exists('/usr/bin/tesseract'):
            pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'
        
        return True
    except Exception as e:
        print(f"Tesseract setup error: {e}")
        # Try alternative paths
        alternative_paths = [
            '/usr/local/bin/tesseract',
            '/opt/homebrew/bin/tesseract'
        ]
        for path in alternative_paths:
            if os.path.exists(path):
                pytesseract.pytesseract.tesseract_cmd = path
                return True
        return False

# Call setup before importing OCR engine
setup_success = setup_tesseract_cloud()

# Import OCR engine after Tesseract setup
try:
    from streamlit_ocr_engine import StreamlitOCREngine
except ImportError as e:
    st.error(f"Failed to import OCR engine: {e}")
    st.error("Please ensure streamlit_ocr_engine.py is in your repository")
    st.stop()

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
    .debug-info {
        background-color: #f8f9fa;
        border: 1px solid #dee2e6;
        border-radius: 0.375rem;
        padding: 1rem;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

def main():
    # Check Tesseract setup
    if not setup_success:
        st.error("⚠️ OCR engine configuration failed. Tesseract OCR is not properly installed.")
        st.info("This is likely a deployment issue. Please contact support or try again later.")
        st.stop()
    
    # Header
    st.markdown('<h1 class="main-header">🔍 OCR Text Extractor</h1>', unsafe_allow_html=True)
    st.markdown("**Extract text from images and PDFs to Excel format while preserving layout**")
    
    # Sidebar for settings
    with st.sidebar:
        st.header("⚙️ Settings")
        confidence = st.slider("Confidence Threshold", 5, 90, 15, 
                              help="Lower values detect more text (try 10-20 for difficult images)")
        row_height = st.slider("Row Height Threshold", 10, 50, 25, 
                              help="Adjust for different document layouts")
        
        # Advanced spacing controls
        st.subheader("📏 Spacing Controls")
        word_spacing = st.slider("Word Spacing Threshold", 10, 50, 30, 
                                help="Pixels - smaller values group words more tightly")
        column_break = st.slider("Column Break Threshold", 50, 150, 80, 
                                help="Pixels - larger gaps indicate new columns")
        
        # Debug mode
        debug_mode = st.checkbox("Enable Debug Mode", 
                                help="Show detailed processing information for troubleshooting")
        
        st.header("📋 Instructions")
        st.markdown("""
        1. **Upload files** using the file uploader
        2. **Adjust settings** if needed (defaults work well)
        3. **Click Process** to extract text
        4. **Download** the Excel files with extracted text
        
        **Supported formats:**
        - Images: PNG, JPG, JPEG, TIFF, BMP
        - Documents: PDF
        
        **Tips for better results:**
        - Use high-resolution images
        - Ensure text is clearly visible
        - Lower confidence threshold for difficult images
        - Adjust spacing controls for better layout preservation
        """)
        
        st.header("ℹ️ About")
        st.markdown("""
        This tool uses advanced OCR technology to extract text from images and PDFs 
        while preserving the original layout and formatting.
        
        **Features:**
        - Smart spacing-based layout preservation
        - Batch processing
        - Excel output format
        - Secure processing (files not stored)
        - Multiple OCR configurations for best results
        - Intelligent column detection
        """)
    
    # Main content area
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📁 File Upload")
        uploaded_files = st.file_uploader(
            "Choose files to process",
            type=['png', 'jpg', 'jpeg', 'pdf', 'tiff', 'bmp'],
            accept_multiple_files=True,
            help="Select one or more images or PDF files (max 200MB per file)"
        )
        
        if uploaded_files:
            st.success(f"✅ {len(uploaded_files)} file(s) selected")
            
            # Show file details
            with st.expander("📄 File Details"):
                for file in uploaded_files:
                    file_size = len(file.read()) / 1024 / 1024  # MB
                    file.seek(0)  # Reset file pointer
                    st.write(f"**{file.name}** - {file_size:.2f} MB - Type: {file.type}")
    
    with col2:
        st.header("🚀 Processing")
        
        if uploaded_files:
            if st.button("🔄 Process Files", type="primary", use_container_width=True):
                process_files(uploaded_files, confidence, row_height, word_spacing, column_break, debug_mode)
        else:
            st.info("👆 Upload files to enable processing")
        
        # System info
        with st.expander("🔧 System Information"):
            st.write("**Tesseract Status:**", "✅ Available" if setup_success else "❌ Not Available")
            try:
                result = subprocess.run(['tesseract', '--version'], 
                                      capture_output=True, text=True, check=True)
                version = result.stdout.split('\n')[0]
                st.write("**Version:**", version)
            except:
                st.write("**Version:**", "Unable to detect")

def process_files(uploaded_files, confidence, row_height, word_spacing, column_break, debug_mode):
    """Process uploaded files and provide download links"""
    
    # Initialize OCR engine
    try:
        ocr_engine = StreamlitOCREngine()
        ocr_engine.confidence_threshold = confidence
        ocr_engine.row_height_threshold = row_height
        ocr_engine.word_spacing_threshold = word_spacing
        ocr_engine.column_break_threshold = column_break
    except Exception as e:
        st.error(f"Failed to initialize OCR engine: {e}")
        return
    
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
            
            if debug_mode:
                st.markdown(f"### 🔍 Debug Info for {uploaded_file.name}")
                debug_container = st.container()
            
            try:
                if uploaded_file.type == "application/pdf":
                    # Process PDF
                    if debug_mode:
                        with debug_container:
                            st.write("📄 Processing PDF file...")
                    
                    excel_files, message = ocr_engine.process_pdf(uploaded_file, uploaded_file.name)
                    
                    if excel_files:
                        for excel_path, page_info in excel_files:
                            # Copy to temp directory with proper name
                            result_name = f"{uploaded_file.name}_{page_info}_extracted.xlsx"
                            result_path = os.path.join(temp_dir, result_name)
                            os.rename(excel_path, result_path)
                            processed_files.append((result_path, result_name))
                        
                        st.success(f"✅ {uploaded_file.name}: {message}")
                        if debug_mode:
                            with debug_container:
                                st.write(f"✅ Created {len(excel_files)} Excel file(s)")
                    else:
                        st.error(f"❌ {uploaded_file.name}: {message}")
                        if debug_mode:
                            with debug_container:
                                st.write(f"❌ PDF processing failed: {message}")
                
                else:
                    # Process image with enhanced debugging
                    if debug_mode:
                        with debug_container:
                            image = Image.open(uploaded_file)
                            st.write(f"🖼️ Image size: {image.size}")
                            st.write(f"🎨 Image mode: {image.mode}")
                            st.image(image, caption="Original Image", width=300)
                            
                            # Test basic OCR first
                            try:
                                basic_text = pytesseract.image_to_string(image)
                                st.write(f"📝 Basic OCR result length: {len(basic_text)} characters")
                                if basic_text.strip():
                                    st.write(f"📄 Sample text: {basic_text[:200]}...")
                                else:
                                    st.warning("⚠️ Basic OCR returned no text")
                            except Exception as e:
                                st.error(f"❌ Basic OCR failed: {e}")
                    
                    # Process with OCR engine
                    excel_path, message = ocr_engine.process_image(uploaded_file, uploaded_file.name)
                    
                    if excel_path:
                        # Copy to temp directory with proper name
                        result_name = f"{os.path.splitext(uploaded_file.name)[0]}_extracted.xlsx"
                        result_path = os.path.join(temp_dir, result_name)
                        os.rename(excel_path, result_path)
                        processed_files.append((result_path, result_name))
                        
                        st.success(f"✅ {uploaded_file.name}: {message}")
                        if debug_mode:
                            with debug_container:
                                st.write(f"✅ Engine processing successful")
                    else:
                        st.error(f"❌ {uploaded_file.name}: {message}")
                        if debug_mode:
                            with debug_container:
                                st.write(f"❌ Engine processing failed: {message}")
                                st.write("💡 Try lowering the confidence threshold or adjusting spacing controls")
            
            except Exception as e:
                st.error(f"❌ Error processing {uploaded_file.name}: {str(e)}")
                if debug_mode:
                    with debug_container:
                        st.write(f"❌ Exception details: {str(e)}")
        
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
            st.info("💡 **Troubleshooting tips:**\n"
                   "- Try lowering the confidence threshold to 10-15\n"
                   "- Ensure images have clear, readable text\n"
                   "- Check that images are not too blurry or low resolution\n"
                   "- Adjust spacing controls for better layout detection\n"
                   "- Enable debug mode to see detailed processing information")

if __name__ == "__main__":
    main()
