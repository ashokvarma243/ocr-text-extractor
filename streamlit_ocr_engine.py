import os
import cv2
import numpy as np
from datetime import datetime
from PIL import Image, ImageEnhance, ImageFilter
import pandas as pd
import pytesseract
from pdf2image import convert_from_path
import re
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
import tempfile
import shutil

class StreamlitOCREngine:
    def __init__(self):
        # Setup cloud environment
        self.setup_cloud_environment()
        self.confidence_threshold = 15  # Lower default for cloud
        self.row_height_threshold = 20
        self.column_gap_threshold = 40
    
    def setup_cloud_environment(self):
        """Setup OCR for cloud deployment"""
        # Set Tesseract path for different environments
        tesseract_paths = [
            '/usr/bin/tesseract',
            '/usr/local/bin/tesseract',
            '/opt/homebrew/bin/tesseract',
            r'C:\Program Files\Tesseract-OCR\tesseract.exe'  # Windows fallback
        ]
        
        for path in tesseract_paths:
            if os.path.exists(path):
                pytesseract.pytesseract.tesseract_cmd = path
                break
        
        # Set tessdata path
        tessdata_paths = [
            '/usr/share/tesseract-ocr/4.00/tessdata',
            '/usr/share/tesseract-ocr/tessdata',
            '/usr/local/share/tessdata',
            '/opt/homebrew/share/tessdata'
        ]
        
        for path in tessdata_paths:
            if os.path.exists(path):
                os.environ['TESSDATA_PREFIX'] = path
                break
    
    def preprocess_image_advanced(self, image):
        """Advanced image preprocessing for better OCR"""
        try:
            # Convert PIL to OpenCV if needed
            if isinstance(image, Image.Image):
                opencv_image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
            else:
                opencv_image = image
            
            # Convert to grayscale
            gray = cv2.cvtColor(opencv_image, cv2.COLOR_BGR2GRAY)
            
            # Try multiple preprocessing approaches
            processed_images = []
            
            # 1. Original grayscale
            processed_images.append(('original', gray))
            
            # 2. Enhanced contrast
            clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8,8))
            enhanced = clahe.apply(gray)
            processed_images.append(('enhanced', enhanced))
            
            # 3. Denoising
            denoised = cv2.fastNlMeansDenoising(enhanced, h=10, templateWindowSize=7, searchWindowSize=21)
            processed_images.append(('denoised', denoised))
            
            # 4. Adaptive thresholding
            adaptive_thresh = cv2.adaptiveThreshold(
                denoised, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2
            )
            processed_images.append(('adaptive', adaptive_thresh))
            
            # 5. OTSU thresholding
            _, otsu_thresh = cv2.threshold(denoised, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            processed_images.append(('otsu', otsu_thresh))
            
            # Return the best preprocessed image (adaptive threshold usually works well)
            return Image.fromarray(adaptive_thresh)
            
        except Exception as e:
            print(f"Error in preprocessing: {str(e)}")
            return image
    
    def extract_text_with_coordinates(self, image):
        """Enhanced text extraction with multiple OCR configurations"""
        try:
            processed_image = self.preprocess_image_advanced(image)
            
            # Try multiple OCR configurations for better results
            ocr_configs = [
                '--psm 6 -c preserve_interword_spaces=1',
                '--psm 4 -c preserve_interword_spaces=1',
                '--psm 3 -c preserve_interword_spaces=1',
                '--psm 11 -c preserve_interword_spaces=1',
                '--psm 12 -c preserve_interword_spaces=1',
                '--psm 6',
                '--psm 4',
                '--psm 3'
            ]
            
            best_result = []
            best_score = 0
            
            for config in ocr_configs:
                try:
                    # Get detailed OCR data
                    ocr_data = pytesseract.image_to_data(
                        processed_image, 
                        output_type=pytesseract.Output.DICT,
                        config=config
                    )
                    
                    # Extract text elements
                    text_elements = []
                    total_confidence = 0
                    valid_elements = 0
                    
                    for i in range(len(ocr_data['text'])):
                        text = ocr_data['text'][i].strip()
                        confidence = int(ocr_data['conf'][i]) if ocr_data['conf'][i] != '-1' else 0
                        
                        if text and len(text) > 1 and confidence > self.confidence_threshold:
                            cleaned_text = self.clean_ocr_text(text)
                            if cleaned_text:
                                text_elements.append({
                                    'text': cleaned_text,
                                    'left': ocr_data['left'][i],
                                    'top': ocr_data['top'][i],
                                    'width': ocr_data['width'][i],
                                    'height': ocr_data['height'][i],
                                    'confidence': confidence
                                })
                                total_confidence += confidence
                                valid_elements += 1
                    
                    # Calculate score for this configuration
                    if valid_elements > 0:
                        avg_confidence = total_confidence / valid_elements
                        text_length = sum(len(elem['text']) for elem in text_elements)
                        score = (len(text_elements) * 0.4) + (avg_confidence * 0.3) + (text_length * 0.3)
                        
                        if score > best_score:
                            best_score = score
                            best_result = text_elements
                            
                except Exception as e:
                    print(f"OCR config {config} failed: {e}")
                    continue
            
            return best_result
            
        except Exception as e:
            print(f"Error extracting text: {str(e)}")
            return []
    
    def clean_ocr_text(self, text):
        """Clean and normalize OCR text"""
        if not text:
            return ""
        
        # Remove excessive whitespace
        text = re.sub(r'\s+', ' ', text.strip())
        
        # Fix common OCR errors
        replacements = {
            '°': '•', '¢': '•', '{f': '•', '|': 'I',
            '@': '', '¥': 'Y', '£': 'E', '€': 'E', '§': 'S',
            '0': 'O', '5': 'S', '1': 'I'  # Common character confusions
        }
        
        # Apply replacements selectively
        for old, new in replacements.items():
            if old in ['0', '5', '1']:  # Only replace in uppercase words
                words = text.split()
                for i, word in enumerate(words):
                    if word.isupper():
                        words[i] = word.replace(old, new)
                text = ' '.join(words)
            else:
                text = text.replace(old, new)
        
        # Remove non-printable characters except common punctuation
        text = re.sub(r'[^\w\s\.\,\!\?\:\;\-\(\)\[\]\{\}\'\"•#$%&*+=<>/\\|`~]', '', text)
        
        return text.strip() if len(text.strip()) > 1 else ""
    
    def create_visual_layout(self, text_elements):
        """Create visual layout preserving document structure"""
        if not text_elements:
            return []
        
        # Sort by vertical position
        text_elements.sort(key=lambda x: x['top'])
        
        # Group into rows with dynamic threshold
        rows = []
        current_row = []
        current_top = None
        
        for element in text_elements:
            if current_top is None or abs(element['top'] - current_top) <= self.row_height_threshold:
                current_row.append(element)
                current_top = element['top'] if current_top is None else current_top
            else:
                if current_row:
                    current_row.sort(key=lambda x: x['left'])
                    rows.append(current_row)
                current_row = [element]
                current_top = element['top']
        
        if current_row:
            current_row.sort(key=lambda x: x['left'])
            rows.append(current_row)
        
        return rows
    
    def create_excel_output(self, rows, filename):
        """Create Excel file with preserved layout and enhanced formatting"""
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "OCR_Results"
            
            if not rows:
                return None
            
            # Process each row
            excel_row = 1
            for row_elements in rows:
                if not row_elements:
                    excel_row += 1
                    continue
                
                # Determine column positions based on horizontal spacing
                col_positions = self.calculate_column_positions(row_elements)
                
                for element, col_pos in zip(row_elements, col_positions):
                    cell = ws.cell(row=excel_row, column=col_pos)
                    cell.value = element['text']
                    cell.alignment = Alignment(wrap_text=True, vertical='top')
                    
                    # Apply smart formatting based on content
                    self.apply_smart_formatting(cell, element)
                
                excel_row += 1
            
            # Auto-adjust column widths
            for column in ws.columns:
                max_length = 0
                column_letter = column[0].column_letter
                
                for cell in column:
                    try:
                        if cell.value and len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                
                adjusted_width = min(max_length + 2, 50)
                if adjusted_width > 10:
                    ws.column_dimensions[column_letter].width = adjusted_width
            
            # Save to temporary file
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
            wb.save(temp_file.name)
            return temp_file.name
            
        except Exception as e:
            print(f"Error creating Excel: {str(e)}")
            return None
    
    def calculate_column_positions(self, row_elements):
        """Calculate appropriate column positions for elements"""
        if len(row_elements) == 1:
            return [1]
        
        positions = []
        current_col = 1
        
        for i, element in enumerate(row_elements):
            if i == 0:
                positions.append(current_col)
            else:
                # Calculate gap from previous element
                prev_element = row_elements[i-1]
                gap = element['left'] - (prev_element['left'] + prev_element['width'])
                
                # If significant gap, move to next column
                if gap > self.column_gap_threshold:
                    current_col += 1
                
                positions.append(current_col)
        
        return positions
    
    def apply_smart_formatting(self, cell, element):
        """Apply smart formatting based on text content and properties"""
        text = element['text']
        
        # Header detection (uppercase, short text, high confidence)
        if (text.isupper() and len(text) < 50 and element['confidence'] > 70):
            cell.font = Font(bold=True, size=12, color="1F4E79")
            cell.fill = PatternFill(start_color="E6E6FA", end_color="E6E6FA", fill_type="solid")
        
        # Bullet point detection
        elif text.startswith('•') or text.startswith('-') or text.startswith('*'):
            cell.font = Font(size=10)
            cell.alignment = Alignment(indent=1, wrap_text=True, vertical='top')
        
        # Number/step detection
        elif re.match(r'^\d+\.', text):
            cell.font = Font(bold=True, size=11, color="2E75B6")
        
        # Default formatting
        else:
            cell.font = Font(size=10)
        
        # Always enable text wrapping
        cell.alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')
    
    def process_image(self, image_file, filename):
        """Process a single image and return Excel file path"""
        try:
            # Open image
            image = Image.open(image_file)
            
            # Extract text with coordinates
            text_elements = self.extract_text_with_coordinates(image)
            
            if not text_elements:
                return None, "No text found in image. Try lowering the confidence threshold or check image quality."
            
            # Create visual layout
            rows = self.create_visual_layout(text_elements)
            
            # Create Excel file
            excel_path = self.create_excel_output(rows, filename)
            
            if excel_path:
                return excel_path, f"Successfully extracted {len(text_elements)} text elements from {len(rows)} rows"
            else:
                return None, "Failed to create Excel file"
            
        except Exception as e:
            return None, f"Error processing image: {str(e)}"
    
    def process_pdf(self, pdf_file, filename):
        """Process PDF and return Excel file paths"""
        try:
            # Convert PDF to images
            with tempfile.TemporaryDirectory() as temp_dir:
                # Save uploaded file temporarily
                temp_pdf_path = os.path.join(temp_dir, "temp.pdf")
                with open(temp_pdf_path, "wb") as f:
                    f.write(pdf_file.read())
                
                # Convert to images with higher DPI for better OCR
                pages = convert_from_path(temp_pdf_path, dpi=300, fmt='png')
                
                excel_files = []
                total_elements = 0
                
                for page_num, page in enumerate(pages, 1):
                    # Process each page
                    text_elements = self.extract_text_with_coordinates(page)
                    
                    if text_elements:
                        rows = self.create_visual_layout(text_elements)
                        excel_path = self.create_excel_output(rows, f"{filename}_page_{page_num}")
                        if excel_path:
                            excel_files.append((excel_path, f"Page {page_num}"))
                            total_elements += len(text_elements)
                
                if excel_files:
                    return excel_files, f"Successfully processed {len(pages)} pages with {total_elements} total text elements"
                else:
                    return [], f"No text found in any of the {len(pages)} pages. Check PDF quality or try lowering confidence threshold."
                
        except Exception as e:
            return [], f"Error processing PDF: {str(e)}"
