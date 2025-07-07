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
        # For Streamlit Cloud deployment, Tesseract path is handled automatically
        self.confidence_threshold = 25
        self.row_height_threshold = 20
        self.column_gap_threshold = 40
    
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
            
            # Enhanced contrast
            clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8,8))
            enhanced = clahe.apply(gray)
            
            # Denoising
            denoised = cv2.fastNlMeansDenoising(enhanced, h=10, templateWindowSize=7, searchWindowSize=21)
            
            # Adaptive thresholding
            adaptive_thresh = cv2.adaptiveThreshold(
                denoised, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2
            )
            
            return Image.fromarray(adaptive_thresh)
            
        except Exception as e:
            print(f"Error in preprocessing: {str(e)}")
            return image
    
    def extract_text_with_coordinates(self, image):
        """Extract text with coordinate information"""
        try:
            processed_image = self.preprocess_image_advanced(image)
            
            # Get detailed OCR data
            ocr_data = pytesseract.image_to_data(
                processed_image, 
                output_type=pytesseract.Output.DICT,
                config='--psm 6 -c preserve_interword_spaces=1'
            )
            
            # Extract text elements with positions
            text_elements = []
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
            
            return text_elements
            
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
            '@': '', '¥': 'Y', '£': 'E', '€': 'E', '§': 'S'
        }
        
        for old, new in replacements.items():
            text = text.replace(old, new)
        
        # Remove non-printable characters
        text = re.sub(r'[^\w\s\.\,\!\?\:\;\-\(\)\[\]\{\}\'\"•#$%&*+=<>/\\|`~]', '', text)
        
        return text.strip() if len(text.strip()) > 1 else ""
    
    def create_visual_layout(self, text_elements):
        """Create visual layout preserving document structure"""
        if not text_elements:
            return []
        
        # Sort by vertical position
        text_elements.sort(key=lambda x: x['top'])
        
        # Group into rows
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
        """Create Excel file with preserved layout"""
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
                
                col = 1
                for element in row_elements:
                    cell = ws.cell(row=excel_row, column=col)
                    cell.value = element['text']
                    cell.alignment = Alignment(wrap_text=True, vertical='top')
                    
                    # Apply smart formatting
                    if element['text'].isupper() and len(element['text']) < 50:
                        cell.font = Font(bold=True, size=12)
                    elif element['text'].startswith('•') or element['text'].startswith('-'):
                        cell.alignment = Alignment(indent=1, wrap_text=True, vertical='top')
                    
                    col += 1
                
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
    
    def process_image(self, image_file, filename):
        """Process a single image and return Excel file path"""
        try:
            # Open image
            image = Image.open(image_file)
            
            # Extract text with coordinates
            text_elements = self.extract_text_with_coordinates(image)
            
            if not text_elements:
                return None, "No text found in image"
            
            # Create visual layout
            rows = self.create_visual_layout(text_elements)
            
            # Create Excel file
            excel_path = self.create_excel_output(rows, filename)
            
            return excel_path, f"Successfully extracted {len(text_elements)} text elements"
            
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
                
                # Convert to images
                pages = convert_from_path(temp_pdf_path, dpi=300)
                
                excel_files = []
                for page_num, page in enumerate(pages, 1):
                    # Process each page
                    text_elements = self.extract_text_with_coordinates(page)
                    
                    if text_elements:
                        rows = self.create_visual_layout(text_elements)
                        excel_path = self.create_excel_output(rows, f"{filename}_page_{page_num}")
                        if excel_path:
                            excel_files.append((excel_path, f"Page {page_num}"))
                
                return excel_files, f"Successfully processed {len(pages)} pages"
                
        except Exception as e:
            return [], f"Error processing PDF: {str(e)}"
