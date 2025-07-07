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
        self.setup_cloud_environment()
        self.confidence_threshold = 10  # Even lower for better detection
        self.row_height_threshold = 25
        self.column_gap_threshold = 50
    
    def setup_cloud_environment(self):
        """Setup OCR for cloud deployment"""
        tesseract_paths = [
            '/usr/bin/tesseract',
            '/usr/local/bin/tesseract',
            '/opt/homebrew/bin/tesseract',
            r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        ]
        
        for path in tesseract_paths:
            if os.path.exists(path):
                pytesseract.pytesseract.tesseract_cmd = path
                break
        
        tessdata_paths = [
            '/usr/share/tesseract-ocr/4.00/tessdata',
            '/usr/share/tesseract-ocr/tessdata',
            '/usr/local/share/tessdata'
        ]
        
        for path in tessdata_paths:
            if os.path.exists(path):
                os.environ['TESSDATA_PREFIX'] = path
                break
    
    def preprocess_image_for_forms(self, image):
        """Specialized preprocessing for form documents"""
        try:
            # Convert PIL to OpenCV
            if isinstance(image, Image.Image):
                opencv_image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
            else:
                opencv_image = image
            
            # Convert to grayscale
            gray = cv2.cvtColor(opencv_image, cv2.COLOR_BGR2GRAY)
            
            # For forms, try minimal preprocessing to preserve text clarity
            # 1. Slight denoising
            denoised = cv2.fastNlMeansDenoising(gray, h=3, templateWindowSize=7, searchWindowSize=21)
            
            # 2. Enhance contrast slightly
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
            enhanced = clahe.apply(denoised)
            
            # 3. For clean documents like forms, often original or lightly enhanced works best
            return Image.fromarray(enhanced)
            
        except Exception as e:
            print(f"Error in preprocessing: {str(e)}")
            return image
    
    def extract_text_comprehensive(self, image):
        """Comprehensive text extraction with multiple methods"""
        try:
            # Try both original and preprocessed versions
            original_pil = image
            processed_pil = self.preprocess_image_for_forms(image)
            
            all_results = []
            
            # Different OCR configurations optimized for forms
            configs = [
                ('--psm 1 -c preserve_interword_spaces=1', 'Auto page segmentation with OSD'),
                ('--psm 3 -c preserve_interword_spaces=1', 'Fully automatic page segmentation'),
                ('--psm 4 -c preserve_interword_spaces=1', 'Single column of text'),
                ('--psm 6 -c preserve_interword_spaces=1', 'Uniform block of text'),
                ('--psm 11 -c preserve_interword_spaces=1', 'Sparse text'),
                ('--psm 12 -c preserve_interword_spaces=1', 'Sparse text with OSD'),
            ]
            
            # Try each configuration with both image versions
            for config, description in configs:
                for img_version, img_name in [(original_pil, 'original'), (processed_pil, 'processed')]:
                    try:
                        # Get OCR data
                        ocr_data = pytesseract.image_to_data(
                            img_version, 
                            output_type=pytesseract.Output.DICT,
                            config=config
                        )
                        
                        # Extract elements with very low threshold
                        elements = self.extract_elements_from_data(ocr_data, threshold=5)
                        
                        if elements:
                            all_results.append({
                                'config': config,
                                'image_version': img_name,
                                'description': description,
                                'elements': elements,
                                'score': self.calculate_result_score(elements)
                            })
                            
                    except Exception as e:
                        print(f"Config {config} with {img_name} failed: {e}")
                        continue
            
            # Return the best result
            if all_results:
                best_result = max(all_results, key=lambda x: x['score'])
                print(f"Best result: {best_result['description']} with {best_result['image_version']} image")
                return best_result['elements']
            else:
                return []
                
        except Exception as e:
            print(f"Error in comprehensive extraction: {str(e)}")
            return []
    
    def extract_elements_from_data(self, ocr_data, threshold=10):
        """Extract text elements from OCR data with flexible threshold"""
        elements = []
        
        for i in range(len(ocr_data['text'])):
            text = ocr_data['text'][i].strip()
            confidence = int(ocr_data['conf'][i]) if ocr_data['conf'][i] != '-1' else 0
            
            # More lenient text filtering
            if text and len(text.strip()) >= 1 and confidence >= threshold:
                # Minimal cleaning to preserve original text
                cleaned_text = self.gentle_text_cleaning(text)
                if cleaned_text:
                    elements.append({
                        'text': cleaned_text,
                        'left': ocr_data['left'][i],
                        'top': ocr_data['top'][i],
                        'width': ocr_data['width'][i],
                        'height': ocr_data['height'][i],
                        'confidence': confidence
                    })
        
        return elements
    
    def gentle_text_cleaning(self, text):
        """Gentle text cleaning that preserves most characters"""
        if not text:
            return ""
        
        # Only remove excessive whitespace
        text = re.sub(r'\s+', ' ', text.strip())
        
        # Only fix obvious OCR errors, preserve most text
        obvious_errors = {
            '|': 'I',  # Common pipe to I error
            '0': 'O',  # Only in all-caps words
        }
        
        # Apply minimal corrections
        words = text.split()
        corrected_words = []
        
        for word in words:
            if word.isupper() and '0' in word:
                word = word.replace('0', 'O')
            if '|' in word:
                word = word.replace('|', 'I')
            corrected_words.append(word)
        
        return ' '.join(corrected_words)
    
    def calculate_result_score(self, elements):
        """Calculate quality score for OCR results"""
        if not elements:
            return 0
        
        total_confidence = sum(elem['confidence'] for elem in elements)
        avg_confidence = total_confidence / len(elements)
        total_text_length = sum(len(elem['text']) for elem in elements)
        
        # Score based on quantity, confidence, and text length
        score = (len(elements) * 0.4) + (avg_confidence * 0.3) + (total_text_length * 0.3)
        return score
    
    def create_structured_layout(self, text_elements):
        """Create structured layout that preserves document organization"""
        if not text_elements:
            return []
        
        # Sort by vertical position first
        text_elements.sort(key=lambda x: x['top'])
        
        # Group into logical rows with more flexible threshold
        rows = []
        current_row = []
        current_top = None
        
        for element in text_elements:
            if current_top is None or abs(element['top'] - current_top) <= self.row_height_threshold:
                current_row.append(element)
                current_top = element['top'] if current_top is None else current_top
            else:
                if current_row:
                    # Sort row by horizontal position
                    current_row.sort(key=lambda x: x['left'])
                    rows.append(current_row)
                current_row = [element]
                current_top = element['top']
        
        # Don't forget the last row
        if current_row:
            current_row.sort(key=lambda x: x['left'])
            rows.append(current_row)
        
        return rows
    
    def create_enhanced_excel(self, rows, filename):
        """Create Excel with better structure preservation"""
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Extracted_Text"
            
            if not rows:
                return None
            
            excel_row = 1
            
            for row_elements in rows:
                if not row_elements:
                    excel_row += 1
                    continue
                
                # For forms, try to detect columns more intelligently
                if len(row_elements) == 1:
                    # Single element - full width
                    cell = ws.cell(row=excel_row, column=1)
                    cell.value = row_elements[0]['text']
                    self.format_cell(cell, row_elements[0])
                else:
                    # Multiple elements - distribute across columns
                    for col_idx, element in enumerate(row_elements, 1):
                        cell = ws.cell(row=excel_row, column=col_idx)
                        cell.value = element['text']
                        self.format_cell(cell, element)
                
                excel_row += 1
            
            # Auto-adjust column widths
            self.adjust_column_widths(ws)
            
            # Save to temporary file
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
            wb.save(temp_file.name)
            return temp_file.name
            
        except Exception as e:
            print(f"Error creating Excel: {str(e)}")
            return None
    
    def format_cell(self, cell, element):
        """Apply formatting based on text characteristics"""
        text = element['text']
        confidence = element['confidence']
        
        # Set base alignment
        cell.alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')
        
        # Format based on content
        if text.isupper() and len(text) > 3:
            # Likely header
            cell.font = Font(bold=True, size=12, color="1F4E79")
        elif text.startswith(('•', '-', '*', '○')):
            # Bullet point
            cell.font = Font(size=10)
            cell.alignment = Alignment(indent=1, wrap_text=True, vertical='top')
        elif re.match(r'^\d+\.', text):
            # Numbered item
            cell.font = Font(bold=True, size=11)
        else:
            # Regular text
            cell.font = Font(size=10)
        
        # Add confidence indicator as comment for debugging
        if confidence < 50:
            cell.font = Font(size=10, italic=True, color="666666")
    
    def adjust_column_widths(self, worksheet):
        """Intelligently adjust column widths"""
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            
            for cell in column:
                try:
                    if cell.value:
                        cell_length = len(str(cell.value))
                        if cell_length > max_length:
                            max_length = cell_length
                except:
                    pass
            
            # Set reasonable width limits
            adjusted_width = min(max(max_length + 2, 15), 80)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    def process_image(self, image_file, filename):
        """Process image with enhanced extraction"""
        try:
            # Open image
            image = Image.open(image_file)
            
            # Use comprehensive extraction
            text_elements = self.extract_text_comprehensive(image)
            
            if not text_elements:
                return None, "No text detected. Try adjusting confidence threshold or check image quality."
            
            # Create structured layout
            rows = self.create_structured_layout(text_elements)
            
            # Create Excel file
            excel_path = self.create_enhanced_excel(rows, filename)
            
            if excel_path:
                return excel_path, f"Extracted {len(text_elements)} text elements organized into {len(rows)} rows"
            else:
                return None, "Failed to create Excel output"
                
        except Exception as e:
            return None, f"Error processing image: {str(e)}"
    
    def process_pdf(self, pdf_file, filename):
        """Process PDF with enhanced extraction"""
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                # Save PDF temporarily
                temp_pdf_path = os.path.join(temp_dir, "temp.pdf")
                with open(temp_pdf_path, "wb") as f:
                    f.write(pdf_file.read())
                
                # Convert with high quality
                pages = convert_from_path(temp_pdf_path, dpi=300, fmt='png')
                
                excel_files = []
                total_elements = 0
                
                for page_num, page in enumerate(pages, 1):
                    # Process each page
                    text_elements = self.extract_text_comprehensive(page)
                    
                    if text_elements:
                        rows = self.create_structured_layout(text_elements)
                        excel_path = self.create_enhanced_excel(rows, f"{filename}_page_{page_num}")
                        if excel_path:
                            excel_files.append((excel_path, f"Page {page_num}"))
                            total_elements += len(text_elements)
                
                if excel_files:
                    return excel_files, f"Processed {len(pages)} pages, extracted {total_elements} text elements"
                else:
                    return [], f"No text found in {len(pages)} pages"
                
        except Exception as e:
            return [], f"Error processing PDF: {str(e)}"
