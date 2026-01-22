#!/usr/bin/env python3
"""
Deep QR Code Analysis Script for Excel Spreadsheet
Processes images from URLs and detects QR codes using multiple detection methods.
"""

import pandas as pd
import requests
from PIL import Image, ImageEnhance, ImageOps
from pyzbar import pyzbar
import cv2
import numpy as np
from io import BytesIO
from tqdm import tqdm
import time
import sys
from datetime import datetime

# Configuration
INPUT_FILE = "EXCEL/Sheet4.xlsx"
OUTPUT_FILE = "EXCEL/Sheet4_QR_Analyzed.xlsx"
ERROR_LOG_FILE = "qr_analysis_errors.log"
TEMP_SAVE_INTERVAL = 10  # Save intermediate results every 10 images
MAX_RETRIES = 3
RETRY_DELAY = 2  # seconds
REQUEST_TIMEOUT = 30  # seconds


class QRCodeAnalyzer:
    """Comprehensive QR code detection and analysis"""

    def __init__(self):
        self.error_log = []
        self.results = {
            'total_processed': 0,
            'qr_codes_found': 0,
            'multiple_qr_found': 0,
            'not_found': 0,
            'errors': 0,
            'corrupted_images': 0
        }

    def log_error(self, row_num, url, error_msg):
        """Log errors to memory and file"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] Row {row_num}: {error_msg} | URL: {url}"
        self.error_log.append(log_entry)
        print(f"ERROR: {log_entry}")

    def download_image(self, url, row_num):
        """Download image from URL with retry logic"""
        for attempt in range(MAX_RETRIES):
            try:
                response = requests.get(url, timeout=REQUEST_TIMEOUT)
                response.raise_for_status()

                # Validate image
                img = Image.open(BytesIO(response.content))
                img.verify()  # Verify it's a valid image

                # Re-open for actual use (verify closes the file)
                img = Image.open(BytesIO(response.content))
                return img

            except requests.exceptions.Timeout:
                if attempt < MAX_RETRIES - 1:
                    time.sleep(RETRY_DELAY * (attempt + 1))
                    continue
                else:
                    self.log_error(row_num, url, "Download timeout after retries")
                    return None

            except requests.exceptions.RequestException as e:
                if attempt < MAX_RETRIES - 1:
                    time.sleep(RETRY_DELAY * (attempt + 1))
                    continue
                else:
                    self.log_error(row_num, url, f"Download failed: {str(e)}")
                    return None

            except Exception as e:
                self.log_error(row_num, url, f"Image validation failed: {str(e)}")
                return None

        return None

    def detect_qr_with_pyzbar(self, img):
        """Primary QR detection using pyzbar"""
        try:
            # Convert PIL image to format pyzbar can use
            decoded_objects = pyzbar.decode(img)

            if decoded_objects:
                qr_codes = [obj for obj in decoded_objects if obj.type == 'QRCODE']
                return qr_codes
            return []
        except Exception as e:
            print(f"Pyzbar detection error: {str(e)}")
            return []

    def detect_qr_with_opencv(self, img):
        """Fallback QR detection using OpenCV"""
        try:
            # Convert PIL Image to OpenCV format
            img_array = np.array(img)

            # Convert RGB to BGR for OpenCV
            if len(img_array.shape) == 3 and img_array.shape[2] == 3:
                img_cv = cv2.cvtColor(img_array, cv2.COLOR_RGB2BGR)
            else:
                img_cv = img_array

            # Try OpenCV QR Code detector
            qr_detector = cv2.QRCodeDetector()
            data, bbox, _ = qr_detector.detectAndDecode(img_cv)

            if data:
                return [data]
            return []
        except Exception as e:
            print(f"OpenCV detection error: {str(e)}")
            return []

    def enhance_image(self, img):
        """Apply various enhancements to improve QR detection"""
        enhanced_images = []

        # Original image
        enhanced_images.append(img)

        # Convert to grayscale
        try:
            gray_img = ImageOps.grayscale(img)
            enhanced_images.append(gray_img)
        except:
            pass

        # Increase contrast
        try:
            enhancer = ImageEnhance.Contrast(img)
            enhanced_images.append(enhancer.enhance(2.0))
        except:
            pass

        # Increase brightness
        try:
            enhancer = ImageEnhance.Brightness(img)
            enhanced_images.append(enhancer.enhance(1.5))
        except:
            pass

        # Increase sharpness
        try:
            enhancer = ImageEnhance.Sharpness(img)
            enhanced_images.append(enhancer.enhance(2.0))
        except:
            pass

        return enhanced_images

    def analyze_image(self, img, row_num, url):
        """
        Comprehensive QR code analysis with multi-stage detection
        Returns: tuple (status_string, qr_data_list)
        """
        if img is None:
            return "ERROR: Download failed", []

        all_qr_data = []

        # Stage 1: Try pyzbar on original image
        qr_codes = self.detect_qr_with_pyzbar(img)
        if qr_codes:
            all_qr_data.extend([obj.data.decode('utf-8', errors='ignore') for obj in qr_codes])

        # Stage 2: If no QR found, try enhanced images with pyzbar
        if not all_qr_data:
            enhanced_images = self.enhance_image(img)
            for enhanced_img in enhanced_images:
                qr_codes = self.detect_qr_with_pyzbar(enhanced_img)
                if qr_codes:
                    all_qr_data.extend([obj.data.decode('utf-8', errors='ignore') for obj in qr_codes])
                    break  # Found QR codes, no need to continue

        # Stage 3: If still no QR found, try OpenCV
        if not all_qr_data:
            opencv_results = self.detect_qr_with_opencv(img)
            if opencv_results:
                all_qr_data.extend(opencv_results)

        # Remove duplicates while preserving order
        seen = set()
        unique_qr_data = []
        for data in all_qr_data:
            if data not in seen:
                seen.add(data)
                unique_qr_data.append(data)

        return unique_qr_data

    def format_qr_result(self, qr_data_list):
        """Format QR detection results according to requirements"""
        if not qr_data_list:
            return "NOT_FOUND"
        elif len(qr_data_list) == 1:
            return qr_data_list[0]
        else:
            return f"{len(qr_data_list)} QR codes found"

    def process_excel(self):
        """Main processing function"""
        print("=" * 70)
        print("QR Code Analysis Script")
        print("=" * 70)

        # Read Excel file
        try:
            print(f"\nReading input file: {INPUT_FILE}")
            df = pd.read_excel(INPUT_FILE)
            print(f"Loaded {len(df)} rows")

            # Validate required columns
            if 'url' not in df.columns:
                print("ERROR: 'url' column not found in Excel file")
                return

            # Ensure QR_CODE column exists and is of object type (string)
            if 'QR_CODE' not in df.columns:
                df['QR_CODE'] = ""

            # Convert QR_CODE column to object type to allow string assignments
            df['QR_CODE'] = df['QR_CODE'].astype('object')

        except Exception as e:
            print(f"ERROR: Failed to read Excel file: {str(e)}")
            return

        # Process each row
        print(f"\nProcessing {len(df)} images...\n")

        for idx, row in tqdm(df.iterrows(), total=len(df), desc="Analyzing QR codes"):
            row_num = idx + 2  # Excel row number (1-indexed + header)
            url = row['url']

            if pd.isna(url) or not url:
                df.at[idx, 'QR_CODE'] = "ERROR: Missing URL"
                self.results['errors'] += 1
                continue

            # Download image
            img = self.download_image(url, row_num)

            if img is None:
                df.at[idx, 'QR_CODE'] = "ERROR: Download failed"
                self.results['errors'] += 1
            else:
                # Analyze for QR codes
                try:
                    qr_data_list = self.analyze_image(img, row_num, url)
                    result_text = self.format_qr_result(qr_data_list)
                    df.at[idx, 'QR_CODE'] = result_text

                    # Update statistics
                    if result_text == "NOT_FOUND":
                        self.results['not_found'] += 1
                    elif " QR codes found" in result_text:
                        self.results['qr_codes_found'] += 1
                        self.results['multiple_qr_found'] += 1
                    else:
                        self.results['qr_codes_found'] += 1

                except Exception as e:
                    df.at[idx, 'QR_CODE'] = f"ERROR: Analysis failed"
                    self.log_error(row_num, url, f"Analysis exception: {str(e)}")
                    self.results['errors'] += 1

            self.results['total_processed'] += 1

            # Intermediate save
            if (idx + 1) % TEMP_SAVE_INTERVAL == 0:
                try:
                    df.to_excel(OUTPUT_FILE, index=False)
                except Exception as e:
                    print(f"\nWarning: Failed to save intermediate results: {str(e)}")

        # Final save
        print(f"\n\nSaving results to: {OUTPUT_FILE}")
        try:
            df.to_excel(OUTPUT_FILE, index=False)
            print("Results saved successfully!")
        except Exception as e:
            print(f"ERROR: Failed to save final results: {str(e)}")
            return

        # Save error log
        if self.error_log:
            with open(ERROR_LOG_FILE, 'w') as f:
                f.write("\n".join(self.error_log))
            print(f"\nError log saved to: {ERROR_LOG_FILE}")

        # Print summary report
        self.print_summary()

    def print_summary(self):
        """Print comprehensive summary report"""
        print("\n" + "=" * 70)
        print("ANALYSIS SUMMARY")
        print("=" * 70)
        print(f"Total images processed:     {self.results['total_processed']}")
        print(f"QR codes found:             {self.results['qr_codes_found']}")
        print(f"  - Single QR code:         {self.results['qr_codes_found'] - self.results['multiple_qr_found']}")
        print(f"  - Multiple QR codes:      {self.results['multiple_qr_found']}")
        print(f"QR codes not found:         {self.results['not_found']}")
        print(f"Errors/Failed downloads:    {self.results['errors']}")

        if self.results['total_processed'] > 0:
            success_rate = (self.results['qr_codes_found'] / self.results['total_processed']) * 100
            print(f"\nSuccess rate:               {success_rate:.1f}%")

        print("=" * 70)


def main():
    """Main entry point"""
    analyzer = QRCodeAnalyzer()

    try:
        analyzer.process_excel()
    except KeyboardInterrupt:
        print("\n\nProcessing interrupted by user.")
        print("Partial results may have been saved.")
        sys.exit(1)
    except Exception as e:
        print(f"\n\nFATAL ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
