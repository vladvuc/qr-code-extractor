"""Enhanced QR code processing script with detailed reporting."""

import logging
import os
from datetime import datetime
import pandas as pd
from typing import Dict, List, Tuple
import sys

from utils.excel_handler import load_excel, save_excel
from utils.image_downloader import download_image
from utils.qr_detector import detect_and_decode_qr


def setup_logging() -> str:
    """Set up logging to both file and console."""
    logs_dir = "logs"
    os.makedirs(logs_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(logs_dir, f"processing_{timestamp}.log")

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )

    return log_file


def process_all_images() -> None:
    """
    Process all 246 images from Test_Sort_V2.xlsx.
    Detect QR codes, extract alphanumeric codes, and generate detailed report.
    """
    logger = logging.getLogger(__name__)

    # Setup logging
    log_file = setup_logging()
    logger.info("="*80)
    logger.info("STARTING QR CODE PROCESSING FOR ALL 246 IMAGES")
    logger.info("="*80)
    logger.info(f"Log file: {log_file}")

    # Configuration
    INPUT_FILE = "Test_Sort_V2.xlsx"
    OUTPUT_FILE = "Test_Sort_V2_Analyzed.xlsx"
    URL_COLUMN = "url"
    STICKER_COLUMN = "Photo Okret Sticker"
    QR_CODE_COLUMN = "QR_CODE"

    try:
        # Load Excel file
        logger.info(f"Loading Excel file: {INPUT_FILE}")
        df = load_excel(INPUT_FILE)
        total_rows = len(df)
        logger.info(f"Loaded {total_rows} rows from Excel")
        logger.info(f"Columns: {df.columns.tolist()}")

        # Verify URL column exists
        if URL_COLUMN not in df.columns:
            raise ValueError(f"Column '{URL_COLUMN}' not found. Available columns: {df.columns.tolist()}")

        # Initialize or reset result columns
        df[STICKER_COLUMN] = False
        df[QR_CODE_COLUMN] = ""

        # Tracking structures
        stats = {
            'total': total_rows,
            'processed': 0,
            'qr_found': 0,
            'no_qr': 0,
            'errors': 0,
            'skipped': 0
        }

        successful_qr_codes: List[Dict] = []
        failed_images: List[Dict] = []
        error_breakdown = {}

        # Process each row
        logger.info("\nStarting image processing...")
        logger.info("-" * 80)

        for index, row in df.iterrows():
            row_num = index + 1
            url = row[URL_COLUMN]

            # Skip empty URLs
            if pd.isna(url) or not str(url).strip():
                logger.info(f"[{row_num}/{total_rows}] Skipping - Empty URL")
                stats['skipped'] += 1
                failed_images.append({
                    'row': row_num,
                    'url': 'EMPTY',
                    'error': 'Empty or missing URL'
                })
                continue

            url = str(url).strip()

            # Progress indicator
            if row_num % 10 == 0:
                progress_pct = (row_num / total_rows) * 100
                logger.info(f"\n*** PROGRESS: {row_num}/{total_rows} ({progress_pct:.1f}%) ***\n")

            logger.info(f"[{row_num}/{total_rows}] Processing: {url[:80]}...")

            try:
                # Download image
                image = download_image(url)

                if image is None:
                    error_msg = "Failed to download image"
                    logger.warning(f"[{row_num}/{total_rows}] {error_msg}")
                    df.at[index, STICKER_COLUMN] = False
                    df.at[index, QR_CODE_COLUMN] = ""
                    stats['errors'] += 1
                    failed_images.append({
                        'row': row_num,
                        'url': url,
                        'error': error_msg
                    })
                    error_breakdown[error_msg] = error_breakdown.get(error_msg, 0) + 1
                    continue

                # Detect and decode QR code
                qr_data = detect_and_decode_qr(image)
                image.close()

                if qr_data:
                    # QR code found
                    logger.info(f"[{row_num}/{total_rows}] ✓ QR CODE FOUND: {qr_data}")
                    df.at[index, STICKER_COLUMN] = True
                    df.at[index, QR_CODE_COLUMN] = qr_data
                    stats['qr_found'] += 1
                    successful_qr_codes.append({
                        'row': row_num,
                        'qr_code': qr_data,
                        'url': url
                    })
                else:
                    # No QR code detected
                    logger.info(f"[{row_num}/{total_rows}] ✗ No QR code detected")
                    df.at[index, STICKER_COLUMN] = False
                    df.at[index, QR_CODE_COLUMN] = ""
                    stats['no_qr'] += 1
                    failed_images.append({
                        'row': row_num,
                        'url': url,
                        'error': 'No QR code detected in image'
                    })
                    error_breakdown['No QR code detected'] = error_breakdown.get('No QR code detected', 0) + 1

                stats['processed'] += 1

            except Exception as e:
                error_msg = f"Unexpected error: {str(e)}"
                logger.error(f"[{row_num}/{total_rows}] {error_msg}")
                df.at[index, STICKER_COLUMN] = False
                df.at[index, QR_CODE_COLUMN] = ""
                stats['errors'] += 1
                failed_images.append({
                    'row': row_num,
                    'url': url,
                    'error': error_msg
                })
                error_breakdown['Unexpected error'] = error_breakdown.get('Unexpected error', 0) + 1

        # Save results
        logger.info(f"\nSaving results to: {OUTPUT_FILE}")
        save_excel(df, OUTPUT_FILE)
        logger.info(f"Results saved successfully!")

        # Generate detailed report
        print_detailed_report(stats, successful_qr_codes, failed_images, error_breakdown, OUTPUT_FILE, log_file)

        # Save detailed report to file
        save_detailed_report(stats, successful_qr_codes, failed_images, error_breakdown, OUTPUT_FILE, log_file)

    except Exception as e:
        logger.error(f"FATAL ERROR: {e}", exc_info=True)
        print(f"\n{'='*80}")
        print(f"ERROR: Processing failed. Check log file: {log_file}")
        print(f"{'='*80}")
        raise


def print_detailed_report(stats: Dict, successful: List[Dict], failed: List[Dict],
                         error_breakdown: Dict, output_file: str, log_file: str) -> None:
    """Print detailed processing report to console."""

    success_rate = (stats['qr_found'] / stats['total'] * 100) if stats['total'] > 0 else 0

    print("\n" + "="*80)
    print("QR CODE PROCESSING COMPLETE - DETAILED REPORT")
    print("="*80)
    print("\n--- SUMMARY STATISTICS ---")
    print(f"Total rows in file:        {stats['total']}")
    print(f"Successfully processed:    {stats['processed']}")
    print(f"QR codes found:            {stats['qr_found']} ({success_rate:.1f}%)")
    print(f"No QR code detected:       {stats['no_qr']}")
    print(f"Errors:                    {stats['errors']}")
    print(f"Skipped (empty URLs):      {stats['skipped']}")

    print("\n--- ERROR BREAKDOWN ---")
    if error_breakdown:
        for error_type, count in sorted(error_breakdown.items(), key=lambda x: x[1], reverse=True):
            print(f"  {error_type}: {count}")
    else:
        print("  No errors!")

    print("\n--- SUCCESSFULLY EXTRACTED QR CODES ---")
    if successful:
        print(f"Total QR codes found: {len(successful)}")
        print("\nFirst 20 QR codes:")
        for item in successful[:20]:
            print(f"  Row {item['row']:3d}: {item['qr_code']}")
        if len(successful) > 20:
            print(f"  ... and {len(successful) - 20} more")
    else:
        print("  No QR codes found.")

    print("\n--- FAILED/NO QR CODE IMAGES ---")
    if failed:
        print(f"Total failed: {len(failed)}")
        print("\nFirst 10 failures:")
        for item in failed[:10]:
            print(f"  Row {item['row']:3d}: {item['error']}")
            if item['url'] != 'EMPTY':
                print(f"         URL: {item['url'][:70]}...")
        if len(failed) > 10:
            print(f"  ... and {len(failed) - 10} more (see detailed report file)")
    else:
        print("  No failures!")

    print("\n--- OUTPUT FILES ---")
    print(f"Excel output:       {output_file}")
    print(f"Processing log:     {log_file}")
    print(f"Detailed report:    processing_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")

    print("\n" + "="*80)
    print(f"SUCCESS RATE: {success_rate:.1f}% ({stats['qr_found']}/{stats['total']})")
    print("="*80 + "\n")


def save_detailed_report(stats: Dict, successful: List[Dict], failed: List[Dict],
                        error_breakdown: Dict, output_file: str, log_file: str) -> None:
    """Save detailed report to a text file."""

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_file = f"processing_report_{timestamp}.txt"

    with open(report_file, 'w') as f:
        success_rate = (stats['qr_found'] / stats['total'] * 100) if stats['total'] > 0 else 0

        f.write("="*80 + "\n")
        f.write("QR CODE PROCESSING - DETAILED REPORT\n")
        f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write("="*80 + "\n\n")

        f.write("SUMMARY STATISTICS\n")
        f.write("-"*80 + "\n")
        f.write(f"Total rows in file:        {stats['total']}\n")
        f.write(f"Successfully processed:    {stats['processed']}\n")
        f.write(f"QR codes found:            {stats['qr_found']} ({success_rate:.1f}%)\n")
        f.write(f"No QR code detected:       {stats['no_qr']}\n")
        f.write(f"Errors:                    {stats['errors']}\n")
        f.write(f"Skipped (empty URLs):      {stats['skipped']}\n\n")

        f.write("ERROR BREAKDOWN\n")
        f.write("-"*80 + "\n")
        if error_breakdown:
            for error_type, count in sorted(error_breakdown.items(), key=lambda x: x[1], reverse=True):
                f.write(f"  {error_type}: {count}\n")
        else:
            f.write("  No errors!\n")
        f.write("\n")

        f.write("SUCCESSFULLY EXTRACTED QR CODES\n")
        f.write("-"*80 + "\n")
        if successful:
            f.write(f"Total: {len(successful)}\n\n")
            for item in successful:
                f.write(f"Row {item['row']:3d}: {item['qr_code']}\n")
                f.write(f"       URL: {item['url']}\n\n")
        else:
            f.write("  No QR codes found.\n")
        f.write("\n")

        f.write("FAILED/NO QR CODE IMAGES\n")
        f.write("-"*80 + "\n")
        if failed:
            f.write(f"Total: {len(failed)}\n\n")
            for item in failed:
                f.write(f"Row {item['row']:3d}: {item['error']}\n")
                f.write(f"       URL: {item['url']}\n\n")
        else:
            f.write("  No failures!\n")
        f.write("\n")

        f.write("OUTPUT FILES\n")
        f.write("-"*80 + "\n")
        f.write(f"Excel output:       {output_file}\n")
        f.write(f"Processing log:     {log_file}\n")
        f.write(f"Detailed report:    {report_file}\n\n")

        f.write("="*80 + "\n")
        f.write(f"FINAL SUCCESS RATE: {success_rate:.1f}% ({stats['qr_found']}/{stats['total']})\n")
        f.write("="*80 + "\n")


if __name__ == "__main__":
    try:
        process_all_images()
    except KeyboardInterrupt:
        print("\n\nProcessing interrupted by user.")
        sys.exit(1)
    except Exception:
        sys.exit(1)
