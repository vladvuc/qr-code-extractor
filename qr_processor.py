"""Main script for processing QR codes in images from Excel file."""

import logging
import os
from datetime import datetime
import pandas as pd

from config import INPUT_FILE, QR_CODE_COLUMN, STICKER_COLUMN
from utils.excel_handler import load_excel, save_excel
from utils.image_downloader import download_image
from utils.qr_detector import detect_and_decode_qr


def setup_logging() -> str:
    """
    Set up logging to both file and console.

    Returns:
        Path to the log file
    """
    # Create logs directory
    logs_dir = "logs"
    os.makedirs(logs_dir, exist_ok=True)

    # Create log filename with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(logs_dir, f"processing_{timestamp}.log")

    # Configure logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )

    return log_file


def process_excel() -> None:
    """
    Main processing function.

    Loads Excel, processes each row, detects QR codes, and saves results.
    """
    logger = logging.getLogger(__name__)

    # Setup logging
    log_file = setup_logging()
    logger.info(f"Starting QR code processing. Log file: {log_file}")

    try:
        # Load Excel file
        logger.info(f"Loading Excel file: {INPUT_FILE}")
        df = load_excel(INPUT_FILE)
        logger.info(f"Loaded {len(df)} rows from Excel")

        # Use 'url' column for image URLs
        url_column = 'url'
        if url_column not in df.columns:
            raise ValueError(f"Column '{url_column}' not found in Excel file. Available columns: {df.columns.tolist()}")
        logger.info(f"URL column: {url_column}")

        # Create or reset QR_CODE and STICKER columns
        if QR_CODE_COLUMN in df.columns:
            logger.info(f"Column '{QR_CODE_COLUMN}' already exists, will overwrite")
        if STICKER_COLUMN in df.columns:
            logger.info(f"Column '{STICKER_COLUMN}' already exists, will overwrite")
        df[QR_CODE_COLUMN] = ""
        df[STICKER_COLUMN] = False

        # Statistics
        stats = {
            'total': len(df),
            'qr_found': 0,
            'no_qr': 0,
            'errors': 0,
            'skipped': 0
        }

        # Process each row
        for index, row in df.iterrows():
            url = row[url_column]

            # Skip if URL is empty or NaN
            if pd.isna(url) or not str(url).strip():
                logger.info(f"Row {index + 1}/{stats['total']}: Skipping empty URL")
                stats['skipped'] += 1
                continue

            url = str(url).strip()
            logger.info(f"Row {index + 1}/{stats['total']}: Processing {url}")

            try:
                # Download image
                image = download_image(url)

                if image is None:
                    # Download failed
                    logger.warning(f"Row {index + 1}: Failed to download image")
                    df.at[index, QR_CODE_COLUMN] = ""
                    df.at[index, STICKER_COLUMN] = False
                    stats['errors'] += 1
                    continue

                # Detect QR code
                qr_data = detect_and_decode_qr(image)

                # Close image to free memory
                image.close()

                if qr_data:
                    # QR code found
                    logger.info(f"Row {index + 1}: QR code found: {qr_data}")
                    df.at[index, QR_CODE_COLUMN] = qr_data
                    df.at[index, STICKER_COLUMN] = True
                    stats['qr_found'] += 1
                else:
                    # No QR code found
                    logger.info(f"Row {index + 1}: No QR code detected")
                    df.at[index, QR_CODE_COLUMN] = ""
                    df.at[index, STICKER_COLUMN] = False
                    stats['no_qr'] += 1

            except Exception as e:
                logger.error(f"Row {index + 1}: Unexpected error: {e}")
                df.at[index, QR_CODE_COLUMN] = ""
                df.at[index, STICKER_COLUMN] = False
                stats['errors'] += 1

        # Save results
        output_dir = "output"
        os.makedirs(output_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = os.path.join(output_dir, f"company_export_processed_{timestamp}.xlsx")

        logger.info(f"Saving results to: {output_file}")
        save_excel(df, output_file)

        # Print summary
        print("\n" + "="*60)
        print("PROCESSING COMPLETE")
        print("="*60)
        print(f"Total rows processed: {stats['total']}")
        print(f"QR codes found: {stats['qr_found']}")
        print(f"No QR code: {stats['no_qr']}")
        print(f"Errors: {stats['errors']}")
        print(f"Skipped (empty URLs): {stats['skipped']}")
        print(f"\nOutput file: {output_file}")
        print(f"Log file: {log_file}")
        print("="*60)

    except Exception as e:
        logger.error(f"Fatal error during processing: {e}", exc_info=True)
        print(f"\nERROR: Processing failed. Check log file for details: {log_file}")
        raise


if __name__ == "__main__":
    process_excel()
