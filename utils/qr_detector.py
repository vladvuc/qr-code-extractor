"""QR code detection and decoding utilities."""

import logging
import numpy as np
import cv2
from PIL import Image
from pyzbar.pyzbar import decode
from typing import Optional

from config import MAX_IMAGE_DIMENSION


logger = logging.getLogger(__name__)


def detect_and_decode_qr(image: Image.Image) -> Optional[str]:
    """
    Detect and decode QR codes in an image.

    Uses multiple detection strategies:
    1. Direct decode attempt
    2. Enhanced image decode (grayscale, resize, equalization, thresholding)

    Args:
        image: PIL Image object

    Returns:
        Decoded QR code data as string, or None if no QR code found
    """
    try:
        # Convert PIL Image to numpy array for OpenCV
        img_array = np.array(image)

        # Convert RGB to BGR if needed (OpenCV uses BGR)
        if len(img_array.shape) == 3 and img_array.shape[2] == 3:
            img_array = cv2.cvtColor(img_array, cv2.COLOR_RGB2BGR)
        elif len(img_array.shape) == 3 and img_array.shape[2] == 4:
            img_array = cv2.cvtColor(img_array, cv2.COLOR_RGBA2BGR)

        # Convert to grayscale
        if len(img_array.shape) == 3:
            gray = cv2.cvtColor(img_array, cv2.COLOR_BGR2GRAY)
        else:
            gray = img_array

        # Attempt 1: Direct decode
        logger.debug("Attempt 1: Direct QR decode")
        qr_data = _decode_qr(gray)
        if qr_data:
            logger.debug(f"QR code found (direct decode): {qr_data}")
            return qr_data

        # Attempt 2: Enhanced image decode
        logger.debug("Attempt 2: Enhanced image decode")
        enhanced = _enhance_image(gray)
        qr_data = _decode_qr(enhanced)
        if qr_data:
            logger.debug(f"QR code found (enhanced decode): {qr_data}")
            return qr_data

        logger.debug("No QR code detected")
        return None

    except Exception as e:
        logger.error(f"Error during QR detection: {e}")
        return None


def _decode_qr(image_array: np.ndarray) -> Optional[str]:
    """
    Decode QR code from a numpy array using pyzbar.

    Args:
        image_array: Numpy array representing the image

    Returns:
        Decoded QR code data as string, or None if no QR code found
    """
    try:
        decoded_objects = decode(image_array)

        if decoded_objects:
            # Return the first QR code found
            qr_data = decoded_objects[0].data.decode('utf-8', errors='ignore')
            return qr_data

        return None
    except Exception as e:
        logger.debug(f"Decode error: {e}")
        return None


def _enhance_image(gray_image: np.ndarray) -> np.ndarray:
    """
    Enhance image for better QR code detection.

    Applies:
    - Resizing if too large
    - Histogram equalization for contrast
    - Adaptive thresholding

    Args:
        gray_image: Grayscale image as numpy array

    Returns:
        Enhanced grayscale image
    """
    enhanced = gray_image.copy()

    # Resize if too large
    height, width = enhanced.shape
    if max(height, width) > MAX_IMAGE_DIMENSION:
        scale = MAX_IMAGE_DIMENSION / max(height, width)
        new_width = int(width * scale)
        new_height = int(height * scale)
        enhanced = cv2.resize(enhanced, (new_width, new_height), interpolation=cv2.INTER_AREA)
        logger.debug(f"Resized image from {width}x{height} to {new_width}x{new_height}")

    # Apply histogram equalization for better contrast
    enhanced = cv2.equalizeHist(enhanced)

    # Apply adaptive thresholding
    enhanced = cv2.adaptiveThreshold(
        enhanced,
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        11,
        2
    )

    return enhanced
