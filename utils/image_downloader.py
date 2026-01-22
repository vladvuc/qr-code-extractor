"""Image downloading utilities with retry logic."""

import logging
import time
from typing import Optional
import requests
from PIL import Image
from io import BytesIO

from config import REQUEST_TIMEOUT, MAX_RETRIES, USER_AGENT


logger = logging.getLogger(__name__)


def download_image(url: str, timeout: int = REQUEST_TIMEOUT, max_retries: int = MAX_RETRIES) -> Optional[Image.Image]:
    """
    Download an image from a URL with retry logic.

    Args:
        url: URL of the image to download
        timeout: Request timeout in seconds
        max_retries: Maximum number of retry attempts

    Returns:
        PIL Image object if successful, None otherwise
    """
    # Validate URL format
    if not url or not isinstance(url, str):
        logger.warning(f"Invalid URL format: {url}")
        return None

    url = url.strip()
    if not url.startswith(('http://', 'https://')):
        logger.warning(f"URL must start with http:// or https://: {url}")
        return None

    headers = {
        'User-Agent': USER_AGENT
    }

    attempt = 0
    last_error = None

    while attempt <= max_retries:
        try:
            logger.debug(f"Downloading image (attempt {attempt + 1}/{max_retries + 1}): {url}")

            response = requests.get(url, headers=headers, timeout=timeout, allow_redirects=True)
            response.raise_for_status()

            # Validate content type
            content_type = response.headers.get('Content-Type', '').lower()
            if not content_type.startswith('image/'):
                logger.warning(f"Non-image content type: {content_type} for URL: {url}")
                return None

            # Convert response to PIL Image
            image = Image.open(BytesIO(response.content))
            logger.debug(f"Successfully downloaded image: {url}")
            return image

        except requests.exceptions.Timeout as e:
            last_error = f"Timeout: {e}"
            logger.warning(f"Timeout downloading {url} (attempt {attempt + 1}/{max_retries + 1})")
            attempt += 1
            if attempt <= max_retries:
                time.sleep(1)  # Brief delay before retry

        except requests.exceptions.RequestException as e:
            last_error = f"Request error: {e}"
            logger.warning(f"Request error downloading {url}: {e}")
            return None

        except Exception as e:
            last_error = f"Unexpected error: {e}"
            logger.warning(f"Unexpected error downloading {url}: {e}")
            return None

    logger.error(f"Failed to download image after {max_retries + 1} attempts: {url}. Last error: {last_error}")
    return None
