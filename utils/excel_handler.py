"""Excel file handling utilities."""

import os
import pandas as pd
from typing import Optional


def load_excel(file_path: str) -> pd.DataFrame:
    """
    Load an Excel file into a pandas DataFrame.

    Args:
        file_path: Path to the Excel file

    Returns:
        DataFrame containing the Excel data

    Raises:
        FileNotFoundError: If the file doesn't exist
        Exception: If the file can't be read
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Excel file not found: {file_path}")

    try:
        df = pd.read_excel(file_path, engine='openpyxl')
        return df
    except Exception as e:
        raise Exception(f"Failed to read Excel file: {e}")


def save_excel(df: pd.DataFrame, output_path: str) -> None:
    """
    Save a DataFrame to an Excel file.

    Args:
        df: DataFrame to save
        output_path: Path where the Excel file will be saved

    Raises:
        Exception: If the file can't be saved
    """
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    try:
        df.to_excel(output_path, engine='openpyxl', index=False)
    except Exception as e:
        raise Exception(f"Failed to save Excel file: {e}")
