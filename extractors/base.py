"""
Base extractor class for PDF processing
"""

import logging
import re
from typing import Optional
from abc import ABC, abstractmethod

# Try to import OCR dependencies, but don't fail if missing
OCR_AVAILABLE = False
try:
    import pytesseract
    from pdf2image import convert_from_path
    OCR_AVAILABLE = True
except ImportError as e:
    # OCR will be disabled but app will still work for text-based PDFs
    pytesseract = None
    convert_from_path = None

from models.bill_data import BillData

class BaseExtractor(ABC):
    """Base class for PDF data extraction"""

    def __init__(self):
        self.logger = self._setup_logging()
        if not OCR_AVAILABLE:
            self.logger.warning("OCR dependencies not available - scanned PDFs may not work")

    def _setup_logging(self):
        logging.basicConfig(level=logging.INFO)
        return logging.getLogger(self.__class__.__name__)

    @abstractmethod
    def extract_data(self, pdf_path: str) -> Optional[BillData]:
        """Extract data from PDF - must be implemented by subclasses"""
        pass

    def _ocr_extract(self, pdf_path: str) -> Optional[str]:
        """Extract text using OCR for scanned PDFs"""
        if not OCR_AVAILABLE:
            self.logger.warning(f"OCR not available - cannot process scanned PDF: {pdf_path}")
            return None

        try:
            images = convert_from_path(pdf_path, dpi=300)
            text = ""
            for image in images:
                try:
                    text += pytesseract.image_to_string(image, config='--psm 6') + "\n"
                except Exception:
                    text += pytesseract.image_to_string(image, config='--psm 4') + "\n"
            return text
        except Exception as e:
            self.logger.error(f"OCR extraction failed: {e}")
            return None

    def _extract_pattern(self, text: str, pattern: str) -> Optional[str]:
        """Extract first match of regex pattern"""
        if not text:
            return None
        match = re.search(pattern, text, re.IGNORECASE)
        return match.group(1).strip() if match else None

    def _extract_currency(self, text: str, pattern: str) -> Optional[float]:
        """Extract currency value and convert to float"""
        if not text:
            return None
        match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
        if not match:
            return None

        value_str = match.group(1)
        value_str = value_str.replace(',', '').replace('$', '').strip()

        is_negative = False
        if value_str.startswith('(') and value_str.endswith(')'):
            is_negative = True
            value_str = value_str[1:-1]

        try:
            value = float(value_str)
            return -value if is_negative else value
        except ValueError:
            return None

    def _extract_number(self, text: str, pattern: str) -> Optional[int]:
        """Extract number and convert to int"""
        if not text:
            return None
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            value_str = match.group(1).replace(',', '')
            try:
                return int(value_str)
            except ValueError:
                return None
        return None