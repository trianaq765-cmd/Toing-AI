"""
Validators - Validate input data
"""

import logging
import re
from typing import Union, Optional
from datetime import datetime
from pathlib import Path

from config import MAX_FILE_SIZE_MB, MAX_EXCEL_ROWS, MAX_EXCEL_COLS

logger = logging.getLogger("office_bot.validators")

# =============================================================================
# VALIDATORS CLASS
# =============================================================================

class Validators:
    """
    Input validation utilities
    """
    
    # =========================================================================
    # TAX VALIDATORS
    # =========================================================================
    
    @staticmethod
    def validate_npwp(npwp: str) -> bool:
        """
        Validate NPWP (Nomor Pokok Wajib Pajak) Indonesia
        Format: XX.XXX.XXX.X-XXX.XXX (15 digits)
        """
        # Remove formatting
        npwp_clean = re.sub(r'[^0-9]', '', npwp)
        
        if len(npwp_clean) != 15:
            return False
        
        return True
    
    @staticmethod
    def validate_nik(nik: str) -> bool:
        """
        Validate NIK (Nomor Induk Kependudukan) Indonesia
        16 digits
        """
        nik_clean = re.sub(r'[^0-9]', '', nik)
        return len(nik_clean) == 16
    
    @staticmethod
    def validate_ptkp_status(status: str) -> bool:
        """
        Validate PTKP status
        Format: TK/0, TK/1, K/0, K/1, K/I/0, etc.
        """
        valid_statuses = [
            "TK/0", "TK/1", "TK/2", "TK/3",
            "K/0", "K/1", "K/2", "K/3",
            "K/I/0", "K/I/1", "K/I/2", "K/I/3"
        ]
        return status.upper() in valid_statuses
    
    # =========================================================================
    # FILE VALIDATORS
    # =========================================================================
    
    @staticmethod
    def validate_file_size(size_bytes: int, max_mb: Optional[int] = None) -> bool:
        """
        Validate file size
        """
        if max_mb is None:
            max_mb = MAX_FILE_SIZE_MB
        
        max_bytes = max_mb * 1024 * 1024
        return size_bytes <= max_bytes
    
    @staticmethod
    def validate_excel_file(file_path: Union[str, Path]) -> tuple[bool, Optional[str]]:
        """
        Validate Excel file
        Returns: (is_valid, error_message)
        """
        file_path = Path(file_path)
        
        # Check existence
        if not file_path.exists():
            return False, "File tidak ditemukan"
        
        # Check extension
        if file_path.suffix.lower() not in ['.xlsx', '.xls', '.xlsm']:
            return False, f"Format file tidak didukung: {file_path.suffix}"
        
        # Check size
        size_bytes = file_path.stat().st_size
        if not Validators.validate_file_size(size_bytes):
            return False, f"File terlalu besar (max {MAX_FILE_SIZE_MB}MB)"
        
        return True, None
    
    # =========================================================================
    # EMAIL & PHONE
    # =========================================================================
    
    @staticmethod
    def validate_email(email: str) -> bool:
        """
        Validate email address
        """
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(pattern, email) is not None
    
    @staticmethod
    def validate_phone_indonesia(phone: str) -> bool:
        """
        Validate Indonesian phone number
        Formats: 08xx, +628xx, 628xx
        """
        # Remove spaces, dashes, parentheses
        phone_clean = re.sub(r'[\s\-\(\)]', '', phone)
        
        # Check patterns
        patterns = [
            r'^08\d{8,11}$',      # 08xxxxxxxxx
            r'^\+628\d{8,11}$',   # +628xxxxxxxxx
            r'^628\d{8,11}$',     # 628xxxxxxxxx
        ]
        
        for pattern in patterns:
            if re.match(pattern, phone_clean):
                return True
        
        return False
    
    # =========================================================================
    # NUMERIC VALIDATORS
    # =========================================================================
    
    @staticmethod
    def validate_positive_number(value: Union[int, float, str]) -> bool:
        """
        Validate positive number
        """
        try:
            num = float(value)
            return num > 0
        except:
            return False
    
    @staticmethod
    def validate_percentage(value: Union[int, float, str]) -> bool:
        """
        Validate percentage (0-100 or 0-1)
        """
        try:
            num = float(value)
            return 0 <= num <= 100 or 0 <= num <= 1
        except:
            return False
    
    @staticmethod
    def validate_range(
        value: Union[int, float],
        min_val: Optional[Union[int, float]] = None,
        max_val: Optional[Union[int, float]] = None
    ) -> bool:
        """
        Validate number is within range
        """
        try:
            num = float(value)
            if min_val is not None and num < min_val:
                return False
            if max_val is not None and num > max_val:
                return False
            return True
        except:
            return False
    
    # =========================================================================
    # DATE VALIDATORS
    # =========================================================================
    
    @staticmethod
    def validate_date(date_str: str, format_str: str = "%d/%m/%Y") -> bool:
        """
        Validate date string
        """
        try:
            datetime.strptime(date_str, format_str)
            return True
        except:
            return False
    
    @staticmethod
    def validate_date_range(
        start_date: Union[str, datetime],
        end_date: Union[str, datetime]
    ) -> bool:
        """
        Validate date range (start <= end)
        """
        try:
            if isinstance(start_date, str):
                start_date = datetime.strptime(start_date, "%d/%m/%Y")
            if isinstance(end_date, str):
                end_date = datetime.strptime(end_date, "%d/%m/%Y")
            
            return start_date <= end_date
        except:
            return False
    
    # =========================================================================
    # EXCEL VALIDATORS
    # =========================================================================
    
    @staticmethod
    def validate_cell_reference(cell_ref: str) -> bool:
        """
        Validate Excel cell reference (e.g., A1, B10, AA100)
        """
        pattern = r'^[A-Z]+[1-9][0-9]*$'
        return re.match(pattern, cell_ref.upper()) is not None
    
    @staticmethod
    def validate_excel_formula(formula: str) -> bool:
        """
        Basic validation for Excel formula
        """
        if not formula.startswith('='):
            return False
        
        # Check balanced parentheses
        return formula.count('(') == formula.count(')')
    
    @staticmethod
    def validate_excel_dimensions(rows: int, cols: int) -> tuple[bool, Optional[str]]:
        """
        Validate Excel dimensions are within limits
        """
        if rows > MAX_EXCEL_ROWS:
            return False, f"Terlalu banyak baris (max {MAX_EXCEL_ROWS})"
        
        if cols > MAX_EXCEL_COLS:
            return False, f"Terlalu banyak kolom (max {MAX_EXCEL_COLS})"
        
        return True, None
