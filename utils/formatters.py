"""
Formatters - Format data untuk Indonesia (Rupiah, tanggal, dll)
"""

import logging
from datetime import datetime, date
from typing import Union, Optional
from decimal import Decimal

from config import CURRENCY_SYMBOL, DATE_FORMAT, DATETIME_FORMAT

logger = logging.getLogger("office_bot.formatters")

# =============================================================================
# FORMATTERS CLASS
# =============================================================================

class Formatters:
    """
    Formatting utilities untuk format Indonesia
    """
    
    # =========================================================================
    # CURRENCY (RUPIAH)
    # =========================================================================
    
    @staticmethod
    def format_rupiah(
        amount: Union[int, float, Decimal, str],
        show_decimal: bool = False,
        show_symbol: bool = True
    ) -> str:
        """
        Format angka ke Rupiah Indonesia
        
        Examples:
            1500000 -> "Rp 1.500.000"
            1500000.50 -> "Rp 1.500.000,50"
        """
        try:
            # Convert to float
            if isinstance(amount, str):
                # Remove existing formatting
                amount = amount.replace(CURRENCY_SYMBOL, "")
                amount = amount.replace(".", "")
                amount = amount.replace(",", ".")
                amount = float(amount)
            else:
                amount = float(amount)
            
            # Split integer and decimal
            is_negative = amount < 0
            amount = abs(amount)
            
            integer_part = int(amount)
            decimal_part = amount - integer_part
            
            # Format integer with dots as thousand separators
            integer_str = f"{integer_part:,}".replace(",", ".")
            
            # Add decimal if needed
            if show_decimal:
                decimal_str = f"{decimal_part:.2f}"[2:]  # Get ".XX" part
                result = f"{integer_str},{decimal_str}"
            else:
                result = integer_str
            
            # Add currency symbol
            if show_symbol:
                result = f"{CURRENCY_SYMBOL} {result}"
            
            # Add negative sign
            if is_negative:
                result = f"-{result}"
            
            return result
            
        except Exception as e:
            logger.error(f"Error formatting Rupiah: {e}")
            return str(amount)
    
    @staticmethod
    def parse_rupiah(rupiah_str: str) -> float:
        """
        Parse Rupiah string to float
        
        Examples:
            "Rp 1.500.000" -> 1500000.0
            "Rp 1.500.000,50" -> 1500000.50
        """
        try:
            # Remove currency symbol
            value = rupiah_str.replace(CURRENCY_SYMBOL, "")
            value = value.strip()
            
            # Replace Indonesian formatting
            value = value.replace(".", "")  # Remove thousand separator
            value = value.replace(",", ".")  # Replace decimal separator
            
            return float(value)
            
        except Exception as e:
            logger.error(f"Error parsing Rupiah: {e}")
            return 0.0
    
    # =========================================================================
    # DATES
    # =========================================================================
    
    @staticmethod
    def format_date(
        date_obj: Union[datetime, date, str],
        format_str: Optional[str] = None
    ) -> str:
        """
        Format date to Indonesian format (DD/MM/YYYY)
        """
        if format_str is None:
            format_str = DATE_FORMAT
        
        try:
            if isinstance(date_obj, str):
                # Try to parse string
                date_obj = datetime.fromisoformat(date_obj)
            
            if isinstance(date_obj, date) and not isinstance(date_obj, datetime):
                date_obj = datetime.combine(date_obj, datetime.min.time())
            
            return date_obj.strftime(format_str)
            
        except Exception as e:
            logger.error(f"Error formatting date: {e}")
            return str(date_obj)
    
    @staticmethod
    def format_datetime(
        datetime_obj: Union[datetime, str],
        format_str: Optional[str] = None
    ) -> str:
        """
        Format datetime to Indonesian format (DD/MM/YYYY HH:MM)
        """
        if format_str is None:
            format_str = DATETIME_FORMAT
        
        return Formatters.format_date(datetime_obj, format_str)
    
    @staticmethod
    def parse_date(date_str: str, format_str: Optional[str] = None) -> datetime:
        """
        Parse Indonesian date string to datetime
        """
        if format_str is None:
            format_str = DATE_FORMAT
        
        try:
            return datetime.strptime(date_str, format_str)
        except Exception as e:
            logger.error(f"Error parsing date: {e}")
            raise
    
    # =========================================================================
    # NUMBERS
    # =========================================================================
    
    @staticmethod
    def format_number(
        number: Union[int, float],
        decimals: int = 0,
        use_separator: bool = True
    ) -> str:
        """
        Format number with Indonesian formatting
        
        Examples:
            1500000 -> "1.500.000"
            1500000.5 -> "1.500.000,50"
        """
        try:
            if decimals > 0:
                formatted = f"{number:,.{decimals}f}"
            else:
                formatted = f"{int(number):,}"
            
            if use_separator:
                # Replace English formatting with Indonesian
                formatted = formatted.replace(",", "TEMP")
                formatted = formatted.replace(".", ",")
                formatted = formatted.replace("TEMP", ".")
            
            return formatted
            
        except Exception as e:
            logger.error(f"Error formatting number: {e}")
            return str(number)
    
    @staticmethod
    def format_percentage(
        value: Union[int, float],
        decimals: int = 2
    ) -> str:
        """
        Format percentage
        
        Examples:
            0.15 -> "15,00%"
            0.1 -> "10,00%"
        """
        try:
            percentage = value * 100 if value <= 1 else value
            formatted = f"{percentage:.{decimals}f}%"
            formatted = formatted.replace(".", ",")
            return formatted
        except Exception as e:
            logger.error(f"Error formatting percentage: {e}")
            return str(value)
    
    # =========================================================================
    # TEXT
    # =========================================================================
    
    @staticmethod
    def terbilang(number: Union[int, float]) -> str:
        """
        Convert number to Indonesian words
        
        Examples:
            1500000 -> "satu juta lima ratus ribu rupiah"
        """
        # Simple implementation for common cases
        # Full implementation would be much longer
        
        ones = ["", "satu", "dua", "tiga", "empat", "lima", "enam", "tujuh", "delapan", "sembilan"]
        teens = ["sepuluh", "sebelas", "dua belas", "tiga belas", "empat belas", 
                 "lima belas", "enam belas", "tujuh belas", "delapan belas", "sembilan belas"]
        tens = ["", "", "dua puluh", "tiga puluh", "empat puluh", "lima puluh",
                "enam puluh", "tujuh puluh", "delapan puluh", "sembilan puluh"]
        
        def convert_below_thousand(n):
            if n == 0:
                return ""
            elif n < 10:
                return ones[n]
            elif n < 20:
                return teens[n - 10]
            elif n < 100:
                return tens[n // 10] + (" " + ones[n % 10] if n % 10 != 0 else "")
            else:
                hundred_part = "seratus" if n // 100 == 1 else ones[n // 100] + " ratus"
                remainder = convert_below_thousand(n % 100)
                return hundred_part + (" " + remainder if remainder else "")
        
        try:
            number = int(number)
            
            if number == 0:
                return "nol rupiah"
            
            if number < 0:
                return "minus " + Formatters.terbilang(abs(number))
            
            # Billions
            if number >= 1_000_000_000:
                billions = number // 1_000_000_000
                remainder = number % 1_000_000_000
                result = convert_below_thousand(billions) + " miliar"
                if remainder:
                    result += " " + Formatters.terbilang(remainder).replace(" rupiah", "")
                return result + " rupiah"
            
            # Millions
            if number >= 1_000_000:
                millions = number // 1_000_000
                remainder = number % 1_000_000
                result = convert_below_thousand(millions) + " juta"
                if remainder:
                    result += " " + Formatters.terbilang(remainder).replace(" rupiah", "")
                return result + " rupiah"
            
            # Thousands
            if number >= 1_000:
                thousands = number // 1_000
                remainder = number % 1_000
                if thousands == 1:
                    result = "seribu"
                else:
                    result = convert_below_thousand(thousands) + " ribu"
                if remainder:
                    result += " " + convert_below_thousand(remainder)
                return result + " rupiah"
            
            return convert_below_thousand(number) + " rupiah"
            
        except Exception as e:
            logger.error(f"Error in terbilang: {e}")
            return str(number)
    
    @staticmethod
    def capitalize_indonesian(text: str) -> str:
        """
        Capitalize text properly for Indonesian
        """
        # Words that should not be capitalized (unless first word)
        lowercase_words = ["dan", "atau", "di", "ke", "dari", "untuk", "pada", "yang"]
        
        words = text.lower().split()
        result = []
        
        for i, word in enumerate(words):
            if i == 0 or word not in lowercase_words:
                result.append(word.capitalize())
            else:
                result.append(word)
        
        return " ".join(result)
    
    # =========================================================================
    # FILE SIZE
    # =========================================================================
    
    @staticmethod
    def format_file_size(size_bytes: int) -> str:
        """
        Format file size to human readable
        
        Examples:
            1024 -> "1 KB"
            1048576 -> "1 MB"
        """
        for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
            if size_bytes < 1024.0:
                return f"{size_bytes:.2f} {unit}"
            size_bytes /= 1024.0
        return f"{size_bytes:.2f} PB"
