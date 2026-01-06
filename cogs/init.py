"""
Discord Cogs - Command modules
"""

from .excel_cog import ExcelCog
from .finance_cog import FinanceCog
from .tax_cog import TaxCog
from .invoice_cog import InvoiceCog
from .analyst_cog import AnalystCog
from .writer_cog import WriterCog
from .hr_cog import HRCog

__all__ = [
    "ExcelCog",
    "FinanceCog", 
    "TaxCog",
    "InvoiceCog",
    "AnalystCog",
    "WriterCog",
    "HRCog",
]
