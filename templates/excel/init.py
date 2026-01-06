"""
Excel Templates - Reusable Excel structures
"""

from .invoice_template import InvoiceTemplate
from .payroll_template import PayrollTemplate
from .expense_template import ExpenseTemplate
from .attendance_template import AttendanceTemplate

class ExcelTemplates:
    """Collection of all Excel templates"""
    
    invoice = InvoiceTemplate
    payroll = PayrollTemplate
    expense = ExpenseTemplate
    attendance = AttendanceTemplate
    
    @classmethod
    def get_template(cls, template_name: str):
        """Get template by name"""
        templates = {
            "invoice": cls.invoice,
            "payroll": cls.payroll,
            "expense": cls.expense,
            "attendance": cls.attendance,
        }
        return templates.get(template_name.lower())
    
    @classmethod
    def list_templates(cls) -> list:
        """List all available templates"""
        return ["invoice", "payroll", "expense", "attendance"]

__all__ = [
    "ExcelTemplates",
    "InvoiceTemplate",
    "PayrollTemplate",
    "ExpenseTemplate",
    "AttendanceTemplate",
]
