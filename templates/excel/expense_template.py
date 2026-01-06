"""
Expense Report Template - Template laporan pengeluaran
"""

from datetime import datetime
from typing import Dict, Any, Optional, List

class ExpenseTemplate:
    """
    Expense Report template untuk laporan pengeluaran
    """
    
    @staticmethod
    def get_structure(
        employee_name: str = "NAMA KARYAWAN",
        department: str = "DEPARTEMEN",
        period: Optional[str] = None,
        expenses: Optional[List[Dict]] = None
    ) -> Dict[str, Any]:
        """
        Generate expense report structure
        
        Args:
            employee_name: Nama karyawan
            department: Departemen
            period: Periode laporan
            expenses: List of expenses [{date, category, description, amount, receipt}]
        """
        
        if period is None:
            now = datetime.now()
            period = f"{now.strftime('%B')} {now.year}"
        
        # Categories
        categories = [
            "Transport",
            "Makan",
            "Akomodasi",
            "Supplies",
            "Entertainment",
            "Lainnya"
        ]
        
        if expenses is None:
            expenses = []
        
        # Build expense rows
        expense_rows = []
        for i, exp in enumerate(expenses, 1):
            expense_rows.append([
                i,
                exp.get("date", ""),
                exp.get("category", ""),
                exp.get("description", ""),
                exp.get("amount", 0),
                exp.get("receipt", "Ya/Tidak"),
                exp.get("notes", "")
            ])
        
        # Add empty rows for template
        if not expenses:
            for i in range(1, 21):  # 20 empty rows
                expense_rows.append([
                    i, "", "", "", 0, "", ""
                ])
        
        last_row = 8 + len(expense_rows)
        
        structure = {
            "sheets": [{
                "name": "Expense Report",
                "data": [
                    ["LAPORAN PENGELUARAN / EXPENSE REPORT", "", "", "", "", "", ""],
                    ["", "", "", "", "", "", ""],
                    [f"Nama: {employee_name}", "", "", f"Departemen: {department}", "", "", ""],
                    [f"Periode: {period}", "", "", f"Tanggal Submit: {datetime.now().strftime('%d/%m/%Y')}", "", "", ""],
                    ["", "", "", "", "", "", ""],
                    ["", "", "", "", "", "", ""],
                    # Table Header
                    ["No", "Tanggal", "Kategori", "Deskripsi", "Jumlah (Rp)", "Bukti/Receipt", "Keterangan"],
                    ["", "", "", "", "", "", ""],
                ] + expense_rows + [
                    ["", "", "", "", "", "", ""],
                    ["", "", "", "TOTAL:", f"=SUM(E9:E{last_row})", "", ""],
                    ["", "", "", "", "", "", ""],
                    ["", "", "", "", "", "", ""],
                    ["Diajukan oleh:", "", "", "", "Disetujui oleh:", "", ""],
                    ["", "", "", "", "", "", ""],
                    ["", "", "", "", "", "", ""],
                    [f"({employee_name})", "", "", "", "(___________________)", "", ""],
                    ["Tanggal:", "", "", "", "Tanggal:", "", ""],
                ],
                "column_widths": {
                    "A": 5,
                    "B": 12,
                    "C": 15,
                    "D": 35,
                    "E": 15,
                    "F": 12,
                    "G": 25
                },
                "formatting": {
                    "currency_columns": ["E"],
                    "date_columns": ["B"],
                    "header_row": 7
                }
            }]
        }
        
        return structure
    
    @staticmethod
    def get_categories() -> List[str]:
        """Get expense categories"""
        return [
            "Transport",
            "Makan",
            "Akomodasi",
            "Supplies",
            "Entertainment",
            "Komunikasi",
            "Parkir & Tol",
            "Lainnya"
        ]
