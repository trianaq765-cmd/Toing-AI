"""
Invoice Template - Template untuk invoice/faktur Indonesia
"""

from datetime import datetime, timedelta
from typing import Dict, Any, Optional

class InvoiceTemplate:
    """
    Invoice/Faktur template dengan format Indonesia
    Include PPN 11%
    """
    
    @staticmethod
    def get_structure(
        company_name: str = "PT. NAMA PERUSAHAAN",
        invoice_number: Optional[str] = None,
        client_name: str = "NAMA CLIENT",
        items: Optional[list] = None,
        include_ppn: bool = True
    ) -> Dict[str, Any]:
        """
        Generate invoice structure
        
        Args:
            company_name: Nama perusahaan pengirim invoice
            invoice_number: Nomor invoice (auto-generate jika None)
            client_name: Nama client/customer
            items: List of items [{description, qty, price}]
            include_ppn: Include PPN 11% atau tidak
        """
        
        # Auto-generate invoice number
        if invoice_number is None:
            now = datetime.now()
            invoice_number = f"INV/{now.year}/{now.month:02d}/{now.day:02d}-001"
        
        # Default items
        if items is None:
            items = [
                {"description": "Item 1", "qty": 1, "price": 0},
                {"description": "Item 2", "qty": 1, "price": 0},
                {"description": "Item 3", "qty": 1, "price": 0},
            ]
        
        # Dates
        today = datetime.now().strftime("%d/%m/%Y")
        due_date = (datetime.now() + timedelta(days=30)).strftime("%d/%m/%Y")
        
        # Build item rows
        item_rows = []
        for i, item in enumerate(items, 1):
            item_rows.append([
                i,
                item.get("description", ""),
                item.get("qty", 1),
                item.get("price", 0),
                f"=C{9+i}*D{9+i}"  # Total = qty * price
            ])
        
        # Padding rows if less than 10 items
        while len(item_rows) < 10:
            row_num = len(item_rows) + 1
            item_rows.append([
                "",
                "",
                "",
                "",
                ""
            ])
        
        last_item_row = 9 + len([i for i in items if i])
        
        structure = {
            "sheets": [{
                "name": "Invoice",
                "data": [
                    # Header
                    [company_name, "", "", "", ""],
                    ["Jl. Alamat Perusahaan No. 123", "", "", "", ""],
                    ["Telp: (021) 1234567 | Email: info@company.com", "", "", "", ""],
                    ["", "", "", "", ""],
                    # Invoice Info
                    ["INVOICE", "", "", "", ""],
                    ["", "", "", "", ""],
                    [f"No. Invoice: {invoice_number}", "", "", f"Tanggal: {today}", ""],
                    [f"Kepada: {client_name}", "", "", f"Jatuh Tempo: {due_date}", ""],
                    ["", "", "", "", ""],
                    # Table Header
                    ["No", "Deskripsi", "Qty", "Harga Satuan", "Total"],
                ] + item_rows + [
                    # Summary
                    ["", "", "", "", ""],
                    ["", "", "", "SUBTOTAL:", f"=SUM(E11:E{10+len(items)})"],
                ],
                "column_widths": {
                    "A": 6,
                    "B": 40,
                    "C": 8,
                    "D": 18,
                    "E": 18
                },
                "formatting": {
                    "currency_columns": ["D", "E"],
                    "header_row": 10,
                    "auto_fit": False
                }
            }]
        }
        
        # Add PPN if needed
        if include_ppn:
            ppn_row = len(structure["sheets"][0]["data"])
            structure["sheets"][0]["data"].append(
                ["", "", "", "PPN 11%:", f"=E{ppn_row}*0.11"]
            )
            structure["sheets"][0]["data"].append(
                ["", "", "", "TOTAL:", f"=E{ppn_row}+E{ppn_row+1}"]
            )
        
        # Add terbilang and signature
        structure["sheets"][0]["data"].extend([
            ["", "", "", "", ""],
            ["", "", "", "", ""],
            ["Terbilang:", "", "", "", ""],
            ["", "", "", "", ""],
            ["", "", "", "", ""],
            ["", "", "", f"{company_name}", ""],
            ["", "", "", "", ""],
            ["", "", "", "", ""],
            ["", "", "", "(___________________)", ""],
            ["", "", "", "Authorized Signature", ""],
        ])
        
        return structure
    
    @staticmethod
    def get_simple_structure() -> Dict[str, Any]:
        """Get simple empty invoice template"""
        return InvoiceTemplate.get_structure()
