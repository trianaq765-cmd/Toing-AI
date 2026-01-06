"""
Invoice Generator - Generate invoices, quotations, PO
"""

import logging
from typing import Dict, Any, List, Optional
from datetime import datetime, timedelta
from dataclasses import dataclass

logger = logging.getLogger("office_bot.invoice.generator")

@dataclass
class InvoiceItem:
    """Invoice line item"""
    description: str
    quantity: float
    unit_price: float
    discount: float = 0
    
    @property
    def subtotal(self) -> float:
        return self.quantity * self.unit_price * (1 - self.discount)

@dataclass
class Invoice:
    """Invoice data"""
    invoice_number: str
    date: str
    due_date: str
    client_name: str
    client_address: str
    items: List[InvoiceItem]
    tax_rate: float = 0.11  # PPN 11%
    notes: str = ""
    
    @property
    def subtotal(self) -> float:
        return sum(item.subtotal for item in self.items)
    
    @property
    def tax_amount(self) -> float:
        return self.subtotal * self.tax_rate
    
    @property
    def total(self) -> float:
        return self.subtotal + self.tax_amount

class InvoiceGenerator:
    """
    Generate invoices and related documents
    """
    
    def __init__(self):
        self.counter = 0
        logger.info("Invoice Generator initialized")
    
    def generate_invoice_number(self, prefix: str = "INV") -> str:
        """Generate unique invoice number"""
        self.counter += 1
        now = datetime.now()
        return f"{prefix}/{now.year}/{now.month:02d}/{self.counter:04d}"
    
    def create_invoice(
        self,
        client_name: str,
        items: List[Dict[str, Any]],
        client_address: str = "",
        due_days: int = 30,
        tax_rate: float = 0.11,
        notes: str = ""
    ) -> Invoice:
        """
        Create invoice from items
        """
        invoice_items = [
            InvoiceItem(
                description=item.get("description", ""),
                quantity=item.get("quantity", 1),
                unit_price=item.get("unit_price", 0),
                discount=item.get("discount", 0)
            )
            for item in items
        ]
        
        now = datetime.now()
        due_date = now + timedelta(days=due_days)
        
        return Invoice(
            invoice_number=self.generate_invoice_number(),
            date=now.strftime("%d/%m/%Y"),
            due_date=due_date.strftime("%d/%m/%Y"),
            client_name=client_name,
            client_address=client_address,
            items=invoice_items,
            tax_rate=tax_rate,
            notes=notes
        )
    
    def create_quotation(
        self,
        client_name: str,
        items: List[Dict[str, Any]],
        valid_days: int = 30
    ) -> Dict[str, Any]:
        """
        Create quotation
        """
        now = datetime.now()
        valid_until = now + timedelta(days=valid_days)
        
        return {
            "quote_number": self.generate_invoice_number("QT"),
            "date": now.strftime("%d/%m/%Y"),
            "valid_until": valid_until.strftime("%d/%m/%Y"),
            "client_name": client_name,
            "items": items,
            "subtotal": sum(i.get("quantity", 1) * i.get("unit_price", 0) for i in items),
            "notes": f"Quotation berlaku sampai {valid_until.strftime('%d/%m/%Y')}"
        }
    
    def create_purchase_order(
        self,
        supplier_name: str,
        items: List[Dict[str, Any]],
        delivery_date: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Create Purchase Order
        """
        now = datetime.now()
        if delivery_date is None:
            delivery = now + timedelta(days=14)
            delivery_date = delivery.strftime("%d/%m/%Y")
        
        return {
            "po_number": self.generate_invoice_number("PO"),
            "date": now.strftime("%d/%m/%Y"),
            "delivery_date": delivery_date,
            "supplier_name": supplier_name,
            "items": items,
            "total": sum(i.get("quantity", 1) * i.get("unit_price", 0) for i in items)
        }
    
    def invoice_to_excel_structure(self, invoice: Invoice) -> Dict[str, Any]:
        """
        Convert Invoice to Excel structure
        """
        # Build item rows
        item_rows = []
        for i, item in enumerate(invoice.items, 1):
            item_rows.append([
                i,
                item.description,
                item.quantity,
                item.unit_price,
                item.subtotal
            ])
        
        return {
            "sheets": [{
                "name": "Invoice",
                "data": [
                    ["INVOICE", "", "", "", ""],
                    ["", "", "", "", ""],
                    [f"No: {invoice.invoice_number}", "", "Tanggal:", invoice.date, ""],
                    [f"Kepada: {invoice.client_name}", "", "Jatuh Tempo:", invoice.due_date, ""],
                    [invoice.client_address, "", "", "", ""],
                    ["", "", "", "", ""],
                    ["No", "Deskripsi", "Qty", "Harga", "Total"],
                ] + item_rows + [
                    ["", "", "", "", ""],
                    ["", "", "", "Subtotal:", invoice.subtotal],
                    ["", "", "", f"PPN {int(invoice.tax_rate*100)}%:", invoice.tax_amount],
                    ["", "", "", "TOTAL:", invoice.total],
                ],
                "column_widths": {"A": 5, "B": 40, "C": 10, "D": 15, "E": 15},
                "formatting": {"currency_columns": ["D", "E"]}
            }]
        }
