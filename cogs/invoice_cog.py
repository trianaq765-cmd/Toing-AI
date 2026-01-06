"""
Invoice Cog - Generate & manage invoices, quotations, PO, etc.
"""

import logging
import re
from typing import Optional
from datetime import datetime, timedelta

import discord
from discord.ext import commands
from discord import app_commands

from config import EMOJIS, COLORS
from utils.formatters import Formatters
from utils.validators import Validators
from utils.helpers import Helpers

logger = logging.getLogger("office_bot.invoice")

# =============================================================================
# HELPER FUNCTION - EXTRACT JSON
# =============================================================================

def extract_json_from_response(response_text: str) -> dict:
    """Extract JSON from AI response"""
    import json
    
    if not response_text:
        raise ValueError("Empty response from AI")
    
    text = response_text.strip()
    
    # Method 1: Direct parse
    try:
        return json.loads(text)
    except:
        pass
    
    # Method 2: Extract from ```json ... ```
    json_match = re.search(r'```json\s*([\s\S]*?)\s*```', text)
    if json_match:
        try:
            return json.loads(json_match.group(1).strip())
        except:
            pass
    
    # Method 3: Extract from ``` ... ```
    code_match = re.search(r'```\s*([\s\S]*?)\s*```', text)
    if code_match:
        try:
            return json.loads(code_match.group(1).strip())
        except:
            pass
    
    # Method 4: Find { ... }
    json_obj_match = re.search(r'\{[\s\S]*\}', text)
    if json_obj_match:
        try:
            return json.loads(json_obj_match.group(0))
        except:
            pass
    
    raise ValueError("Could not extract valid JSON")


def create_default_invoice(client_name: str = "Client") -> dict:
    """Create default invoice structure"""
    today = datetime.now()
    due_date = today + timedelta(days=30)
    
    return {
        "company_name": client_name,
        "invoice_number": f"INV/{today.year}/{today.month:02d}/001",
        "date": today.strftime("%d/%m/%Y"),
        "due_date": due_date.strftime("%d/%m/%Y"),
        "items": [
            {"description": "Item 1", "qty": 1, "price": 0},
            {"description": "Item 2", "qty": 1, "price": 0},
            {"description": "Item 3", "qty": 1, "price": 0},
        ],
        "subtotal": 0,
        "tax": 0.11,
        "total": 0
    }

# =============================================================================
# INVOICE COG
# =============================================================================

class InvoiceCog(commands.Cog):
    """
    Commands untuk Invoice, Quotation, PO, dan dokumen bisnis
    """
    
    def __init__(self, bot):
        self.bot = bot
        logger.info("Invoice Cog loaded")
    
    # =========================================================================
    # INVOICE COMMANDS
    # =========================================================================
    
    @commands.command(name="invoice", aliases=["inv"])
    @commands.cooldown(1, 5, commands.BucketType.user)
    async def create_invoice(self, ctx, *, details: str):
        """
        Buat invoice otomatis
        
        Usage: !invoice PT ABC, 3 item, total 15jt
        """
        loading = await ctx.send(embed=Helpers.create_loading_embed("Membuat invoice..."))
        
        try:
            # Parse dengan AI
            ai_response = await self.bot.ai_engine.generate_invoice(
                f"""Buatkan struktur invoice dari deskripsi berikut:
                
{details}

Berikan output JSON dengan format:
{{"company_name": "...", "invoice_number": "INV/YYYY/MM/XXX", "date": "DD/MM/YYYY", "due_date": "DD/MM/YYYY", "items": [{{"description": "...", "qty": 1, "price": 100000}}], "subtotal": 0, "tax": 0.11, "total": 0}}

PENTING: Output HANYA JSON, tanpa penjelasan tambahan."""
            )
            
            if not ai_response.success:
                raise Exception(ai_response.error)
            
            # Parse JSON response
            try:
                invoice_data = extract_json_from_response(ai_response.content)
            except ValueError as e:
                logger.warning(f"JSON extraction failed for invoice: {e}")
                invoice_data = create_default_invoice(details.split(',')[0] if ',' in details else details)
                await ctx.send("⚠️ Menggunakan template default. Silakan edit sesuai kebutuhan.")
            
            # Calculate totals if not present
            if 'items' in invoice_data:
                subtotal = sum(item.get('qty', 1) * item.get('price', 0) for item in invoice_data['items'])
                invoice_data['subtotal'] = subtotal
                tax_rate = invoice_data.get('tax', 0.11)
                invoice_data['tax_amount'] = subtotal * tax_rate
                invoice_data['total'] = subtotal + invoice_data['tax_amount']
            
            # Generate Excel structure
            structure = await self._create_invoice_structure(invoice_data)
            
            # Create Excel file
            file_path = await self.bot.excel_engine.create_excel(structure)
            
            # Create embed
            embed = Helpers.create_success_embed(
                title="Invoice Berhasil Dibuat",
                description=f"**{invoice_data.get('company_name', 'N/A')}**\n"
                           f"Invoice: {invoice_data.get('invoice_number', 'N/A')}\n"
                           f"Total: {Formatters.format_rupiah(invoice_data.get('total', 0))}"
            )
            
            embed.add_field(
                name="Detail",
                value=f"Tanggal: {invoice_data.get('date', 'N/A')}\n"
                      f"Jatuh Tempo: {invoice_data.get('due_date', 'N/A')}\n"
                      f"Item: {len(invoice_data.get('items', []))} item",
                inline=False
            )
            
            # Send file
            discord_file = self.bot.file_handler.create_discord_file(file_path)
            await loading.delete()
            await ctx.send(embed=embed, file=discord_file)
            
            # Log usage
            await self.bot.log_usage(ctx.author.id, "invoice", "invoice", True)
            
            # Cleanup
            await self.bot.file_handler.delete_file(file_path)
            
        except Exception as e:
            logger.error(f"Error creating invoice: {e}")
            await loading.delete()
            embed = Helpers.create_error_embed(
                title="Error Membuat Invoice",
                description=f"```{str(e)[:500]}```\n\n"
                           f"**Tips:** Coba format seperti:\n"
                           f"`!invoice PT ABC, laptop 5 unit @10jt, mouse 10 unit @200rb`"
            )
            await ctx.send(embed=embed)
    
    @commands.command(name="quotation", aliases=["quote"])
    @commands.cooldown(1, 5, commands.BucketType.user)
    async def create_quotation(self, ctx, *, details: str):
        """
        Buat quotation/penawaran harga
        
        Usage: !quotation untuk PT XYZ, software development, 3 bulan
        """
        loading = await ctx.send(embed=Helpers.create_loading_embed("Membuat quotation..."))
        
        try:
            ai_response = await self.bot.ai_engine.generate_invoice(
                f"""Buatkan quotation/penawaran harga dari: {details}

Format output JSON:
{{"quote_number": "QT/YYYY/MM/XXX", "client_name": "...", "valid_until": "DD/MM/YYYY", "items": [{{"description": "...", "qty": 1, "unit": "bulan", "price": 100000}}]}}"""
            )
            
            if not ai_response.success:
                raise Exception(ai_response.error)
            
            try:
                quote_data = extract_json_from_response(ai_response.content)
            except ValueError:
                quote_data = {
                    "quote_number": f"QT/{datetime.now().year}/{datetime.now().month:02d}/001",
                    "client_name": details.split(',')[0] if ',' in details else details,
                    "valid_until": (datetime.now() + timedelta(days=30)).strftime("%d/%m/%Y"),
                    "items": [{"description": "Item 1", "qty": 1, "unit": "unit", "price": 0}]
                }
            
            # Create structure
            structure = await self._create_quotation_structure(quote_data)
            file_path = await self.bot.excel_engine.create_excel(structure)
            
            # Send
            embed = Helpers.create_success_embed(
                title="Quotation Berhasil Dibuat",
                description=f"**{quote_data.get('client_name', 'N/A')}**\n"
                           f"Quote: {quote_data.get('quote_number', 'N/A')}"
            )
            
            discord_file = self.bot.file_handler.create_discord_file(file_path)
            await loading.delete()
            await ctx.send(embed=embed, file=discord_file)
            
            await self.bot.log_usage(ctx.author.id, "quotation", "invoice", True)
            await self.bot.file_handler.delete_file(file_path)
            
        except Exception as e:
            logger.error(f"Error creating quotation: {e}")
            await loading.delete()
            embed = Helpers.create_error_embed(
                title="Error Membuat Quotation",
                description=f"```{str(e)[:500]}```"
            )
            await ctx.send(embed=embed)
    
    @commands.command(name="po", aliases=["purchase_order"])
    @commands.cooldown(1, 5, commands.BucketType.user)
    async def create_po(self, ctx, *, details: str):
        """
        Buat Purchase Order
        
        Usage: !po ke PT Supplier, 100 unit produk A
        """
        loading = await ctx.send(embed=Helpers.create_loading_embed("Membuat Purchase Order..."))
        
        try:
            ai_response = await self.bot.ai_engine.generate_invoice(
                f"Buatkan Purchase Order (PO) dari: {details}"
            )
            
            if not ai_response.success:
                raise Exception(ai_response.error)
            
            try:
                po_data = extract_json_from_response(ai_response.content)
            except ValueError:
                po_data = {
                    "po_number": f"PO/{datetime.now().year}/{datetime.now().month:02d}/001",
                    "supplier_name": details.split(',')[0] if ',' in details else details,
                    "date": datetime.now().strftime("%d/%m/%Y"),
                    "items": [{"description": "Item 1", "qty": 1, "unit": "unit"}]
                }
            
            structure = await self._create_po_structure(po_data)
            file_path = await self.bot.excel_engine.create_excel(structure)
            
            embed = Helpers.create_success_embed(
                title="Purchase Order Berhasil Dibuat",
                description=f"PO Number: {po_data.get('po_number', 'N/A')}"
            )
            
            discord_file = self.bot.file_handler.create_discord_file(file_path)
            await loading.delete()
            await ctx.send(embed=embed, file=discord_file)
            
            await self.bot.log_usage(ctx.author.id, "po", "invoice", True)
            await self.bot.file_handler.delete_file(file_path)
            
        except Exception as e:
            logger.error(f"Error creating PO: {e}")
            await loading.delete()
            embed = Helpers.create_error_embed(
                title="Error Membuat PO",
                description=f"```{str(e)[:500]}```"
            )
            await ctx.send(embed=embed)
    
    # =========================================================================
    # HELPER METHODS
    # =========================================================================
    
    async def _create_invoice_structure(self, data: dict) -> dict:
        """Create invoice Excel structure"""
        items = data.get('items', [])
        subtotal = data.get('subtotal', 0)
        tax_rate = data.get('tax', 0.11)
        tax_amount = data.get('tax_amount', subtotal * tax_rate)
        total = data.get('total', subtotal + tax_amount)
        
        item_rows = []
        for i, item in enumerate(items, 1):
            item_rows.append([
                i,
                item.get('description', ''),
                item.get('qty', 1),
                item.get('price', 0),
                f"=C{10+i}*D{10+i}"
            ])
        
        # Pad to at least 5 rows
        while len(item_rows) < 5:
            item_rows.append(["", "", "", "", ""])
        
        last_item_row = 10 + len(items)
        
        return {
            "sheets": [{
                "name": "Invoice",
                "data": [
                    [data.get('company_name', 'INVOICE'), "", "", "", ""],
                    ["", "", "", "", ""],
                    ["INVOICE", "", "", "", ""],
                    ["", "", "", "", ""],
                    [f"No: {data.get('invoice_number', '')}", "", "", f"Tanggal: {data.get('date', '')}", ""],
                    [f"Kepada:", "", "", f"Jatuh Tempo: {data.get('due_date', '')}", ""],
                    [data.get('client_name', data.get('company_name', '')), "", "", "", ""],
                    ["", "", "", "", ""],
                    ["", "", "", "", ""],
                    ["No", "Deskripsi", "Qty", "Harga Satuan", "Total"],
                ] + item_rows + [
                    ["", "", "", "", ""],
                    ["", "", "", "SUBTOTAL:", f"=SUM(E11:E{last_item_row})"],
                    ["", "", "", f"PPN {int(tax_rate*100)}%:", f"=E{last_item_row+2}*{tax_rate}"],
                    ["", "", "", "TOTAL:", f"=E{last_item_row+2}+E{last_item_row+3}"],
                ],
                "column_widths": {"A": 6, "B": 40, "C": 8, "D": 18, "E": 18},
                "formatting": {"currency_columns": ["D", "E"]}
            }]
        }
    
    async def _create_quotation_structure(self, data: dict) -> dict:
        """Create quotation Excel structure"""
        items = data.get('items', [])
        
        item_rows = []
        for i, item in enumerate(items, 1):
            item_rows.append([
                i,
                item.get('description', ''),
                item.get('qty', 1),
                item.get('unit', 'unit'),
                item.get('price', 0),
                f"=C{8+i}*E{8+i}"
            ])
        
        while len(item_rows) < 5:
            item_rows.append(["", "", "", "", "", ""])
        
        return {
            "sheets": [{
                "name": "Quotation",
                "data": [
                    ["QUOTATION / PENAWARAN HARGA", "", "", "", "", ""],
                    ["", "", "", "", "", ""],
                    [f"No: {data.get('quote_number', '')}", "", "", "", f"Valid Until: {data.get('valid_until', '')}", ""],
                    [f"Kepada: {data.get('client_name', '')}", "", "", "", "", ""],
                    ["", "", "", "", "", ""],
                    ["", "", "", "", "", ""],
                    ["No", "Deskripsi", "Qty", "Unit", "Harga Satuan", "Total"],
                ] + item_rows + [
                    ["", "", "", "", "", ""],
                    ["", "", "", "", "TOTAL:", f"=SUM(F8:F{7+len(items)})"],
                ],
                "column_widths": {"A": 5, "B": 35, "C": 8, "D": 10, "E": 15, "F": 15},
                "formatting": {"currency_columns": ["E", "F"]}
            }]
        }
    
    async def _create_po_structure(self, data: dict) -> dict:
        """Create PO Excel structure"""
        items = data.get('items', [])
        
        item_rows = []
        for i, item in enumerate(items, 1):
            item_rows.append([
                i,
                item.get('description', ''),
                item.get('qty', 0),
                item.get('unit', 'unit'),
                item.get('delivery_date', ''),
                item.get('notes', '')
            ])
        
        while len(item_rows) < 5:
            item_rows.append(["", "", "", "", "", ""])
        
        return {
            "sheets": [{
                "name": "Purchase Order",
                "data": [
                    ["PURCHASE ORDER", "", "", "", "", ""],
                    ["", "", "", "", "", ""],
                    [f"PO Number: {data.get('po_number', '')}", "", "", f"Date: {data.get('date', '')}", "", ""],
                    [f"Supplier: {data.get('supplier_name', '')}", "", "", "", "", ""],
                    ["", "", "", "", "", ""],
                    ["No", "Item", "Qty", "Unit", "Delivery Date", "Notes"],
                ] + item_rows,
                "column_widths": {"A": 5, "B": 30, "C": 8, "D": 10, "E": 15, "F": 25}
            }]
        }

# =============================================================================
# SETUP
# =============================================================================

async def setup(bot):
    await bot.add_cog(InvoiceCog(bot))
