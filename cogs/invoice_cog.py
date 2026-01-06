"""
Invoice Cog - Generate & manage invoices, quotations, PO, etc.
"""

import logging
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
{{
    "company_name": "...",
    "invoice_number": "INV/YYYY/MM/XXX",
    "date": "DD/MM/YYYY",
    "due_date": "DD/MM/YYYY",
    "items": [
        {{"description": "...", "qty": 1, "price": 100000}}
    ],
    "subtotal": 0,
    "tax": 0.11,
    "total": 0
}}"""
            )
            
            if not ai_response.success:
                raise Exception(ai_response.error)
            
            # Parse JSON response
            import json
            invoice_data = json.loads(ai_response.content)
            
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
                description=f"```{str(e)[:500]}```"
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
            ai_response = await self.bot.ai_engine.query(
                f"""Buatkan quotation/penawaran harga dari:

{details}

Format output JSON:
{{
    "quote_number": "QT/YYYY/MM/XXX",
    "client_name": "...",
    "valid_until": "DD/MM/YYYY",
    "items": [
        {{"description": "...", "qty": 1, "unit": "bulan", "price": 100000}}
    ]
}}""",
                task_type=self.bot.ai_engine.TaskType.INVOICE
            )
            
            if not ai_response.success:
                raise Exception(ai_response.error)
            
            import json
            quote_data = json.loads(ai_response.content)
            
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
            
            import json
            po_data = json.loads(ai_response.content)
            
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
        
        # Calculate totals
        items = data.get('items', [])
        subtotal = sum(item['qty'] * item['price'] for item in items)
        tax_rate = data.get('tax', 0.11)
        tax_amount = subtotal * tax_rate
        total = subtotal + tax_amount
        
        # Update data
        data['subtotal'] = subtotal
        data['tax_amount'] = tax_amount
        data['total'] = total
        
        # Build rows
        item_rows = []
        for i, item in enumerate(items, 1):
            item_total = item['qty'] * item['price']
            item_rows.append([
                i,
                item['description'],
                item['qty'],
                item['price'],
                f"=C{i+10}*D{i+10}"  # Formula for total
            ])
        
        # Create structure
        return {
            "sheets": [{
                "name": "Invoice",
                "data": [
                    # Header section
                    ["INVOICE", "", "", "", ""],
                    ["", "", "", "", ""],
                    [f"No: {data.get('invoice_number', '')}", "", "", "Tanggal:", data.get('date', '')],
                    [f"Kepada: {data.get('company_name', '')}", "", "", "Jatuh Tempo:", data.get('due_date', '')],
                    ["", "", "", "", ""],
                    ["", "", "", "", ""],
                    # Items header
                    ["No", "Deskripsi", "Qty", "Harga Satuan", "Total"],
                    # Separator row
                    ["", "", "", "", ""],
                    # Items (will be added)
                ] + item_rows + [
                    # Summary
                    ["", "", "", "", ""],
                    ["", "", "", "SUBTOTAL:", f"=SUM(E8:E{7+len(items)})"],
                    ["", "", "", f"PPN {int(tax_rate*100)}%:", f"=E{10+len(items)}*{tax_rate}"],
                    ["", "", "", "TOTAL:", f"=E{10+len(items)}+E{11+len(items)}"],
                    ["", "", "", "", ""],
                    [f"Terbilang: {Formatters.terbilang(total)}", "", "", "", ""],
                ],
                "column_widths": {
                    "A": 5,
                    "B": 40,
                    "C": 8,
                    "D": 15,
                    "E": 15
                },
                "formatting": {
                    "currency_columns": ["D", "E"],
                    "auto_fit": False
                }
            }]
        }
    
    async def _create_quotation_structure(self, data: dict) -> dict:
        """Create quotation Excel structure"""
        
        items = data.get('items', [])
        total = sum(item['qty'] * item['price'] for item in items)
        
        item_rows = []
        for i, item in enumerate(items, 1):
            item_rows.append([
                i,
                item['description'],
                item['qty'],
                item.get('unit', 'unit'),
                item['price'],
                f"=C{i+8}*E{i+8}"
            ])
        
        return {
            "sheets": [{
                "name": "Quotation",
                "data": [
                    ["QUOTATION / PENAWARAN HARGA", "", "", "", "", ""],
                    ["", "", "", "", "", ""],
                    [f"No: {data.get('quote_number', '')}", "", "", "", "Valid Until:", data.get('valid_until', '')],
                    [f"Kepada: {data.get('client_name', '')}", "", "", "", "", ""],
                    ["", "", "", "", "", ""],
                    # Headers
                    ["No", "Deskripsi", "Qty", "Unit", "Harga Satuan", "Total"],
                    ["", "", "", "", "", ""],
                ] + item_rows + [
                    ["", "", "", "", "", ""],
                    ["", "", "", "", "TOTAL:", f"=SUM(F7:F{6+len(items)})"],
                ],
                "column_widths": {
                    "A": 5,
                    "B": 35,
                    "C": 8,
                    "D": 10,
                    "E": 15,
                    "F": 15
                },
                "formatting": {
                    "currency_columns": ["E", "F"]
                }
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
        
        return {
            "sheets": [{
                "name": "Purchase Order",
                "data": [
                    ["PURCHASE ORDER", "", "", "", "", ""],
                    ["", "", "", "", "", ""],
                    [f"PO Number: {data.get('po_number', '')}", "", "", "Date:", data.get('date', '')],
                    [f"Supplier: {data.get('supplier_name', '')}", "", "", "", ""],
                    ["", "", "", "", "", ""],
                    ["No", "Item", "Qty", "Unit", "Delivery Date", "Notes"],
                    ["", "", "", "", "", ""],
                ] + item_rows,
                "column_widths": {
                    "A": 5,
                    "B": 30,
                    "C": 8,
                    "D": 10,
                    "E": 15,
                    "F": 25
                }
            }]
        }

# =============================================================================
# SETUP
# =============================================================================

async def setup(bot):
    await bot.add_cog(InvoiceCog(bot))
