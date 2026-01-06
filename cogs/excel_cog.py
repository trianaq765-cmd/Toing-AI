"""
Excel Cog - Main Excel manipulation commands
Buat, edit, perbaiki, analisis file Excel
"""

import logging
import re
from typing import Optional
import json

import discord
from discord.ext import commands

from config import EMOJIS, COLORS
from utils.formatters import Formatters
from utils.validators import Validators
from utils.helpers import Helpers

logger = logging.getLogger("office_bot.excel")

# =============================================================================
# HELPER FUNCTION - EXTRACT JSON
# =============================================================================

def extract_json_from_response(response_text: str) -> dict:
    """
    Extract JSON from AI response, handling markdown code blocks and extra text
    """
    if not response_text:
        raise ValueError("Empty response from AI")
    
    text = response_text.strip()
    
    # Method 1: Try direct JSON parse
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    
    # Method 2: Extract from ```json ... ``` code block
    json_match = re.search(r'```json\s*([\s\S]*?)\s*```', text)
    if json_match:
        try:
            return json.loads(json_match.group(1).strip())
        except json.JSONDecodeError:
            pass
    
    # Method 3: Extract from ``` ... ``` code block (without json tag)
    code_match = re.search(r'```\s*([\s\S]*?)\s*```', text)
    if code_match:
        try:
            return json.loads(code_match.group(1).strip())
        except json.JSONDecodeError:
            pass
    
    # Method 4: Find JSON object pattern { ... }
    json_obj_match = re.search(r'\{[\s\S]*\}', text)
    if json_obj_match:
        try:
            return json.loads(json_obj_match.group(0))
        except json.JSONDecodeError:
            pass
    
    # Method 5: Find JSON array pattern [ ... ]
    json_arr_match = re.search(r'\[[\s\S]*\]', text)
    if json_arr_match:
        try:
            return json.loads(json_arr_match.group(0))
        except json.JSONDecodeError:
            pass
    
    # If all methods fail, raise error with helpful message
    raise ValueError(f"Could not extract valid JSON from response. First 200 chars: {text[:200]}")


def create_default_structure(description: str) -> dict:
    """
    Create a default Excel structure when AI fails
    """
    return {
        "sheets": [{
            "name": "Sheet1",
            "headers": ["No", "Data 1", "Data 2", "Data 3", "Keterangan"],
            "data": [
                [1, "", "", "", ""],
                [2, "", "", "", ""],
                [3, "", "", "", ""],
                [4, "", "", "", ""],
                [5, "", "", "", ""],
            ],
            "column_widths": {"A": 5, "B": 20, "C": 20, "D": 20, "E": 30},
            "formatting": {}
        }]
    }

# =============================================================================
# EXCEL COG
# =============================================================================

class ExcelCog(commands.Cog):
    """
    Main Excel commands - create, fix, analyze, formula
    """
    
    def __init__(self, bot):
        self.bot = bot
        logger.info("Excel Cog loaded")
    
    # =========================================================================
    # CREATE COMMANDS
    # =========================================================================
    
    @commands.command(name="buat", aliases=["create", "buatexcel"])
    @commands.cooldown(1, 10, commands.BucketType.user)
    async def create_excel(self, ctx, *, description: str):
        """
        Buat file Excel dari deskripsi
        
        Usage: !buat tabel penjualan dengan kolom: tanggal, produk, qty, harga, total
        """
        loading = await ctx.send(embed=Helpers.create_loading_embed("Membuat Excel... (AI sedang memproses)"))
        
        try:
            # Get AI to create structure
            ai_response = await self.bot.ai_engine.create_excel_structure(description)
            
            if not ai_response.success:
                raise Exception(f"AI Error: {ai_response.error}")
            
            # Log raw response for debugging
            logger.info(f"AI Response (first 500 chars): {ai_response.content[:500]}")
            
            # Parse JSON response with robust extraction
            try:
                structure = extract_json_from_response(ai_response.content)
            except ValueError as e:
                logger.warning(f"JSON extraction failed: {e}")
                # Use default structure
                structure = create_default_structure(description)
                await ctx.send(f"⚠️ AI response tidak sempurna, menggunakan template default. Silakan edit sesuai kebutuhan.")
            
            # Validate structure
            if "sheets" not in structure:
                structure = {"sheets": [structure]} if isinstance(structure, dict) else create_default_structure(description)
            
            # Create Excel file
            file_path = await self.bot.excel_engine.create_excel(structure)
            
            # Get file info
            file_info = self.bot.file_handler.get_file_info(file_path)
            
            # Create success embed
            embed = Helpers.create_success_embed(
                title="Excel Berhasil Dibuat",
                description=f"**File:** {file_info['name']}\n"
                           f"**Size:** {Formatters.format_file_size(file_info['size'])}\n"
                           f"**Sheets:** {len(structure.get('sheets', []))}"
            )
            
            # Add preview
            if structure.get('sheets'):
                first_sheet = structure['sheets'][0]
                preview = f"**Sheet: {first_sheet.get('name', 'Sheet1')}**\n"
                
                headers = first_sheet.get('headers', [])
                if headers:
                    preview += f"Columns: {', '.join(str(h) for h in headers[:5])}"
                    if len(headers) > 5:
                        preview += f" ... (+{len(headers)-5} more)"
                
                embed.add_field(
                    name=f"{EMOJIS['excel']} Preview",
                    value=preview,
                    inline=False
                )
            
            # Send file
            discord_file = self.bot.file_handler.create_discord_file(file_path)
            await loading.delete()
            await ctx.send(embed=embed, file=discord_file)
            
            # Log usage
            await self.bot.log_usage(ctx.author.id, "buat", "excel", True)
            
            # Cleanup
            await self.bot.file_handler.delete_file(file_path)
            
        except Exception as e:
            logger.error(f"Error creating Excel: {e}")
            await loading.delete()
            embed = Helpers.create_error_embed(
                title="Error Membuat Excel",
                description=f"```{str(e)[:500]}```\n\n"
                           f"**Tips:** Coba perintah yang lebih sederhana, contoh:\n"
                           f"`!buat tabel karyawan dengan kolom NIK, Nama, Gaji`"
            )
            await ctx.send(embed=embed)
    
    # =========================================================================
    # FIX COMMANDS
    # =========================================================================
    
    @commands.command(name="perbaiki", aliases=["fix"])
    @commands.cooldown(1, 10, commands.BucketType.user)
    async def fix_excel(self, ctx):
        """
        Perbaiki error dalam file Excel
        
        Usage: Upload Excel file + !perbaiki
        """
        # Check for attachment
        if not ctx.message.attachments:
            embed = Helpers.create_error_embed(
                title="File Tidak Ditemukan",
                description=f"Upload file Excel bersamaan dengan command ini!\n\n"
                           f"**Contoh:**\n"
                           f"1. Upload file Excel\n"
                           f"2. Ketik: `!perbaiki` di caption"
            )
            await ctx.send(embed=embed)
            return
        
        loading = await ctx.send(embed=Helpers.create_loading_embed("Memeriksa dan memperbaiki Excel..."))
        
        try:
            # Download attachment
            attachment = ctx.message.attachments[0]
            
            # Validate
            if not attachment.filename.endswith(('.xlsx', '.xls', '.xlsm')):
                raise ValueError("File harus berformat Excel (.xlsx, .xls, .xlsm)")
            
            file_path = await self.bot.file_handler.download_attachment(
                attachment,
                subfolder="fix"
            )
            
            # Read Excel and detect errors
            excel_data = await self.bot.excel_engine.read_excel(file_path)
            errors = excel_data.get('errors', [])
            
            if not errors:
                embed = Helpers.create_success_embed(
                    title="Tidak Ada Error",
                    description=f"{EMOJIS['check']} File Excel tidak memiliki error!"
                )
                
                # Show stats
                stats = excel_data.get('stats')
                if stats:
                    embed.add_field(
                        name="📊 Info File",
                        value=f"Sheets: {stats.total_sheets}\n"
                              f"Rows: {stats.total_rows:,}\n"
                              f"Columns: {stats.total_columns}\n"
                              f"Formulas: {'Ya' if stats.has_formulas else 'Tidak'}",
                        inline=False
                    )
                
                await loading.delete()
                await ctx.send(embed=embed)
                await self.bot.file_handler.delete_file(file_path)
                return
            
            # Show errors found
            errors_text = f"Ditemukan **{len(errors)} error**:\n\n"
            for i, error in enumerate(errors[:10], 1):
                errors_text += f"{i}. **{error.cell}**: {error.error_type}\n"
                errors_text += f"   _{error.message}_\n\n"
            
            if len(errors) > 10:
                errors_text += f"_...dan {len(errors)-10} error lainnya_\n"
            
            # Get AI suggestions
            errors_description = "\n".join([
                f"Cell {e.cell}: {e.error_type} - Formula: {e.formula}"
                for e in errors[:5]
            ])
            
            ai_response = await self.bot.ai_engine.fix_excel_errors(errors_description)
            
            # Try auto-fix
            fixed_path, original_errors = await self.bot.excel_engine.fix_excel_errors(
                file_path,
                auto_fix=True
            )
            
            # Create result embed
            embed = Helpers.create_success_embed(
                title="Excel Berhasil Diperbaiki",
                description=errors_text[:1000]
            )
            
            # Add AI suggestions
            if ai_response.success:
                suggestions = Helpers.truncate(ai_response.content, 1024)
                embed.add_field(
                    name=f"{EMOJIS['ai']} Saran Perbaikan",
                    value=suggestions,
                    inline=False
                )
            
            embed.add_field(
                name="ℹ️ Catatan",
                value="File telah diperbaiki secara otomatis untuk error umum.\n"
                      "Silakan cek hasil dan sesuaikan jika perlu.",
                inline=False
            )
            
            # Send fixed file
            discord_file = self.bot.file_handler.create_discord_file(
                fixed_path,
                filename=f"fixed_{attachment.filename}"
            )
            
            await loading.delete()
            await ctx.send(embed=embed, file=discord_file)
            
            # Log usage
            await self.bot.log_usage(ctx.author.id, "perbaiki", "excel", True)
            
            # Cleanup
            await self.bot.file_handler.delete_file(file_path)
            await self.bot.file_handler.delete_file(fixed_path)
            
        except Exception as e:
            logger.error(f"Error fixing Excel: {e}")
            await loading.delete()
            embed = Helpers.create_error_embed(
                title="Error Memperbaiki Excel",
                description=f"```{str(e)[:500]}```"
            )
            await ctx.send(embed=embed)
    
    # =========================================================================
    # FORMULA COMMANDS
    # =========================================================================
    
    @commands.command(name="rumus", aliases=["formula"])
    @commands.cooldown(1, 3, commands.BucketType.user)
    async def explain_formula(self, ctx, *, formula_name: str):
        """
        Jelaskan rumus Excel dan berikan contoh
        
        Usage: !rumus VLOOKUP
               !rumus SUMIFS
        """
        loading = await ctx.send(embed=Helpers.create_loading_embed("Mencari rumus..."))
        
        try:
            # Check if formula exists in library
            from core.formula_library import FormulaLibrary
            
            formula_lib = FormulaLibrary()
            formula_info = formula_lib.get_formula(formula_name.upper())
            
            if formula_info:
                # Use library data
                embed = discord.Embed(
                    title=f"📐 {formula_info['name']}",
                    description=formula_info['description'],
                    color=COLORS["excel"]
                )
                
                embed.add_field(
                    name="📝 Syntax",
                    value=f"```excel\n{formula_info['syntax']}\n```",
                    inline=False
                )
                
                if formula_info.get('parameters'):
                    params_text = "\n".join([
                        f"• **{p['name']}**: {p['description']}"
                        for p in formula_info['parameters']
                    ])
                    embed.add_field(
                        name="📋 Parameters",
                        value=params_text,
                        inline=False
                    )
                
                if formula_info.get('examples'):
                    examples_text = "\n\n".join([
                        f"**{ex['title']}**\n```excel\n{ex['formula']}\n```\n_{ex['explanation']}_"
                        for ex in formula_info['examples'][:2]
                    ])
                    embed.add_field(
                        name="💡 Contoh Penggunaan",
                        value=Helpers.truncate(examples_text, 1024),
                        inline=False
                    )
                
                if formula_info.get('tips'):
                    embed.add_field(
                        name="⭐ Tips",
                        value=formula_info['tips'],
                        inline=False
                    )
                
            else:
                # Use AI for explanation
                ai_response = await self.bot.ai_engine.explain_formula(formula_name)
                
                if not ai_response.success:
                    raise Exception(ai_response.error)
                
                embed = discord.Embed(
                    title=f"📐 Rumus: {formula_name.upper()}",
                    description=Helpers.truncate(ai_response.content, 4000),
                    color=COLORS["excel"]
                )
            
            embed.set_footer(text="Gunakan !daftarrumus untuk melihat semua rumus yang tersedia")
            
            await loading.delete()
            await ctx.send(embed=embed)
            
            await self.bot.log_usage(ctx.author.id, "rumus", "excel", True)
            
        except Exception as e:
            logger.error(f"Error explaining formula: {e}")
            await loading.delete()
            embed = Helpers.create_error_embed(
                title="Error Mencari Rumus",
                description=f"```{str(e)[:500]}```\n\n"
                           f"Coba gunakan: `!daftarrumus` untuk melihat rumus yang tersedia"
            )
            await ctx.send(embed=embed)
    
    @commands.command(name="daftarrumus", aliases=["listformulas", "formulas"])
    async def list_formulas(self, ctx, category: Optional[str] = None):
        """
        Tampilkan daftar rumus Excel
        
        Usage: !daftarrumus
               !daftarrumus lookup
        """
        try:
            from core.formula_library import FormulaLibrary
            
            formula_lib = FormulaLibrary()
            
            if category:
                # Show formulas in category
                category = category.lower()
                formulas = formula_lib.get_formulas_by_category(category)
                
                if not formulas:
                    available_cats = ", ".join(formula_lib.get_categories())
                    embed = Helpers.create_error_embed(
                        title="Kategori Tidak Ditemukan",
                        description=f"Kategori yang tersedia:\n{available_cats}"
                    )
                    await ctx.send(embed=embed)
                    return
                
                embed = discord.Embed(
                    title=f"📚 Rumus Excel: {category.upper()}",
                    color=COLORS["excel"]
                )
                
                formulas_text = "\n".join([
                    f"• **{f['name']}** - {f['description'][:50]}..."
                    for f in formulas[:20]
                ])
                
                embed.description = formulas_text
                
            else:
                # Show categories
                categories = formula_lib.get_categories_with_count()
                
                embed = discord.Embed(
                    title="📚 Kategori Rumus Excel",
                    description="Gunakan `!daftarrumus <kategori>` untuk melihat rumus dalam kategori",
                    color=COLORS["excel"]
                )
                
                for cat_name, count in categories.items():
                    emoji = self._get_category_emoji(cat_name)
                    embed.add_field(
                        name=f"{emoji} {cat_name.upper()}",
                        value=f"{count} rumus",
                        inline=True
                    )
            
            embed.set_footer(text="Gunakan !rumus <nama> untuk detail rumus")
            await ctx.send(embed=embed)
            
        except Exception as e:
            logger.error(f"Error listing formulas: {e}")
            embed = Helpers.create_error_embed(
                title="Error Menampilkan Rumus",
                description=f"```{str(e)[:500]}```"
            )
            await ctx.send(embed=embed)
    
    # =========================================================================
    # TEMPLATE COMMANDS
    # =========================================================================
    
    @commands.command(name="template")
    @commands.cooldown(1, 5, commands.BucketType.user)
    async def get_template(self, ctx, template_type: str):
        """
        Dapatkan template Excel siap pakai
        
        Usage: !template invoice
               !template payroll
               !template expense
               !template attendance
        """
        loading = await ctx.send(embed=Helpers.create_loading_embed("Membuat template..."))
        
        try:
            # Create template
            file_path = await self.bot.excel_engine.create_template(template_type.lower())
            
            # Template info
            template_info = {
                "invoice": "Template invoice/faktur dengan PPN 11%",
                "payroll": "Template slip gaji dengan BPJS dan PPh21",
                "expense": "Template laporan pengeluaran",
                "attendance": "Template absensi karyawan"
            }
            
            embed = Helpers.create_success_embed(
                title=f"Template {template_type.upper()}",
                description=template_info.get(
                    template_type.lower(),
                    "Template Excel siap pakai"
                )
            )
            
            embed.add_field(
                name="ℹ️ Cara Pakai",
                value="1. Download template\n"
                      "2. Isi data sesuai kebutuhan\n"
                      "3. Rumus sudah otomatis tersedia",
                inline=False
            )
            
            # Send file
            discord_file = self.bot.file_handler.create_discord_file(file_path)
            await loading.delete()
            await ctx.send(embed=embed, file=discord_file)
            
            await self.bot.log_usage(ctx.author.id, "template", "excel", True)
            await self.bot.file_handler.delete_file(file_path)
            
        except Exception as e:
            logger.error(f"Error getting template: {e}")
            await loading.delete()
            
            embed = Helpers.create_error_embed(
                title="Error Membuat Template",
                description=f"Template `{template_type}` tidak tersedia.\n\n"
                           f"**Template yang tersedia:**\n"
                           f"• invoice\n"
                           f"• payroll\n"
                           f"• expense\n"
                           f"• attendance"
            )
            await ctx.send(embed=embed)
    
    # =========================================================================
    # INFO COMMANDS
    # =========================================================================
    
    @commands.command(name="infoexcel", aliases=["excelinfo"])
    @commands.cooldown(1, 5, commands.BucketType.user)
    async def excel_info(self, ctx):
        """
        Tampilkan info file Excel
        
        Usage: Upload Excel + !infoexcel
        """
        if not ctx.message.attachments:
            embed = Helpers.create_error_embed(
                title="File Tidak Ditemukan",
                description="Upload file Excel bersamaan dengan command ini!"
            )
            await ctx.send(embed=embed)
            return
        
        loading = await ctx.send(embed=Helpers.create_loading_embed("Membaca info Excel..."))
        
        try:
            attachment = ctx.message.attachments[0]
            file_path = await self.bot.file_handler.download_attachment(attachment)
            
            # Read Excel
            excel_data = await self.bot.excel_engine.read_excel(file_path)
            stats = excel_data['stats']
            
            # Create info embed
            embed = discord.Embed(
                title=f"{EMOJIS['excel']} Info Excel: {attachment.filename}",
                color=COLORS["excel"]
            )
            
            # Basic info
            embed.add_field(
                name="📊 Basic Info",
                value=f"File Size: {Formatters.format_file_size(excel_data['file_size'])}\n"
                      f"Total Sheets: {stats.total_sheets}\n"
                      f"Active Sheet: {excel_data['active_sheet']}",
                inline=False
            )
            
            # Sheets list
            sheets_text = "\n".join([f"• {sheet}" for sheet in excel_data['sheets']])
            embed.add_field(
                name="📑 Sheets",
                value=sheets_text,
                inline=True
            )
            
            # Stats
            embed.add_field(
                name="📈 Statistics",
                value=f"Total Rows: {stats.total_rows:,}\n"
                      f"Total Columns: {stats.total_columns}\n"
                      f"Has Formulas: {'Ya' if stats.has_formulas else 'Tidak'}",
                inline=True
            )
            
            # Errors
            if excel_data['errors']:
                errors_text = f"Ditemukan {len(excel_data['errors'])} error:\n"
                error_types = {}
                for error in excel_data['errors']:
                    error_types[error.error_type] = error_types.get(error.error_type, 0) + 1
                
                for error_type, count in error_types.items():
                    errors_text += f"• {error_type}: {count}\n"
                
                embed.add_field(
                    name=f"{EMOJIS['warning']} Errors",
                    value=errors_text,
                    inline=False
                )
            else:
                embed.add_field(
                    name=f"{EMOJIS['success']} Status",
                    value="Tidak ada error terdeteksi",
                    inline=False
                )
            
            await loading.delete()
            await ctx.send(embed=embed)
            
            await self.bot.log_usage(ctx.author.id, "infoexcel", "excel", True)
            await self.bot.file_handler.delete_file(file_path)
            
        except Exception as e:
            logger.error(f"Error getting Excel info: {e}")
            await loading.delete()
            embed = Helpers.create_error_embed(
                title="Error Membaca Info",
                description=f"```{str(e)[:500]}```"
            )
            await ctx.send(embed=embed)
    
    # =========================================================================
    # HELPER METHODS
    # =========================================================================
    
    def _get_category_emoji(self, category: str) -> str:
        """Get emoji for formula category"""
        emojis = {
            "lookup": "🔍",
            "math": "➕",
            "statistical": "📊",
            "text": "📝",
            "date": "📅",
            "logical": "🔀",
            "financial": "💰",
            "array": "🔢",
            "database": "🗄️"
        }
        return emojis.get(category.lower(), "📐")

# =============================================================================
# SETUP
# =============================================================================

async def setup(bot):
    await bot.add_cog(ExcelCog(bot))
