"""
Finance Cog - Financial analysis, accounting, budgeting
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

logger = logging.getLogger("office_bot.finance")

# =============================================================================
# JSON EXTRACTION HELPER
# =============================================================================

def extract_json_from_response(response_text: str) -> dict:
    """Extract JSON from AI response"""
    if not response_text:
        raise ValueError("Empty response")
    
    text = response_text.strip()
    
    # Method 1: Direct parse
    try:
        return json.loads(text)
    except:
        pass
    
    # Method 2: From ```json ... ```
    match = re.search(r'```json\s*([\s\S]*?)\s*```', text)
    if match:
        try:
            return json.loads(match.group(1).strip())
        except:
            pass
    
    # Method 3: From ``` ... ```
    match = re.search(r'```\s*([\s\S]*?)\s*```', text)
    if match:
        try:
            return json.loads(match.group(1).strip())
        except:
            pass
    
    # Method 4: Find { ... }
    match = re.search(r'\{[\s\S]*\}', text)
    if match:
        try:
            return json.loads(match.group(0))
        except:
            pass
    
    raise ValueError("Could not extract JSON")

# =============================================================================
# FINANCE COG
# =============================================================================

class FinanceCog(commands.Cog):
    def __init__(self, bot):
        self.bot = bot
        logger.info("Finance Cog loaded")
    
    # =========================================================================
    # FINANCIAL STATEMENTS
    # =========================================================================
    
    @commands.command(name="neraca", aliases=["balance_sheet"])
    @commands.cooldown(1, 10, commands.BucketType.user)
    async def create_balance_sheet(self, ctx, *, details: str):
        """Buat template Neraca"""
        loading = await ctx.send(embed=Helpers.create_loading_embed("Membuat neraca..."))
        
        try:
            ai_response = await self.bot.ai_engine.finance_analysis(
                f"""Buatkan struktur Neraca untuk: {details}

Output JSON format:
{{"company": "...", "date": "DD/MM/YYYY", "assets": {{"kas": 0, "piutang": 0}}, "liabilities": {{"hutang": 0}}, "equity": {{"modal": 0}}}}"""
            )
            
            if not ai_response.success:
                raise Exception(ai_response.error)
            
            try:
                data = extract_json_from_response(ai_response.content)
            except:
                # Default structure
                data = {
                    "company": details.split()[0] if details else "Perusahaan",
                    "date": "31/12/2024",
                    "assets": {"kas": 0, "piutang": 0, "persediaan": 0, "aset_tetap": 0},
                    "liabilities": {"hutang_usaha": 0, "hutang_bank": 0},
                    "equity": {"modal": 0, "laba_ditahan": 0}
                }
            
            structure = await self._create_balance_sheet_structure(data)
            file_path = await self.bot.excel_engine.create_excel(structure)
            
            embed = Helpers.create_success_embed(
                title="Neraca Berhasil Dibuat",
                description=f"**{data.get('company', 'N/A')}**\nPer: {data.get('date', 'N/A')}"
            )
            
            discord_file = self.bot.file_handler.create_discord_file(file_path)
            await loading.delete()
            await ctx.send(embed=embed, file=discord_file)
            
            await self.bot.log_usage(ctx.author.id, "neraca", "finance", True)
            await self.bot.file_handler.delete_file(file_path)
            
        except Exception as e:
            logger.error(f"Error creating balance sheet: {e}")
            await loading.delete()
            await ctx.send(embed=Helpers.create_error_embed("Error", f"```{str(e)[:500]}```"))
    
    @commands.command(name="labarugi", aliases=["income_statement", "pnl"])
    @commands.cooldown(1, 10, commands.BucketType.user)
    async def create_income_statement(self, ctx, *, details: str):
        """Buat Laporan Laba Rugi"""
        loading = await ctx.send(embed=Helpers.create_loading_embed("Membuat laporan laba rugi..."))
        
        try:
            ai_response = await self.bot.ai_engine.finance_analysis(
                f"""Buatkan struktur Laporan Laba Rugi dari: {details}

Output JSON format:
{{"company": "...", "period": "...", "pendapatan": 0, "hpp": 0, "beban_operasional": {{"gaji": 0, "sewa": 0}}, "laba_bersih": 0}}"""
            )
            
            if not ai_response.success:
                raise Exception(ai_response.error)
            
            try:
                data = extract_json_from_response(ai_response.content)
            except:
                # Parse from details string
                data = self._parse_income_statement_details(details)
            
            structure = await self._create_income_statement_structure(data)
            file_path = await self.bot.excel_engine.create_excel(structure)
            
            # Calculate key metrics
            pendapatan = data.get('pendapatan', 0)
            laba_bersih = data.get('laba_bersih', 0)
            margin = (laba_bersih / pendapatan * 100) if pendapatan > 0 else 0
            
            embed = Helpers.create_success_embed(
                title="Laporan Laba Rugi",
                description=f"**{data.get('company', 'N/A')}**\nPeriode: {data.get('period', 'N/A')}"
            )
            
            embed.add_field(
                name="📊 Ringkasan",
                value=f"Pendapatan: {Formatters.format_rupiah(pendapatan)}\n"
                      f"Laba Bersih: {Formatters.format_rupiah(laba_bersih)}\n"
                      f"Net Margin: {margin:.2f}%",
                inline=False
            )
            
            discord_file = self.bot.file_handler.create_discord_file(file_path)
            await loading.delete()
            await ctx.send(embed=embed, file=discord_file)
            
            await self.bot.log_usage(ctx.author.id, "labarugi", "finance", True)
            await self.bot.file_handler.delete_file(file_path)
            
        except Exception as e:
            logger.error(f"Error creating income statement: {e}")
            await loading.delete()
            await ctx.send(embed=Helpers.create_error_embed("Error", f"```{str(e)[:500]}```"))
    
    @commands.command(name="cashflow", aliases=["arus_kas"])
    @commands.cooldown(1, 10, commands.BucketType.user)
    async def create_cashflow(self, ctx, *, details: str):
        """Buat Laporan Arus Kas"""
        loading = await ctx.send(embed=Helpers.create_loading_embed("Membuat laporan arus kas..."))
        
        try:
            structure = await self._create_cashflow_structure({"company": details.split()[0], "period": "2024"})
            file_path = await self.bot.excel_engine.create_excel(structure)
            
            embed = Helpers.create_success_embed(
                title="Laporan Arus Kas",
                description="Template arus kas berhasil dibuat"
            )
            
            discord_file = self.bot.file_handler.create_discord_file(file_path)
            await loading.delete()
            await ctx.send(embed=embed, file=discord_file)
            
            await self.bot.log_usage(ctx.author.id, "cashflow", "finance", True)
            await self.bot.file_handler.delete_file(file_path)
            
        except Exception as e:
            logger.error(f"Error creating cashflow: {e}")
            await loading.delete()
            await ctx.send(embed=Helpers.create_error_embed("Error", f"```{str(e)[:500]}```"))
    
    # =========================================================================
    # FINANCIAL CALCULATIONS
    # =========================================================================
    
    @commands.command(name="bep", aliases=["breakeven"])
    @commands.cooldown(1, 3, commands.BucketType.user)
    async def calculate_bep(self, ctx, fixed_cost: str, price: str, variable_cost: str):
        """Hitung Break Even Point"""
        try:
            fc = Formatters.parse_rupiah(fixed_cost)
            p = Formatters.parse_rupiah(price)
            vc = Formatters.parse_rupiah(variable_cost)
            
            contribution_margin = p - vc
            
            if contribution_margin <= 0:
                raise ValueError("Harga jual harus lebih besar dari biaya variabel")
            
            bep_units = fc / contribution_margin
            bep_rupiah = bep_units * p
            
            embed = discord.Embed(
                title=f"{EMOJIS['calculator']} Break Even Point",
                color=COLORS["finance"]
            )
            
            embed.add_field(
                name="Input",
                value=f"Fixed Cost: {Formatters.format_rupiah(fc)}\n"
                      f"Harga Jual/unit: {Formatters.format_rupiah(p)}\n"
                      f"Biaya Variabel/unit: {Formatters.format_rupiah(vc)}\n"
                      f"Contribution Margin: {Formatters.format_rupiah(contribution_margin)}",
                inline=False
            )
            
            embed.add_field(
                name="📊 BEP",
                value=f"**Unit: {Formatters.format_number(bep_units, 0)} unit**\n"
                      f"**Rupiah: {Formatters.format_rupiah(bep_rupiah)}**",
                inline=False
            )
            
            await ctx.send(embed=embed)
            await self.bot.log_usage(ctx.author.id, "bep", "finance", True)
            
        except Exception as e:
            await ctx.send(embed=Helpers.create_error_embed("Error BEP", f"```{str(e)}```"))
    
    @commands.command(name="roi")
    @commands.cooldown(1, 3, commands.BucketType.user)
    async def calculate_roi(self, ctx, investment: str, return_value: str):
        """Hitung ROI"""
        try:
            inv = Formatters.parse_rupiah(investment)
            ret = Formatters.parse_rupiah(return_value)
            
            profit = ret - inv
            roi_pct = (profit / inv) * 100 if inv > 0 else 0
            
            embed = discord.Embed(
                title=f"{EMOJIS['chart']} Return on Investment",
                color=COLORS["finance"]
            )
            
            embed.add_field(name="Investasi", value=Formatters.format_rupiah(inv), inline=True)
            embed.add_field(name="Return", value=Formatters.format_rupiah(ret), inline=True)
            embed.add_field(name="Profit", value=Formatters.format_rupiah(profit), inline=True)
            
            roi_emoji = EMOJIS['success'] if roi_pct > 0 else EMOJIS['error']
            embed.add_field(name=f"{roi_emoji} ROI", value=f"**{roi_pct:.2f}%**", inline=False)
            
            await ctx.send(embed=embed)
            await self.bot.log_usage(ctx.author.id, "roi", "finance", True)
            
        except Exception as e:
            await ctx.send(embed=Helpers.create_error_embed("Error ROI", f"```{str(e)}```"))
    
    @commands.command(name="depresiasi", aliases=["depreciation"])
    @commands.cooldown(1, 5, commands.BucketType.user)
    async def calculate_depreciation(self, ctx, cost: str, salvage: str, life: int, method: str = "straight"):
        """Hitung depresiasi"""
        try:
            cost_val = Formatters.parse_rupiah(cost)
            salvage_val = Formatters.parse_rupiah(salvage)
            
            annual_depreciation = (cost_val - salvage_val) / life
            
            embed = discord.Embed(
                title=f"{EMOJIS['calculator']} Depresiasi - Straight Line",
                color=COLORS["finance"]
            )
            
            embed.add_field(
                name="Input",
                value=f"Harga Perolehan: {Formatters.format_rupiah(cost_val)}\n"
                      f"Nilai Residu: {Formatters.format_rupiah(salvage_val)}\n"
                      f"Umur Ekonomis: {life} tahun",
                inline=False
            )
            
            embed.add_field(
                name="Depresiasi/Tahun",
                value=f"**{Formatters.format_rupiah(annual_depreciation)}**",
                inline=False
            )
            
            # Schedule preview
            schedule = "```\nTahun | Depresiasi     | Nilai Buku\n"
            schedule += "------|----------------|---------------\n"
            accumulated = 0
            for year in range(1, min(life + 1, 6)):
                accumulated += annual_depreciation
                book_value = cost_val - accumulated
                schedule += f"  {year}   | {annual_depreciation:>13,.0f} | {book_value:>13,.0f}\n"
            if life > 5:
                schedule += "...   | ...            | ...\n"
            schedule += "```"
            
            embed.add_field(name="Jadwal", value=schedule, inline=False)
            
            await ctx.send(embed=embed)
            await self.bot.log_usage(ctx.author.id, "depresiasi", "finance", True)
            
        except Exception as e:
            await ctx.send(embed=Helpers.create_error_embed("Error Depresiasi", f"```{str(e)}```"))
    
    # =========================================================================
    # HELPER METHODS
    # =========================================================================
    
    def _parse_income_statement_details(self, details: str) -> dict:
        """Parse income statement from natural language"""
        data = {
            "company": "Perusahaan",
            "period": "2024",
            "pendapatan": 0,
            "hpp": 0,
            "beban_operasional": {},
            "laba_bersih": 0
        }
        
        # Extract company name
        if "PT " in details.upper():
            match = re.search(r'PT\.?\s*([A-Za-z\s]+)', details, re.IGNORECASE)
            if match:
                data["company"] = f"PT {match.group(1).strip()}"
        
        # Extract period
        period_match = re.search(r'periode\s+([A-Za-z\-\s]+\d{4})', details, re.IGNORECASE)
        if period_match:
            data["period"] = period_match.group(1).strip()
        
        # Parse numbers (simplified)
        def parse_amount(text, keyword):
            pattern = rf'{keyword}[:\s]+(\d+(?:[.,]\d+)?)\s*(?:jt|juta|M|miliar)?'
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                num = float(match.group(1).replace(',', '.'))
                # Check for jt/M suffix
                if 'jt' in text[match.end():match.end()+5].lower() or 'juta' in text[match.end():match.end()+10].lower():
                    num *= 1_000_000
                elif 'm' in text[match.end():match.end()+3].lower() or 'miliar' in text[match.end():match.end()+10].lower():
                    num *= 1_000_000_000
                return num
            return 0
        
        data["pendapatan"] = parse_amount(details, "penjualan") or parse_amount(details, "pendapatan")
        data["hpp"] = parse_amount(details, "hpp") or parse_amount(details, "harga pokok")
        data["beban_operasional"] = {
            "gaji": parse_amount(details, "gaji") or parse_amount(details, "beban gaji"),
            "sewa": parse_amount(details, "sewa") or parse_amount(details, "beban sewa"),
            "listrik": parse_amount(details, "listrik") or parse_amount(details, "beban listrik"),
            "marketing": parse_amount(details, "marketing") or parse_amount(details, "beban marketing"),
            "lainnya": parse_amount(details, "lainnya") or parse_amount(details, "beban lain"),
        }
        
        # Calculate laba bersih
        total_beban = sum(data["beban_operasional"].values())
        laba_kotor = data["pendapatan"] - data["hpp"]
        data["laba_bersih"] = laba_kotor - total_beban
        
        return data
    
    async def _create_balance_sheet_structure(self, data: dict) -> dict:
        return {
            "sheets": [{
                "name": "Neraca",
                "headers": [],
                "data": [
                    ["NERACA (BALANCE SHEET)", "", ""],
                    [data.get('company', ''), "", ""],
                    [f"Per: {data.get('date', '')}", "", ""],
                    ["", "", ""],
                    ["ASET", "", ""],
                    ["Kas dan Setara Kas", "", data.get('assets', {}).get('kas', 0)],
                    ["Piutang Usaha", "", data.get('assets', {}).get('piutang', 0)],
                    ["Persediaan", "", data.get('assets', {}).get('persediaan', 0)],
                    ["Aset Tetap", "", data.get('assets', {}).get('aset_tetap', 0)],
                    ["TOTAL ASET", "", "=SUM(C6:C9)"],
                    ["", "", ""],
                    ["LIABILITAS", "", ""],
                    ["Hutang Usaha", "", data.get('liabilities', {}).get('hutang_usaha', 0)],
                    ["Hutang Bank", "", data.get('liabilities', {}).get('hutang_bank', 0)],
                    ["TOTAL LIABILITAS", "", "=SUM(C13:C14)"],
                    ["", "", ""],
                    ["EKUITAS", "", ""],
                    ["Modal", "", data.get('equity', {}).get('modal', 0)],
                    ["Laba Ditahan", "", data.get('equity', {}).get('laba_ditahan', 0)],
                    ["TOTAL EKUITAS", "", "=SUM(C18:C19)"],
                    ["", "", ""],
                    ["TOTAL LIABILITAS & EKUITAS", "", "=C15+C20"],
                ],
                "column_widths": {"A": 35, "B": 5, "C": 20},
                "formatting": {"currency_columns": ["C"]}
            }]
        }
    
    async def _create_income_statement_structure(self, data: dict) -> dict:
        beban = data.get('beban_operasional', {})
        
        return {
            "sheets": [{
                "name": "Laba Rugi",
                "headers": [],
                "data": [
                    ["LAPORAN LABA RUGI", ""],
                    [data.get('company', ''), ""],
                    [f"Periode: {data.get('period', '')}", ""],
                    ["", ""],
                    ["Pendapatan / Penjualan", data.get('pendapatan', 0)],
                    ["Harga Pokok Penjualan", data.get('hpp', 0)],
                    ["LABA KOTOR", "=B5-B6"],
                    ["", ""],
                    ["Beban Operasional:", ""],
                    ["  Beban Gaji", beban.get('gaji', 0)],
                    ["  Beban Sewa", beban.get('sewa', 0)],
                    ["  Beban Listrik", beban.get('listrik', 0)],
                    ["  Beban Marketing", beban.get('marketing', 0)],
                    ["  Beban Lainnya", beban.get('lainnya', 0)],
                    ["Total Beban Operasional", "=SUM(B10:B14)"],
                    ["", ""],
                    ["LABA OPERASIONAL", "=B7-B15"],
                    ["", ""],
                    ["Pajak Penghasilan (22%)", "=B17*0.22"],
                    ["", ""],
                    ["LABA BERSIH", "=B17-B19"],
                ],
                "column_widths": {"A": 35, "B": 20},
                "formatting": {"currency_columns": ["B"]}
            }]
        }
    
    async def _create_cashflow_structure(self, data: dict) -> dict:
        return {
            "sheets": [{
                "name": "Arus Kas",
                "headers": [],
                "data": [
                    ["LAPORAN ARUS KAS", ""],
                    [data.get('company', ''), ""],
                    [f"Periode: {data.get('period', '')}", ""],
                    ["", ""],
                    ["ARUS KAS DARI AKTIVITAS OPERASI", ""],
                    ["  Penerimaan dari Pelanggan", 0],
                    ["  Pembayaran kepada Pemasok", 0],
                    ["  Pembayaran Beban Operasi", 0],
                    ["Kas Bersih Aktivitas Operasi", "=SUM(B6:B8)"],
                    ["", ""],
                    ["ARUS KAS DARI AKTIVITAS INVESTASI", ""],
                    ["  Pembelian Aset Tetap", 0],
                    ["Kas Bersih Aktivitas Investasi", "=B12"],
                    ["", ""],
                    ["ARUS KAS DARI AKTIVITAS PENDANAAN", ""],
                    ["  Penerimaan Pinjaman", 0],
                    ["  Pembayaran Dividen", 0],
                    ["Kas Bersih Aktivitas Pendanaan", "=SUM(B16:B17)"],
                    ["", ""],
                    ["KENAIKAN/PENURUNAN KAS", "=B9+B13+B18"],
                    ["Kas Awal", 0],
                    ["KAS AKHIR", "=B20+B21"],
                ],
                "column_widths": {"A": 40, "B": 20},
                "formatting": {"currency_columns": ["B"]}
            }]
        }

# =============================================================================
# SETUP
# =============================================================================

async def setup(bot):
    await bot.add_cog(FinanceCog(bot))
