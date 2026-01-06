"""
Finance Cog - Financial analysis, accounting, budgeting
"""

import logging
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
# FINANCE COG
# =============================================================================

class FinanceCog(commands.Cog):
    """
    Commands untuk finance, accounting, dan analisis keuangan
    """
    
    def __init__(self, bot):
        self.bot = bot
        logger.info("Finance Cog loaded")
    
    # =========================================================================
    # FINANCIAL STATEMENTS
    # =========================================================================
    
    @commands.command(name="neraca", aliases=["balance_sheet"])
    @commands.cooldown(1, 10, commands.BucketType.user)
    async def create_balance_sheet(self, ctx, *, details: str):
        """
        Buat template Neraca (Balance Sheet)
        
        Usage: !neraca untuk PT ABC per 31 Des 2024
        """
        loading = await ctx.send(embed=Helpers.create_loading_embed("Membuat neraca..."))
        
        try:
            ai_response = await self.bot.ai_engine.finance_analysis(
                f"""Buatkan struktur Neraca (Balance Sheet) untuk:

{details}

Format JSON:
{{
    "company": "...",
    "date": "DD/MM/YYYY",
    "assets": {{
        "current": {{"kas": 0, "piutang": 0, "persediaan": 0}},
        "non_current": {{"aset_tetap": 0, "akumulasi_penyusutan": 0}}
    }},
    "liabilities": {{
        "current": {{"hutang_usaha": 0, "hutang_jangka_pendek": 0}},
        "non_current": {{"hutang_jangka_panjang": 0}}
    }},
    "equity": {{
        "modal": 0,
        "laba_ditahan": 0
    }}
}}"""
            )
            
            if not ai_response.success:
                raise Exception(ai_response.error)
            
            # Parse response
            data = json.loads(ai_response.content)
            
            # Create Excel structure
            structure = await self._create_balance_sheet_structure(data)
            file_path = await self.bot.excel_engine.create_excel(structure)
            
            # Create embed
            embed = Helpers.create_success_embed(
                title="Neraca (Balance Sheet)",
                description=f"**{data.get('company', 'N/A')}**\n"
                           f"Per: {data.get('date', 'N/A')}"
            )
            
            # Calculate totals
            total_assets = self._sum_nested_dict(data.get('assets', {}))
            total_liabilities = self._sum_nested_dict(data.get('liabilities', {}))
            total_equity = self._sum_nested_dict(data.get('equity', {}))
            
            embed.add_field(
                name=f"{EMOJIS['money']} Summary",
                value=f"Total Aset: {Formatters.format_rupiah(total_assets)}\n"
                      f"Total Liabilitas: {Formatters.format_rupiah(total_liabilities)}\n"
                      f"Total Ekuitas: {Formatters.format_rupiah(total_equity)}",
                inline=False
            )
            
            if abs(total_assets - (total_liabilities + total_equity)) > 0.01:
                embed.add_field(
                    name=f"{EMOJIS['warning']} Perhatian",
                    value="Aset ≠ Liabilitas + Ekuitas (belum balance)",
                    inline=False
                )
            
            # Send file
            discord_file = self.bot.file_handler.create_discord_file(file_path)
            await loading.delete()
            await ctx.send(embed=embed, file=discord_file)
            
            await self.bot.log_usage(ctx.author.id, "neraca", "finance", True)
            await self.bot.file_handler.delete_file(file_path)
            
        except Exception as e:
            logger.error(f"Error creating balance sheet: {e}")
            await loading.delete()
            embed = Helpers.create_error_embed(
                title="Error Membuat Neraca",
                description=f"```{str(e)[:500]}```"
            )
            await ctx.send(embed=embed)
    
    @commands.command(name="labarugi", aliases=["income_statement", "pnl"])
    @commands.cooldown(1, 10, commands.BucketType.user)
    async def create_income_statement(self, ctx, *, details: str):
        """
        Buat Laporan Laba Rugi (Income Statement)
        
        Usage: !labarugi PT ABC periode Jan-Des 2024
        """
        loading = await ctx.send(embed=Helpers.create_loading_embed("Membuat laporan laba rugi..."))
        
        try:
            ai_response = await self.bot.ai_engine.finance_analysis(
                f"""Buatkan struktur Laporan Laba Rugi untuk:

{details}

Format JSON dengan komponen:
- Pendapatan
- Harga Pokok Penjualan (HPP)
- Laba Kotor
- Beban Operasional
- Laba Operasional
- Pendapatan/Beban Lain-lain
- Laba Sebelum Pajak
- Pajak Penghasilan
- Laba Bersih"""
            )
            
            if not ai_response.success:
                raise Exception(ai_response.error)
            
            data = json.loads(ai_response.content)
            
            # Create structure
            structure = await self._create_income_statement_structure(data)
            file_path = await self.bot.excel_engine.create_excel(structure)
            
            # Calculate key metrics
            pendapatan = data.get('pendapatan', 0)
            laba_bersih = data.get('laba_bersih', 0)
            margin = (laba_bersih / pendapatan * 100) if pendapatan > 0 else 0
            
            embed = Helpers.create_success_embed(
                title="Laporan Laba Rugi",
                description=f"**{data.get('company', 'N/A')}**\n"
                           f"Periode: {data.get('period', 'N/A')}"
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
            embed = Helpers.create_error_embed(
                title="Error Membuat Laba Rugi",
                description=f"```{str(e)[:500]}```"
            )
            await ctx.send(embed=embed)
    
    @commands.command(name="cashflow", aliases=["arus_kas"])
    @commands.cooldown(1, 10, commands.BucketType.user)
    async def create_cashflow(self, ctx, *, details: str):
        """
        Buat Laporan Arus Kas (Cash Flow Statement)
        
        Usage: !cashflow PT ABC Q1 2024
        """
        loading = await ctx.send(embed=Helpers.create_loading_embed("Membuat laporan arus kas..."))
        
        try:
            ai_response = await self.bot.ai_engine.finance_analysis(
                f"""Buatkan struktur Laporan Arus Kas untuk:

{details}

Format dengan 3 aktivitas:
1. Aktivitas Operasi
2. Aktivitas Investasi
3. Aktivitas Pendanaan

Output JSON"""
            )
            
            if not ai_response.success:
                raise Exception(ai_response.error)
            
            data = json.loads(ai_response.content)
            structure = await self._create_cashflow_structure(data)
            file_path = await self.bot.excel_engine.create_excel(structure)
            
            embed = Helpers.create_success_embed(
                title="Laporan Arus Kas",
                description=f"**{data.get('company', 'N/A')}**\n"
                           f"Periode: {data.get('period', 'N/A')}"
            )
            
            discord_file = self.bot.file_handler.create_discord_file(file_path)
            await loading.delete()
            await ctx.send(embed=embed, file=discord_file)
            
            await self.bot.log_usage(ctx.author.id, "cashflow", "finance", True)
            await self.bot.file_handler.delete_file(file_path)
            
        except Exception as e:
            logger.error(f"Error creating cashflow: {e}")
            await loading.delete()
            embed = Helpers.create_error_embed(
                title="Error Membuat Arus Kas",
                description=f"```{str(e)[:500]}```"
            )
            await ctx.send(embed=embed)
    
    # =========================================================================
    # FINANCIAL CALCULATIONS
    # =========================================================================
    
    @commands.command(name="bep", aliases=["breakeven"])
    @commands.cooldown(1, 3, commands.BucketType.user)
    async def calculate_bep(self, ctx, fixed_cost: str, price: str, variable_cost: str):
        """
        Hitung Break Even Point
        
        Usage: !bep 100jt 50000 30000
               (fixed cost, harga jual, biaya variabel per unit)
        """
        try:
            fc = Formatters.parse_rupiah(fixed_cost)
            p = Formatters.parse_rupiah(price)
            vc = Formatters.parse_rupiah(variable_cost)
            
            # BEP (unit) = Fixed Cost / (Price - Variable Cost)
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
                value=f"**Unit: {Formatters.format_number(bep_units, 2)} unit**\n"
                      f"**Rupiah: {Formatters.format_rupiah(bep_rupiah)}**",
                inline=False
            )
            
            embed.add_field(
                name="💡 Insight",
                value=f"Anda perlu menjual {int(bep_units)+1} unit untuk mulai profit.\n"
                      f"Margin per unit: {Formatters.format_rupiah(contribution_margin)}",
                inline=False
            )
            
            await ctx.send(embed=embed)
            await self.bot.log_usage(ctx.author.id, "bep", "finance", True)
            
        except Exception as e:
            logger.error(f"Error calculating BEP: {e}")
            embed = Helpers.create_error_embed(
                title="Error Menghitung BEP",
                description=f"```{str(e)[:500]}```"
            )
            await ctx.send(embed=embed)
    
    @commands.command(name="roi")
    @commands.cooldown(1, 3, commands.BucketType.user)
    async def calculate_roi(self, ctx, investment: str, return_value: str):
        """
        Hitung Return on Investment (ROI)
        
        Usage: !roi 100jt 150jt
        """
        try:
            inv = Formatters.parse_rupiah(investment)
            ret = Formatters.parse_rupiah(return_value)
            
            profit = ret - inv
            roi_pct = (profit / inv) * 100
            
            embed = discord.Embed(
                title=f"{EMOJIS['chart']} Return on Investment",
                color=COLORS["finance"]
            )
            
            embed.add_field(
                name="Investasi",
                value=Formatters.format_rupiah(inv),
                inline=True
            )
            
            embed.add_field(
                name="Return",
                value=Formatters.format_rupiah(ret),
                inline=True
            )
            
            embed.add_field(
                name="Profit",
                value=Formatters.format_rupiah(profit),
                inline=True
            )
            
            roi_emoji = EMOJIS['success'] if roi_pct > 0 else EMOJIS['error']
            embed.add_field(
                name=f"{roi_emoji} ROI",
                value=f"**{roi_pct:.2f}%**",
                inline=False
            )
            
            await ctx.send(embed=embed)
            await self.bot.log_usage(ctx.author.id, "roi", "finance", True)
            
        except Exception as e:
            logger.error(f"Error calculating ROI: {e}")
            embed = Helpers.create_error_embed(
                title="Error Menghitung ROI",
                description=f"```{str(e)[:500]}```"
            )
            await ctx.send(embed=embed)
    
    @commands.command(name="depresiasi", aliases=["depreciation"])
    @commands.cooldown(1, 5, commands.BucketType.user)
    async def calculate_depreciation(self, ctx, cost: str, salvage: str, life: int, method: str = "straight"):
        """
        Hitung depresiasi aset
        
        Usage: !depresiasi 100jt 10jt 5 straight
        Methods: straight, declining
        """
        try:
            cost_val = Formatters.parse_rupiah(cost)
            salvage_val = Formatters.parse_rupiah(salvage)
            
            if method.lower() == "straight":
                # Straight Line: (Cost - Salvage) / Life
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
                
                # Schedule
                schedule = "```\n"
                schedule += "Tahun | Depresiasi      | Akumulasi       | Nilai Buku\n"
                schedule += "------|-----------------|-----------------|----------------\n"
                
                accumulated = 0
                for year in range(1, min(life + 1, 6)):  # Max 5 years preview
                    accumulated += annual_depreciation
                    book_value = cost_val - accumulated
                    schedule += f"  {year}   | {annual_depreciation:>13,.0f} | {accumulated:>13,.0f} | {book_value:>13,.0f}\n"
                
                if life > 5:
                    schedule += f"...   | ...             | ...             | ...\n"
                
                schedule += "```"
                
                embed.add_field(
                    name="Jadwal Depresiasi",
                    value=schedule,
                    inline=False
                )
                
            else:
                embed = Helpers.create_error_embed(
                    title="Method Tidak Didukung",
                    description="Gunakan: straight atau declining"
                )
            
            await ctx.send(embed=embed)
            await self.bot.log_usage(ctx.author.id, "depresiasi", "finance", True)
            
        except Exception as e:
            logger.error(f"Error calculating depreciation: {e}")
            embed = Helpers.create_error_embed(
                title="Error Menghitung Depresiasi",
                description=f"```{str(e)[:500]}```"
            )
            await ctx.send(embed=embed)
    
    @commands.command(name="rasio", aliases=["ratio"])
    @commands.cooldown(1, 5, commands.BucketType.user)
    async def financial_ratios(self, ctx):
        """
        Hitung rasio keuangan dari file Excel
        
        Usage: Upload balance sheet + income statement + !rasio
        """
        if not ctx.message.attachments:
            embed = Helpers.create_error_embed(
                title="File Tidak Ditemukan",
                description="Upload file Excel dengan data keuangan!"
            )
            await ctx.send(embed=embed)
            return
        
        loading = await ctx.send(embed=Helpers.create_loading_embed("Menghitung rasio keuangan..."))
        
        try:
            attachment = ctx.message.attachments[0]
            file_path = await self.bot.file_handler.download_attachment(attachment)
            
            # Read Excel
            excel_data = await self.bot.excel_engine.read_excel(file_path)
            
            # Use AI to analyze
            ai_response = await self.bot.ai_engine.finance_analysis(
                f"""Dari data Excel berikut, hitung rasio keuangan:

{str(excel_data['data'][:100])}

Hitung:
1. Current Ratio = Aset Lancar / Liabilitas Lancar
2. Quick Ratio = (Aset Lancar - Persediaan) / Liabilitas Lancar
3. Debt to Equity Ratio
4. Return on Assets (ROA)
5. Return on Equity (ROE)
6. Profit Margin

Berikan hasil dan interpretasi."""
            )
            
            if not ai_response.success:
                raise Exception(ai_response.error)
            
            embed = Helpers.create_success_embed(
                title="Analisis Rasio Keuangan",
                description=Helpers.truncate(ai_response.content, 4000)
            )
            
            await loading.delete()
            await ctx.send(embed=embed)
            
            await self.bot.log_usage(ctx.author.id, "rasio", "finance", True)
            await self.bot.file_handler.delete_file(file_path)
            
        except Exception as e:
            logger.error(f"Error calculating ratios: {e}")
            await loading.delete()
            embed = Helpers.create_error_embed(
                title="Error Menghitung Rasio",
                description=f"```{str(e)[:500]}```"
            )
            await ctx.send(embed=embed)
    
    # =========================================================================
    # HELPER METHODS
    # =========================================================================
    
    def _sum_nested_dict(self, d: dict) -> float:
        """Sum all numeric values in nested dict"""
        total = 0
        for value in d.values():
            if isinstance(value, dict):
                total += self._sum_nested_dict(value)
            elif isinstance(value, (int, float)):
                total += value
        return total
    
    async def _create_balance_sheet_structure(self, data: dict) -> dict:
        """Create balance sheet Excel structure"""
        return {
            "sheets": [{
                "name": "Neraca",
                "data": [
                    ["NERACA (BALANCE SHEET)", "", ""],
                    [data.get('company', ''), "", ""],
                    [f"Per: {data.get('date', '')}", "", ""],
                    ["", "", ""],
                    ["ASET", "", ""],
                    ["Aset Lancar:", "", ""],
                    ["  Kas dan Setara Kas", "", 0],
                    ["  Piutang Usaha", "", 0],
                    ["  Persediaan", "", 0],
                    ["Total Aset Lancar", "", "=SUM(C7:C9)"],
                    ["", "", ""],
                    ["Aset Tidak Lancar:", "", ""],
                    ["  Aset Tetap", "", 0],
                    ["  Akumulasi Penyusutan", "", 0],
                    ["Total Aset Tidak Lancar", "", "=C13+C14"],
                    ["", "", ""],
                    ["TOTAL ASET", "", "=C10+C15"],
                    ["", "", ""],
                    ["LIABILITAS", "", ""],
                    ["Liabilitas Jangka Pendek:", "", ""],
                    ["  Hutang Usaha", "", 0],
                    ["Total Liabilitas Jangka Pendek", "", "=C21"],
                    ["", "", ""],
                    ["EKUITAS", "", ""],
                    ["  Modal", "", 0],
                    ["  Laba Ditahan", "", 0],
                    ["Total Ekuitas", "", "=C25+C26"],
                    ["", "", ""],
                    ["TOTAL LIABILITAS & EKUITAS", "", "=C22+C27"],
                ],
                "column_widths": {"A": 35, "B": 5, "C": 20},
                "formatting": {
                    "currency_columns": ["C"]
                }
            }]
        }
    
    async def _create_income_statement_structure(self, data: dict) -> dict:
        """Create income statement structure"""
        return {
            "sheets": [{
                "name": "Laba Rugi",
                "data": [
                    ["LAPORAN LABA RUGI", ""],
                    [data.get('company', ''), ""],
                    [f"Periode: {data.get('period', '')}", ""],
                    ["", ""],
                    ["Pendapatan", 0],
                    ["Harga Pokok Penjualan", 0],
                    ["Laba Kotor", "=B5-B6"],
                    ["", ""],
                    ["Beban Operasional:", ""],
                    ["  Beban Gaji", 0],
                    ["  Beban Sewa", 0],
                    ["  Beban Lain-lain", 0],
                    ["Total Beban Operasional", "=SUM(B10:B12)"],
                    ["", ""],
                    ["Laba Operasional", "=B7-B13"],
                    ["", ""],
                    ["Pendapatan/Beban Lain-lain", 0],
                    ["", ""],
                    ["Laba Sebelum Pajak", "=B15+B17"],
                    ["Pajak Penghasilan", "=B19*0.22"],
                    ["", ""],
                    ["LABA BERSIH", "=B19-B20"],
                ],
                "column_widths": {"A": 35, "B": 20},
                "formatting": {
                    "currency_columns": ["B"]
                }
            }]
        }
    
    async def _create_cashflow_structure(self, data: dict) -> dict:
        """Create cash flow statement structure"""
        return {
            "sheets": [{
                "name": "Arus Kas",
                "data": [
                    ["LAPORAN ARUS KAS", ""],
                    [data.get('company', ''), ""],
                    [f"Periode: {data.get('period', '')}", ""],
                    ["", ""],
                    ["ARUS KAS DARI AKTIVITAS OPERASI", ""],
                    ["  Penerimaan dari Pelanggan", 0],
                    ["  Pembayaran kepada Pemasok", 0],
                    ["  Pembayaran Beban Operasi", 0],
                    ["Kas Bersih dari Aktivitas Operasi", "=B6+B7+B8"],
                    ["", ""],
                    ["ARUS KAS DARI AKTIVITAS INVESTASI", ""],
                    ["  Pembelian Aset Tetap", 0],
                    ["Kas Bersih dari Aktivitas Investasi", "=B12"],
                    ["", ""],
                    ["ARUS KAS DARI AKTIVITAS PENDANAAN", ""],
                    ["  Penerimaan Pinjaman", 0],
                    ["  Pembayaran Dividen", 0],
                    ["Kas Bersih dari Aktivitas Pendanaan", "=B16+B17"],
                    ["", ""],
                    ["KENAIKAN/PENURUNAN KAS BERSIH", "=B9+B13+B18"],
                    ["Kas Awal Periode", 0],
                    ["KAS AKHIR PERIODE", "=B20+B21"],
                ],
                "column_widths": {"A": 40, "B": 20},
                "formatting": {
                    "currency_columns": ["B"]
                }
            }]
        }

# =============================================================================
# SETUP
# =============================================================================

async def setup(bot):
    await bot.add_cog(FinanceCog(bot))
