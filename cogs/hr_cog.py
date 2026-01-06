"""
HR Cog - Payroll, attendance, employee management
"""

import logging
from typing import Optional
from datetime import datetime

import discord
from discord.ext import commands

from config import EMOJIS, COLORS, PTKP_2024
from utils.formatters import Formatters
from utils.validators import Validators
from utils.helpers import Helpers

logger = logging.getLogger("office_bot.hr")

# =============================================================================
# HR COG
# =============================================================================

class HRCog(commands.Cog):
    """
    Commands untuk HR, Payroll, dan employee management
    """
    
    def __init__(self, bot):
        self.bot = bot
        logger.info("HR Cog loaded")
    
    # =========================================================================
    # PAYROLL COMMANDS
    # =========================================================================
    
    @commands.command(name="gaji", aliases=["salary"])
    @commands.cooldown(1, 5, commands.BucketType.user)
    async def calculate_salary(self, ctx, gaji_pokok: str, *, status_ptkp: str = "TK/0"):
        """
        Hitung take home pay (gaji bersih)
        
        Usage: !gaji 10000000 K/1
        """
        loading = await ctx.send(embed=Helpers.create_loading_embed("Menghitung gaji..."))
        
        try:
            # Parse salary
            gaji_pokok_val = Formatters.parse_rupiah(gaji_pokok)
            
            if not Validators.validate_positive_number(gaji_pokok_val):
                raise ValueError("Gaji pokok harus angka positif")
            
            # Validate PTKP
            status_ptkp = status_ptkp.upper()
            if not Validators.validate_ptkp_status(status_ptkp):
                raise ValueError(f"Status PTKP tidak valid. Gunakan: TK/0, K/1, K/I/0, dll")
            
            # Calculate payroll
            result = await self.bot.ai_engine.hr_calculation(
                f"""Hitung take home pay (gaji bersih) dengan detail:

Gaji Pokok: {Formatters.format_rupiah(gaji_pokok_val)}
Status PTKP: {status_ptkp}

Hitung:
1. BPJS Kesehatan (1% dari gaji)
2. BPJS Ketenagakerjaan/JHT (2% dari gaji)
3. PPh 21 (gunakan PTKP {status_ptkp} = {Formatters.format_rupiah(PTKP_2024.get(status_ptkp, 54000000))})

Format output JSON:
{{
    "gaji_pokok": {gaji_pokok_val},
    "bpjs_kesehatan": 0,
    "bpjs_tk": 0,
    "pph21": 0,
    "total_potongan": 0,
    "gaji_bersih": 0
}}
"""
            )
            
            if not result.success:
                raise Exception(result.error)
            
            import json
            data = json.loads(result.content)
            
            # Create embed
            embed = discord.Embed(
                title=f"{EMOJIS['money']} Perhitungan Gaji",
                color=COLORS["finance"]
            )
            
            embed.add_field(
                name="💰 Pendapatan",
                value=f"Gaji Pokok: **{Formatters.format_rupiah(data['gaji_pokok'])}**",
                inline=False
            )
            
            embed.add_field(
                name="➖ Potongan",
                value=f"BPJS Kesehatan (1%): {Formatters.format_rupiah(data['bpjs_kesehatan'])}\n"
                      f"BPJS TK/JHT (2%): {Formatters.format_rupiah(data['bpjs_tk'])}\n"
                      f"PPh 21 ({status_ptkp}): {Formatters.format_rupiah(data['pph21'])}\n"
                      f"**Total Potongan:** {Formatters.format_rupiah(data['total_potongan'])}",
                inline=False
            )
            
            embed.add_field(
                name="✅ Gaji Bersih (Take Home Pay)",
                value=f"**{Formatters.format_rupiah(data['gaji_bersih'])}**\n"
                      f"_{Formatters.terbilang(data['gaji_bersih'])}_",
                inline=False
            )
            
            await loading.delete()
            await ctx.send(embed=embed)
            
            await self.bot.log_usage(ctx.author.id, "gaji", "hr", True)
            
        except Exception as e:
            logger.error(f"Error calculating salary: {e}")
            await loading.delete()
            embed = Helpers.create_error_embed(
                title="Error Menghitung Gaji",
                description=f"```{str(e)[:500]}```"
            )
            await ctx.send(embed=embed)
    
    @commands.command(name="slipgaji", aliases=["payslip"])
    @commands.cooldown(1, 5, commands.BucketType.user)
    async def generate_payslip(self, ctx, *, details: str):
        """
        Generate slip gaji Excel
        
        Usage: !slipgaji nama: John, NIK: 123, gaji: 10jt, status: K/1
        """
        loading = await ctx.send(embed=Helpers.create_loading_embed("Membuat slip gaji..."))
        
        try:
            # Use template
            template_path = await self.bot.excel_engine.create_template("payroll")
            
            # Get employee details from AI
            ai_response = await self.bot.ai_engine.hr_calculation(
                f"Parse data karyawan dari: {details} dan buatkan struktur slip gaji"
            )
            
            embed = Helpers.create_success_embed(
                title="Slip Gaji Berhasil Dibuat",
                description="Template slip gaji telah digenerate. Silakan edit sesuai kebutuhan."
            )
            
            discord_file = self.bot.file_handler.create_discord_file(template_path)
            await loading.delete()
            await ctx.send(embed=embed, file=discord_file)
            
            await self.bot.log_usage(ctx.author.id, "slipgaji", "hr", True)
            await self.bot.file_handler.delete_file(template_path)
            
        except Exception as e:
            logger.error(f"Error generating payslip: {e}")
            await loading.delete()
            embed = Helpers.create_error_embed(
                title="Error Membuat Slip Gaji",
                description=f"```{str(e)[:500]}```"
            )
            await ctx.send(embed=embed)
    
    @commands.command(name="lembur", aliases=["overtime"])
    @commands.cooldown(1, 3, commands.BucketType.user)
    async def calculate_overtime(self, ctx, gaji_pokok: str, jam_lembur: float):
        """
        Hitung uang lembur
        
        Usage: !lembur 8000000 10
        """
        try:
            gaji_val = Formatters.parse_rupiah(gaji_pokok)
            
            # Rumus lembur Indonesia (simplified)
            # Upah per jam = 1/173 x gaji pokok
            # Jam 1: 1.5x upah/jam
            # Jam 2+: 2x upah/jam
            
            upah_per_jam = gaji_val / 173
            
            if jam_lembur <= 1:
                uang_lembur = jam_lembur * upah_per_jam * 1.5
            else:
                uang_lembur = (upah_per_jam * 1.5) + ((jam_lembur - 1) * upah_per_jam * 2)
            
            embed = discord.Embed(
                title=f"{EMOJIS['calculator']} Perhitungan Lembur",
                color=COLORS["finance"]
            )
            
            embed.add_field(
                name="Detail",
                value=f"Gaji Pokok: {Formatters.format_rupiah(gaji_val)}\n"
                      f"Jam Lembur: {jam_lembur} jam\n"
                      f"Upah/Jam: {Formatters.format_rupiah(upah_per_jam)}",
                inline=False
            )
            
            embed.add_field(
                name="Uang Lembur",
                value=f"**{Formatters.format_rupiah(uang_lembur)}**",
                inline=False
            )
            
            await ctx.send(embed=embed)
            await self.bot.log_usage(ctx.author.id, "lembur", "hr", True)
            
        except Exception as e:
            logger.error(f"Error calculating overtime: {e}")
            embed = Helpers.create_error_embed(
                title="Error Menghitung Lembur",
                description=f"```{str(e)[:500]}```"
            )
            await ctx.send(embed=embed)
    
    @commands.command(name="ptkp")
    async def show_ptkp(self, ctx):
        """
        Tampilkan tabel PTKP 2024
        """
        embed = discord.Embed(
            title="📋 PTKP (Penghasilan Tidak Kena Pajak) 2024",
            color=COLORS["tax"]
        )
        
        ptkp_text = "```\n"
        ptkp_text += "Status    | PTKP\n"
        ptkp_text += "----------|------------------\n"
        for status, amount in PTKP_2024.items():
            ptkp_text += f"{status:9} | Rp {amount:>13,}\n"
        ptkp_text += "```"
        
        embed.description = ptkp_text
        embed.set_footer(text="TK = Tidak Kawin | K = Kawin | I = Istri Bekerja | 0-3 = Tanggungan")
        
        await ctx.send(embed=embed)

# =============================================================================
# SETUP
# =============================================================================

async def setup(bot):
    await bot.add_cog(HRCog(bot))
