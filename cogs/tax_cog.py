"""
Tax Cog - Indonesian tax calculations (PPh, PPN, e-Faktur)
"""

import logging
from typing import Optional
import json

import discord
from discord.ext import commands

from config import EMOJIS, COLORS, PTKP_2024, PPH21_TARIF, PPN_RATE
from utils.formatters import Formatters
from utils.validators import Validators
from utils.helpers import Helpers

logger = logging.getLogger("office_bot.tax")

# =============================================================================
# TAX COG
# =============================================================================

class TaxCog(commands.Cog):
    """
    Commands untuk perhitungan pajak Indonesia
    """
    
    def __init__(self, bot):
        self.bot = bot
        logger.info("Tax Cog loaded")
    
    # =========================================================================
    # PPH 21 COMMANDS
    # =========================================================================
    
    @commands.command(name="pph21")
    @commands.cooldown(1, 5, commands.BucketType.user)
    async def calculate_pph21(self, ctx, gaji_setahun: str, status_ptkp: str = "TK/0"):
        """
        Hitung PPh 21 (Pajak Penghasilan Karyawan)
        
        Usage: !pph21 120jt K/1
        """
        loading = await ctx.send(embed=Helpers.create_loading_embed("Menghitung PPh 21..."))
        
        try:
            # Parse input
            gaji_tahunan = Formatters.parse_rupiah(gaji_setahun)
            status_ptkp = status_ptkp.upper()
            
            # Validate
            if not Validators.validate_ptkp_status(status_ptkp):
                raise ValueError(f"Status PTKP tidak valid: {status_ptkp}")
            
            # Get PTKP
            ptkp = PTKP_2024.get(status_ptkp, 54_000_000)
            
            # Calculate PKP (Penghasilan Kena Pajak)
            # Gaji tahunan - PTKP
            pkp = max(0, gaji_tahunan - ptkp)
            
            # Calculate PPh 21 using progressive tax rates
            pph21 = self._calculate_progressive_tax(pkp)
            
            # Create detailed embed
            embed = discord.Embed(
                title=f"{EMOJIS['tax']} Perhitungan PPh 21",
                color=COLORS["tax"]
            )
            
            embed.add_field(
                name="💰 Input",
                value=f"Gaji Setahun: {Formatters.format_rupiah(gaji_tahunan)}\n"
                      f"Status PTKP: {status_ptkp}\n"
                      f"Nilai PTKP: {Formatters.format_rupiah(ptkp)}",
                inline=False
            )
            
            embed.add_field(
                name="📊 Perhitungan",
                value=f"Penghasilan Setahun: {Formatters.format_rupiah(gaji_tahunan)}\n"
                      f"PTKP: {Formatters.format_rupiah(ptkp)}\n"
                      f"**PKP (Penghasilan Kena Pajak): {Formatters.format_rupiah(pkp)}**",
                inline=False
            )
            
            # Progressive tax breakdown
            breakdown = self._get_tax_breakdown(pkp)
            embed.add_field(
                name="🧮 Rincian Tarif Progresif",
                value=breakdown,
                inline=False
            )
            
            embed.add_field(
                name="💳 PPh 21",
                value=f"**Setahun: {Formatters.format_rupiah(pph21)}**\n"
                      f"Per Bulan: {Formatters.format_rupiah(pph21/12)}\n"
                      f"Effective Rate: {(pph21/gaji_tahunan*100) if gaji_tahunan > 0 else 0:.2f}%",
                inline=False
            )
            
            embed.set_footer(text="Perhitungan berdasarkan UU HPP 2022 (tarif progresif baru)")
            
            await loading.delete()
            await ctx.send(embed=embed)
            
            await self.bot.log_usage(ctx.author.id, "pph21", "tax", True)
            
        except Exception as e:
            logger.error(f"Error calculating PPh 21: {e}")
            await loading.delete()
            embed = Helpers.create_error_embed(
                title="Error Menghitung PPh 21",
                description=f"```{str(e)[:500]}```\n\n"
                           f"**Contoh penggunaan:**\n"
                           f"`!pph21 120000000 K/1`\n"
                           f"`!pph21 120jt TK/0`"
            )
            await ctx.send(embed=embed)
    
    @commands.command(name="pph23")
    @commands.cooldown(1, 3, commands.BucketType.user)
    async def calculate_pph23(self, ctx, jumlah: str, jenis: str = "jasa"):
        """
        Hitung PPh 23 (Jasa, Royalti, Dividen, Bunga)
        
        Usage: !pph23 10jt jasa
        Jenis: jasa (2%), dividen (15%), bunga (15%), royalti (15%)
        """
        try:
            nilai = Formatters.parse_rupiah(jumlah)
            jenis = jenis.lower()
            
            # PPh 23 rates
            rates = {
                "jasa": 0.02,      # 2%
                "dividen": 0.15,   # 15% (atau 10% jika ada DTA)
                "bunga": 0.15,     # 15%
                "royalti": 0.15,   # 15%
            }
            
            if jenis not in rates:
                raise ValueError(f"Jenis tidak valid. Gunakan: {', '.join(rates.keys())}")
            
            rate = rates[jenis]
            pph23 = nilai * rate
            
            embed = discord.Embed(
                title=f"{EMOJIS['tax']} PPh 23 - {jenis.upper()}",
                color=COLORS["tax"]
            )
            
            embed.add_field(
                name="Detail",
                value=f"Jumlah Bruto: {Formatters.format_rupiah(nilai)}\n"
                      f"Tarif: {rate*100}%\n"
                      f"**PPh 23: {Formatters.format_rupiah(pph23)}**\n"
                      f"Netto: {Formatters.format_rupiah(nilai - pph23)}",
                inline=False
            )
            
            embed.set_footer(text="PPh 23 dipotong oleh pemberi penghasilan")
            
            await ctx.send(embed=embed)
            await self.bot.log_usage(ctx.author.id, "pph23", "tax", True)
            
        except Exception as e:
            logger.error(f"Error calculating PPh 23: {e}")
            embed = Helpers.create_error_embed(
                title="Error Menghitung PPh 23",
                description=f"```{str(e)[:500]}```"
            )
            await ctx.send(embed=embed)
    
    # =========================================================================
    # PPN COMMANDS
    # =========================================================================
    
    @commands.command(name="ppn")
    @commands.cooldown(1, 3, commands.BucketType.user)
    async def calculate_ppn(self, ctx, dpp: str, include_ppn: bool = False):
        """
        Hitung PPN 11%
        
        Usage: !ppn 10jt
               !ppn 11.1jt true  (jika harga sudah include PPN)
        """
        try:
            nilai = Formatters.parse_rupiah(dpp)
            
            if include_ppn:
                # Harga sudah termasuk PPN, hitung DPP
                dpp_value = nilai / (1 + PPN_RATE)
                ppn_value = nilai - dpp_value
            else:
                # Harga belum termasuk PPN
                dpp_value = nilai
                ppn_value = nilai * PPN_RATE
            
            total = dpp_value + ppn_value
            
            embed = discord.Embed(
                title=f"{EMOJIS['tax']} PPN {int(PPN_RATE*100)}%",
                color=COLORS["tax"]
            )
            
            embed.add_field(
                name="📋 Rincian",
                value=f"DPP (Dasar Pengenaan Pajak): {Formatters.format_rupiah(dpp_value)}\n"
                      f"PPN {int(PPN_RATE*100)}%: {Formatters.format_rupiah(ppn_value)}\n"
                      f"**Total: {Formatters.format_rupiah(total)}**",
                inline=False
            )
            
            if include_ppn:
                embed.set_footer(text="Perhitungan dari harga yang sudah include PPN")
            else:
                embed.set_footer(text="PPN 11% berlaku sejak 1 April 2022")
            
            await ctx.send(embed=embed)
            await self.bot.log_usage(ctx.author.id, "ppn", "tax", True)
            
        except Exception as e:
            logger.error(f"Error calculating PPN: {e}")
            embed = Helpers.create_error_embed(
                title="Error Menghitung PPN",
                description=f"```{str(e)[:500]}```"
            )
            await ctx.send(embed=embed)
    
    @commands.command(name="ppnbm")
    @commands.cooldown(1, 3, commands.BucketType.user)
    async def calculate_ppnbm(self, ctx, dpp: str, tarif_ppnbm: float):
        """
        Hitung PPN + PPnBM (Pajak Penjualan Barang Mewah)
        
        Usage: !ppnbm 100jt 20
               (DPP, tarif PPnBM dalam %)
        """
        try:
            dpp_value = Formatters.parse_rupiah(dpp)
            
            ppn_value = dpp_value * PPN_RATE
            ppnbm_value = dpp_value * (tarif_ppnbm / 100)
            total = dpp_value + ppn_value + ppnbm_value
            
            embed = discord.Embed(
                title=f"{EMOJIS['tax']} PPN + PPnBM",
                color=COLORS["tax"]
            )
            
            embed.add_field(
                name="📋 Rincian",
                value=f"DPP: {Formatters.format_rupiah(dpp_value)}\n"
                      f"PPN 11%: {Formatters.format_rupiah(ppn_value)}\n"
                      f"PPnBM {tarif_ppnbm}%: {Formatters.format_rupiah(ppnbm_value)}\n"
                      f"**Total: {Formatters.format_rupiah(total)}**",
                inline=False
            )
            
            embed.set_footer(text="PPnBM untuk barang mewah (mobil, properti, dll)")
            
            await ctx.send(embed=embed)
            await self.bot.log_usage(ctx.author.id, "ppnbm", "tax", True)
            
        except Exception as e:
            logger.error(f"Error calculating PPnBM: {e}")
            embed = Helpers.create_error_embed(
                title="Error Menghitung PPnBM",
                description=f"```{str(e)[:500]}```"
            )
            await ctx.send(embed=embed)
    
    # =========================================================================
    # CORPORATE TAX
    # =========================================================================
    
    @commands.command(name="pph25", aliases=["pph_badan"])
    @commands.cooldown(1, 5, commands.BucketType.user)
    async def calculate_pph_badan(self, ctx, laba_kena_pajak: str, omzet_setahun: Optional[str] = None):
        """
        Hitung PPh Badan (Pasal 25/29)
        
        Usage: !pph25 500jt
               !pph25 500jt 10M  (dengan omzet untuk cek fasilitas)
        """
        try:
            laba = Formatters.parse_rupiah(laba_kena_pajak)
            
            # Default rate 22%
            rate = 0.22
            fasilitas = None
            
            if omzet_setahun:
                omzet = Formatters.parse_rupiah(omzet_setahun)
                
                # Check for UMKM facility (omzet < 4.8M: 0.5% final)
                if omzet <= 4_800_000_000:
                    fasilitas = "UMKM PP 23"
                    rate = 0.005  # 0.5% final
                # Check for SME facility (omzet < 50M: reduction)
                elif omzet <= 50_000_000_000:
                    fasilitas = "UU HPP - Pengurangan 50% untuk omzet < 50M (apply ke PKP s.d 4.8M)"
            
            pph_badan = laba * rate
            
            embed = discord.Embed(
                title=f"{EMOJIS['tax']} PPh Badan",
                color=COLORS["tax"]
            )
            
            if omzet_setahun:
                embed.add_field(
                    name="Input",
                    value=f"Laba Kena Pajak: {Formatters.format_rupiah(laba)}\n"
                          f"Omzet Setahun: {Formatters.format_rupiah(omzet)}",
                    inline=False
                )
            else:
                embed.add_field(
                    name="Input",
                    value=f"Laba Kena Pajak: {Formatters.format_rupiah(laba)}",
                    inline=False
                )
            
            embed.add_field(
                name="Tarif",
                value=f"{rate*100}%" + (f" ({fasilitas})" if fasilitas else " (Tarif Umum)"),
                inline=False
            )
            
            embed.add_field(
                name="PPh Badan",
                value=f"**Setahun: {Formatters.format_rupiah(pph_badan)}**\n"
                      f"Per Bulan (PPh 25): {Formatters.format_rupiah(pph_badan/12)}",
                inline=False
            )
            
            if fasilitas == "UMKM PP 23":
                embed.add_field(
                    name="ℹ️ Catatan",
                    value="UMKM dengan omzet < Rp 4.8M dapat PPh Final 0.5% dari omzet",
                    inline=False
                )
            
            embed.set_footer(text="Tarif PPh Badan 22% berlaku sejak 2020 (UU Cipta Kerja)")
            
            await ctx.send(embed=embed)
            await self.bot.log_usage(ctx.author.id, "pph25", "tax", True)
            
        except Exception as e:
            logger.error(f"Error calculating PPh Badan: {e}")
            embed = Helpers.create_error_embed(
                title="Error Menghitung PPh Badan",
                description=f"```{str(e)[:500]}```"
            )
            await ctx.send(embed=embed)
    
    # =========================================================================
    # INFO COMMANDS
    # =========================================================================
    
    @commands.command(name="infopajak", aliases=["taxptkp"])
    async def tax_info(self, ctx):
        """
        Informasi PTKP dan tarif pajak terbaru
        """
        embed = discord.Embed(
            title=f"{EMOJIS['tax']} Info Pajak Indonesia 2024",
            color=COLORS["tax"]
        )
        
        # PTKP
        ptkp_text = "```\n"
        ptkp_text += "Status  | PTKP/Tahun\n"
        ptkp_text += "--------|------------------\n"
        for status, amount in list(PTKP_2024.items())[:8]:
            ptkp_text += f"{status:7} | Rp {amount:>13,}\n"
        ptkp_text += "```"
        
        embed.add_field(
            name="📋 PTKP 2024",
            value=ptkp_text,
            inline=False
        )
        
        # Progressive tax rates
        tarif_text = "```\n"
        tarif_text += "PKP (Penghasilan Kena Pajak) | Tarif\n"
        tarif_text += "-----------------------------|------\n"
        tarif_text += "s.d Rp 60.000.000            |  5%\n"
        tarif_text += "Rp 60jt - Rp 250jt           | 15%\n"
        tarif_text += "Rp 250jt - Rp 500jt          | 25%\n"
        tarif_text += "Rp 500jt - Rp 5M             | 30%\n"
        tarif_text += "Di atas Rp 5M                | 35%\n"
        tarif_text += "```"
        
        embed.add_field(
            name="📊 Tarif PPh 21 Progresif",
            value=tarif_text,
            inline=False
        )
        
        # Other taxes
        embed.add_field(
            name="💼 Pajak Lainnya",
            value=f"• **PPN**: {int(PPN_RATE*100)}% (sejak 1 April 2022)\n"
                  f"• **PPh Badan**: 22%\n"
                  f"• **PPh 23 Jasa**: 2%\n"
                  f"• **PPh 23 Dividen/Bunga/Royalti**: 15%\n"
                  f"• **PPh Final UMKM**: 0.5% (omzet < Rp 4.8M)",
            inline=False
        )
        
        embed.set_footer(text="Berdasarkan UU HPP (Harmonisasi Peraturan Perpajakan) 2022")
        
        await ctx.send(embed=embed)
    
    # =========================================================================
    # HELPER METHODS
    # =========================================================================
    
    def _calculate_progressive_tax(self, pkp: float) -> float:
        """Calculate tax using progressive rates"""
        if pkp <= 0:
            return 0
        
        tax = 0
        remaining = pkp
        previous_bracket = 0
        
        for bracket, rate in PPH21_TARIF:
            taxable_in_bracket = min(remaining, bracket - previous_bracket)
            
            if taxable_in_bracket > 0:
                tax += taxable_in_bracket * rate
                remaining -= taxable_in_bracket
            
            if remaining <= 0:
                break
            
            previous_bracket = bracket
        
        return tax
    
    def _get_tax_breakdown(self, pkp: float) -> str:
        """Get progressive tax breakdown text"""
        if pkp <= 0:
            return "PKP = 0, tidak ada pajak"
        
        breakdown = "```\n"
        breakdown += "Lapisan          | PKP          | Tarif | Pajak\n"
        breakdown += "-----------------|--------------|-------|-------------\n"
        
        remaining = pkp
        previous_bracket = 0
        
        for i, (bracket, rate) in enumerate(PPH21_TARIF):
            taxable_in_bracket = min(remaining, bracket - previous_bracket)
            
            if taxable_in_bracket > 0:
                tax_in_bracket = taxable_in_bracket * rate
                
                # Format bracket name
                if i == 0:
                    bracket_name = f"s.d {bracket/1_000_000:.0f}jt"
                elif bracket == float('inf'):
                    bracket_name = f"> {previous_bracket/1_000_000:.0f}jt"
                else:
                    bracket_name = f"{previous_bracket/1_000_000:.0f}-{bracket/1_000_000:.0f}jt"
                
                breakdown += f"{bracket_name:16} | {taxable_in_bracket:>12,.0f} | {rate*100:>4.0f}% | {tax_in_bracket:>11,.0f}\n"
                
                remaining -= taxable_in_bracket
            
            if remaining <= 0:
                break
            
            previous_bracket = bracket
        
        breakdown += "```"
        return breakdown

# =============================================================================
# SETUP
# =============================================================================

async def setup(bot):
    await bot.add_cog(TaxCog(bot))
