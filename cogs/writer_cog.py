"""
Writer Cog - Copywriting & document generation
"""

import logging
from typing import Optional

import discord
from discord.ext import commands

from config import EMOJIS, COLORS
from utils.helpers import Helpers

logger = logging.getLogger("office_bot.writer")

# =============================================================================
# WRITER COG
# =============================================================================

class WriterCog(commands.Cog):
    """
    Commands untuk copywriting dan pembuatan dokumen
    """
    
    def __init__(self, bot):
        self.bot = bot
        logger.info("Writer Cog loaded")
    
    # =========================================================================
    # DOCUMENT COMMANDS
    # =========================================================================
    
    @commands.command(name="email")
    @commands.cooldown(1, 5, commands.BucketType.user)
    async def write_email(self, ctx, *, topic: str):
        """
        Tulis email bisnis profesional
        
        Usage: !email meeting request untuk project review
        """
        loading = await ctx.send(embed=Helpers.create_loading_embed("Menulis email..."))
        
        try:
            ai_response = await self.bot.ai_engine.generate_document(
                f"""Buatkan email bisnis profesional dalam Bahasa Indonesia untuk:

{topic}

Format:
- Subject line
- Greeting formal
- Body email yang jelas dan profesional
- Closing yang tepat
- Signature placeholder

Gunakan bahasa yang sopan dan profesional."""
            )
            
            if not ai_response.success:
                raise Exception(ai_response.error)
            
            embed = discord.Embed(
                title=f"{EMOJIS['file']} Email Draft",
                description=Helpers.truncate(ai_response.content, 4000),
                color=COLORS["primary"]
            )
            
            embed.set_footer(text="Silakan edit sesuai kebutuhan")
            
            await loading.delete()
            await ctx.send(embed=embed)
            
            await self.bot.log_usage(ctx.author.id, "email", "writer", True)
            
        except Exception as e:
            logger.error(f"Error writing email: {e}")
            await loading.delete()
            embed = Helpers.create_error_embed(
                title="Error Menulis Email",
                description=f"```{str(e)[:500]}```"
            )
            await ctx.send(embed=embed)
    
    @commands.command(name="memo")
    @commands.cooldown(1, 5, commands.BucketType.user)
    async def write_memo(self, ctx, *, topic: str):
        """
        Buat memo internal
        
        Usage: !memo perubahan jam kerja mulai besok
        """
        loading = await ctx.send(embed=Helpers.create_loading_embed("Membuat memo..."))
        
        try:
            ai_response = await self.bot.ai_engine.generate_document(
                f"""Buatkan memo internal perusahaan tentang:

{topic}

Format memo standar Indonesia:
- Header: MEMO INTERNAL
- Kepada:
- Dari:
- Tanggal:
- Perihal:
- Isi memo yang jelas dan ringkas
"""
            )
            
            if not ai_response.success:
                raise Exception(ai_response.error)
            
            embed = discord.Embed(
                title=f"📝 Memo Internal",
                description=Helpers.truncate(ai_response.content, 4000),
                color=COLORS["info"]
            )
            
            await loading.delete()
            await ctx.send(embed=embed)
            
            await self.bot.log_usage(ctx.author.id, "memo", "writer", True)
            
        except Exception as e:
            logger.error(f"Error writing memo: {e}")
            await loading.delete()
            embed = Helpers.create_error_embed(
                title="Error Membuat Memo",
                description=f"```{str(e)[:500]}```"
            )
            await ctx.send(embed=embed)
    
    @commands.command(name="proposal")
    @commands.cooldown(1, 5, commands.BucketType.user)
    async def write_proposal(self, ctx, *, topic: str):
        """
        Buat draft proposal
        
        Usage: !proposal kerjasama partnership dengan vendor baru
        """
        loading = await ctx.send(embed=Helpers.create_loading_embed("Membuat proposal..."))
        
        try:
            ai_response = await self.bot.ai_engine.generate_document(
                f"""Buatkan outline proposal bisnis untuk:

{topic}

Sertakan:
1. Latar Belakang
2. Tujuan
3. Ruang Lingkup
4. Metodologi/Rencana Kerja
5. Timeline
6. Anggaran (outline)
7. Penutup

Gunakan bahasa formal dan profesional."""
            )
            
            if not ai_response.success:
                raise Exception(ai_response.error)
            
            # Split into pages if too long
            pages = Helpers.paginate_text(ai_response.content, 2000)
            
            for i, page in enumerate(pages, 1):
                embed = discord.Embed(
                    title=f"📄 Proposal Draft (Halaman {i}/{len(pages)})",
                    description=page,
                    color=COLORS["primary"]
                )
                await ctx.send(embed=embed)
            
            await loading.delete()
            await self.bot.log_usage(ctx.author.id, "proposal", "writer", True)
            
        except Exception as e:
            logger.error(f"Error writing proposal: {e}")
            await loading.delete()
            embed = Helpers.create_error_embed(
                title="Error Membuat Proposal",
                description=f"```{str(e)[:500]}```"
            )
            await ctx.send(embed=embed)
    
    @commands.command(name="notulen", aliases=["minutes"])
    @commands.cooldown(1, 5, commands.BucketType.user)
    async def write_minutes(self, ctx, *, details: str):
        """
        Buat notulensi rapat
        
        Usage: !notulen meeting review Q1, peserta: team leads, hasil: ...
        """
        loading = await ctx.send(embed=Helpers.create_loading_embed("Membuat notulensi..."))
        
        try:
            ai_response = await self.bot.ai_engine.generate_document(
                f"""Buatkan notulensi rapat dari informasi:

{details}

Format:
- Judul Rapat
- Tanggal & Waktu
- Tempat
- Peserta
- Agenda
- Pembahasan (poin-poin penting)
- Keputusan
- Action Items (siapa, apa, kapan)
- Rapat ditutup jam berapa
"""
            )
            
            if not ai_response.success:
                raise Exception(ai_response.error)
            
            embed = discord.Embed(
                title="📋 Notulensi Rapat",
                description=Helpers.truncate(ai_response.content, 4000),
                color=COLORS["info"]
            )
            
            await loading.delete()
            await ctx.send(embed=embed)
            
            await self.bot.log_usage(ctx.author.id, "notulen", "writer", True)
            
        except Exception as e:
            logger.error(f"Error writing minutes: {e}")
            await loading.delete()
            embed = Helpers.create_error_embed(
                title="Error Membuat Notulensi",
                description=f"```{str(e)[:500]}```"
            )
            await ctx.send(embed=embed)
    
    @commands.command(name="surat", aliases=["letter"])
    @commands.cooldown(1, 5, commands.BucketType.user)
    async def write_letter(self, ctx, *, purpose: str):
        """
        Buat surat resmi
        
        Usage: !surat permohonan kerjasama ke PT ABC
        """
        loading = await ctx.send(embed=Helpers.create_loading_embed("Membuat surat..."))
        
        try:
            ai_response = await self.bot.ai_engine.generate_document(
                f"""Buatkan surat resmi untuk:

{purpose}

Format surat resmi Indonesia:
- Kop surat (placeholder)
- Nomor surat
- Lampiran
- Perihal
- Kepada Yth.
- Pembuka
- Isi surat
- Penutup
- Hormat kami,
- (Nama & Jabatan)
"""
            )
            
            if not ai_response.success:
                raise Exception(ai_response.error)
            
            embed = discord.Embed(
                title="✉️ Surat Resmi",
                description=Helpers.truncate(ai_response.content, 4000),
                color=COLORS["primary"]
            )
            
            await loading.delete()
            await ctx.send(embed=embed)
            
            await self.bot.log_usage(ctx.author.id, "surat", "writer", True)
            
        except Exception as e:
            logger.error(f"Error writing letter: {e}")
            await loading.delete()
            embed = Helpers.create_error_embed(
                title="Error Membuat Surat",
                description=f"```{str(e)[:500]}```"
            )
            await ctx.send(embed=embed)

# =============================================================================
# SETUP
# =============================================================================

async def setup(bot):
    await bot.add_cog(WriterCog(bot))
