"""
Analyst Cog - Data analysis commands
"""

import logging
from typing import Optional

import discord
from discord.ext import commands
import pandas as pd

from config import EMOJIS, COLORS
from utils.formatters import Formatters
from utils.helpers import Helpers

logger = logging.getLogger("office_bot.analyst")

# =============================================================================
# ANALYST COG
# =============================================================================

class AnalystCog(commands.Cog):
    """
    Commands untuk data analysis
    """
    
    def __init__(self, bot):
        self.bot = bot
        logger.info("Analyst Cog loaded")
    
    # =========================================================================
    # ANALYSIS COMMANDS
    # =========================================================================
    
    @commands.command(name="analyze", aliases=["analisa"])
    @commands.cooldown(1, 10, commands.BucketType.user)
    async def analyze_data(self, ctx):
        """
        Analisis data dari file Excel
        
        Usage: Upload Excel file + !analyze
        """
        # Check for attachment
        if not ctx.message.attachments:
            embed = Helpers.create_error_embed(
                title="File Tidak Ditemukan",
                description="Upload file Excel bersamaan dengan command ini!"
            )
            await ctx.send(embed=embed)
            return
        
        loading = await ctx.send(embed=Helpers.create_loading_embed("Menganalisis data..."))
        
        try:
            # Download file
            attachment = ctx.message.attachments[0]
            file_path = await self.bot.file_handler.download_attachment(
                attachment,
                subfolder="analysis"
            )
            
            # Read Excel
            excel_data = await self.bot.excel_engine.read_excel(file_path)
            
            # Convert to DataFrame for analysis
            df = await self.bot.excel_engine.excel_to_dataframe(file_path)
            
            # Get AI analysis
            data_summary = self._create_data_summary(df)
            
            ai_response = await self.bot.ai_engine.analyze_data(
                f"""Analisis data berikut dan berikan insight:

{data_summary}

Data preview (5 baris pertama):
{df.head().to_string()}

Berikan:
1. Summary statistics
2. Trend analysis (jika ada time series)
3. Key insights dan findings
4. Rekomendasi actionable
"""
            )
            
            if not ai_response.success:
                raise Exception(ai_response.error)
            
            # Create result embed
            embed = Helpers.create_success_embed(
                title="Hasil Analisis Data",
                description=f"**File:** {attachment.filename}\n"
                           f"**Rows:** {len(df):,} | **Columns:** {len(df.columns)}"
            )
            
            # Add AI insights
            insights = Helpers.truncate(ai_response.content, 1024)
            embed.add_field(
                name=f"{EMOJIS['chart']} Insights & Recommendations",
                value=insights,
                inline=False
            )
            
            # Add statistics
            stats = self._get_basic_stats(df)
            embed.add_field(
                name=f"{EMOJIS['calculator']} Statistics",
                value=stats,
                inline=False
            )
            
            await loading.delete()
            await ctx.send(embed=embed)
            
            await self.bot.log_usage(ctx.author.id, "analyze", "analyst", True)
            await self.bot.file_handler.delete_file(file_path)
            
        except Exception as e:
            logger.error(f"Error analyzing data: {e}")
            await loading.delete()
            embed = Helpers.create_error_embed(
                title="Error Analisis Data",
                description=f"```{str(e)[:500]}```"
            )
            await ctx.send(embed=embed)
    
    @commands.command(name="summary", aliases=["ringkasan"])
    @commands.cooldown(1, 5, commands.BucketType.user)
    async def data_summary(self, ctx):
        """
        Ringkasan statistik dari data Excel
        
        Usage: Upload Excel + !summary
        """
        if not ctx.message.attachments:
            embed = Helpers.create_error_embed(
                title="File Tidak Ditemukan",
                description="Upload file Excel bersamaan dengan command ini!"
            )
            await ctx.send(embed=embed)
            return
        
        loading = await ctx.send(embed=Helpers.create_loading_embed("Membuat ringkasan..."))
        
        try:
            attachment = ctx.message.attachments[0]
            file_path = await self.bot.file_handler.download_attachment(attachment)
            
            df = await self.bot.excel_engine.excel_to_dataframe(file_path)
            
            # Create summary
            embed = discord.Embed(
                title=f"{EMOJIS['excel']} Data Summary",
                color=COLORS["excel"]
            )
            
            # Basic info
            embed.add_field(
                name="📋 Info",
                value=f"Rows: {len(df):,}\n"
                      f"Columns: {len(df.columns)}\n"
                      f"Size: {Formatters.format_file_size(file_path.stat().st_size)}",
                inline=True
            )
            
            # Column info
            numeric_cols = df.select_dtypes(include=['number']).columns
            text_cols = df.select_dtypes(include=['object']).columns
            
            embed.add_field(
                name="📊 Columns",
                value=f"Numeric: {len(numeric_cols)}\n"
                      f"Text: {len(text_cols)}\n"
                      f"Total: {len(df.columns)}",
                inline=True
            )
            
            # Missing values
            missing = df.isnull().sum().sum()
            embed.add_field(
                name="⚠️ Data Quality",
                value=f"Missing values: {missing:,}\n"
                      f"Duplicates: {df.duplicated().sum():,}",
                inline=True
            )
            
            # Statistics for numeric columns
            if len(numeric_cols) > 0:
                stats_text = ""
                for col in numeric_cols[:3]:  # First 3 numeric columns
                    mean_val = df[col].mean()
                    stats_text += f"**{col}:**\n"
                    stats_text += f"  Mean: {Formatters.format_number(mean_val, 2)}\n"
                    stats_text += f"  Min: {Formatters.format_number(df[col].min(), 2)}\n"
                    stats_text += f"  Max: {Formatters.format_number(df[col].max(), 2)}\n\n"
                
                embed.add_field(
                    name="📈 Numeric Summary",
                    value=Helpers.truncate(stats_text, 1024),
                    inline=False
                )
            
            await loading.delete()
            await ctx.send(embed=embed)
            
            await self.bot.log_usage(ctx.author.id, "summary", "analyst", True)
            await self.bot.file_handler.delete_file(file_path)
            
        except Exception as e:
            logger.error(f"Error creating summary: {e}")
            await loading.delete()
            embed = Helpers.create_error_embed(
                title="Error Membuat Summary",
                description=f"```{str(e)[:500]}```"
            )
            await ctx.send(embed=embed)
    
    # =========================================================================
    # HELPER METHODS
    # =========================================================================
    
    def _create_data_summary(self, df: pd.DataFrame) -> str:
        """Create text summary of DataFrame"""
        summary = f"Dataset dengan {len(df)} rows dan {len(df.columns)} columns\n\n"
        
        summary += "Columns:\n"
        for col in df.columns:
            dtype = df[col].dtype
            null_count = df[col].isnull().sum()
            summary += f"- {col} ({dtype}): {null_count} missing values\n"
        
        return summary
    
    def _get_basic_stats(self, df: pd.DataFrame) -> str:
        """Get basic statistics text"""
        numeric_cols = df.select_dtypes(include=['number']).columns
        
        if len(numeric_cols) == 0:
            return "Tidak ada kolom numerik untuk dianalisis"
        
        stats = "```\n"
        for col in numeric_cols[:5]:  # Max 5 columns
            stats += f"{col}:\n"
            stats += f"  Count: {df[col].count():,}\n"
            stats += f"  Mean:  {df[col].mean():.2f}\n"
            stats += f"  Std:   {df[col].std():.2f}\n"
            stats += f"  Min:   {df[col].min():.2f}\n"
            stats += f"  Max:   {df[col].max():.2f}\n\n"
        stats += "```"
        
        return stats

# =============================================================================
# SETUP
# =============================================================================

async def setup(bot):
    await bot.add_cog(AnalystCog(bot))
