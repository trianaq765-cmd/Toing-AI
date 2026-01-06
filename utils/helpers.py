"""
Helpers - Utility helper functions
"""

import asyncio
import logging
import hashlib
from typing import Any, Optional, List, Dict
from datetime import datetime
import discord

from config import EMOJIS, COLORS

logger = logging.getLogger("office_bot.helpers")

# =============================================================================
# HELPERS CLASS
# =============================================================================

class Helpers:
    """
    General helper utilities
    """
    
    # =========================================================================
    # DISCORD HELPERS
    # =========================================================================
    
    @staticmethod
    def create_embed(
        title: str,
        description: Optional[str] = None,
        color: int = COLORS["primary"],
        fields: Optional[List[Dict[str, Any]]] = None,
        footer: Optional[str] = None,
        thumbnail: Optional[str] = None,
        image: Optional[str] = None
    ) -> discord.Embed:
        """
        Create Discord embed with Indonesian formatting
        """
        embed = discord.Embed(
            title=title,
            description=description,
            color=color,
            timestamp=datetime.utcnow()
        )
        
        if fields:
            for field in fields:
                embed.add_field(
                    name=field.get("name", ""),
                    value=field.get("value", ""),
                    inline=field.get("inline", True)
                )
        
        if footer:
            embed.set_footer(text=footer)
        
        if thumbnail:
            embed.set_thumbnail(url=thumbnail)
        
        if image:
            embed.set_image(url=image)
        
        return embed
    
    @staticmethod
    def create_success_embed(
        title: str,
        description: Optional[str] = None,
        **kwargs
    ) -> discord.Embed:
        """Create success embed"""
        return Helpers.create_embed(
            title=f"{EMOJIS['success']} {title}",
            description=description,
            color=COLORS["success"],
            **kwargs
        )
    
    @staticmethod
    def create_error_embed(
        title: str,
        description: Optional[str] = None,
        **kwargs
    ) -> discord.Embed:
        """Create error embed"""
        return Helpers.create_embed(
            title=f"{EMOJIS['error']} {title}",
            description=description,
            color=COLORS["error"],
            **kwargs
        )
    
    @staticmethod
    def create_warning_embed(
        title: str,
        description: Optional[str] = None,
        **kwargs
    ) -> discord.Embed:
        """Create warning embed"""
        return Helpers.create_embed(
            title=f"{EMOJIS['warning']} {title}",
            description=description,
            color=COLORS["warning"],
            **kwargs
        )
    
    @staticmethod
    def create_loading_embed(message: str = "Memproses...") -> discord.Embed:
        """Create loading embed"""
        return Helpers.create_embed(
            title=f"{EMOJIS['loading']} {message}",
            color=COLORS["info"]
        )
    
    # =========================================================================
    # TEXT HELPERS
    # =========================================================================
    
    @staticmethod
    def truncate(text: str, max_length: int = 1024, suffix: str = "...") -> str:
        """
        Truncate text to max length
        """
        if len(text) <= max_length:
            return text
        return text[:max_length - len(suffix)] + suffix
    
    @staticmethod
    def code_block(text: str, language: str = "") -> str:
        """
        Wrap text in Discord code block
        """
        return f"```{language}\n{text}\n```"
    
    @staticmethod
    def inline_code(text: str) -> str:
        """
        Wrap text in inline code
        """
        return f"`{text}`"
    
    @staticmethod
    def bold(text: str) -> str:
        """Bold text"""
        return f"**{text}**"
    
    @staticmethod
    def italic(text: str) -> str:
        """Italic text"""
        return f"*{text}*"
    
    # =========================================================================
    # LIST HELPERS
    # =========================================================================
    
    @staticmethod
    def chunk_list(lst: List[Any], chunk_size: int) -> List[List[Any]]:
        """
        Split list into chunks
        """
        return [lst[i:i + chunk_size] for i in range(0, len(lst), chunk_size)]
    
    @staticmethod
    def paginate_text(
        text: str,
        max_length: int = 2000,
        code_block: bool = False
    ) -> List[str]:
        """
        Paginate long text for Discord (2000 char limit)
        """
        if len(text) <= max_length:
            return [text]
        
        pages = []
        lines = text.split('\n')
        current_page = ""
        
        for line in lines:
            if len(current_page) + len(line) + 1 > max_length:
                pages.append(current_page)
                current_page = line + "\n"
            else:
                current_page += line + "\n"
        
        if current_page:
            pages.append(current_page)
        
        if code_block:
            pages = [f"```\n{page}\n```" for page in pages]
        
        return pages
    
    # =========================================================================
    # ASYNC HELPERS
    # =========================================================================
    
    @staticmethod
    async def run_with_timeout(coro, timeout: float = 30.0):
        """
        Run coroutine with timeout
        """
        try:
            return await asyncio.wait_for(coro, timeout=timeout)
        except asyncio.TimeoutError:
            logger.error(f"Operation timed out after {timeout}s")
            raise
    
    @staticmethod
    async def safe_delete_message(message: discord.Message, delay: float = 0):
        """
        Safely delete message (ignore if already deleted)
        """
        try:
            if delay > 0:
                await asyncio.sleep(delay)
            await message.delete()
        except discord.NotFound:
            pass
        except Exception as e:
            logger.warning(f"Failed to delete message: {e}")
    
    # =========================================================================
    # DATA HELPERS
    # =========================================================================
    
    @staticmethod
    def generate_hash(data: str) -> str:
        """
        Generate SHA256 hash
        """
        return hashlib.sha256(data.encode()).hexdigest()
    
    @staticmethod
    def safe_get(dictionary: dict, *keys, default=None):
        """
        Safely get nested dictionary value
        
        Example:
            safe_get({"a": {"b": {"c": 1}}}, "a", "b", "c") -> 1
        """
        result = dictionary
        for key in keys:
            if isinstance(result, dict):
                result = result.get(key)
            else:
                return default
            if result is None:
                return default
        return result
    
    # =========================================================================
    # TABLE FORMATTING
    # =========================================================================
    
    @staticmethod
    def create_table(
        headers: List[str],
        rows: List[List[Any]],
        max_width: int = 80
    ) -> str:
        """
        Create simple ASCII table
        """
        # Calculate column widths
        col_widths = [len(str(h)) for h in headers]
        
        for row in rows:
            for i, cell in enumerate(row):
                col_widths[i] = max(col_widths[i], len(str(cell)))
        
        # Limit widths
        col_widths = [min(w, max_width // len(headers)) for w in col_widths]
        
        # Create separator
        separator = "+" + "+".join(["-" * (w + 2) for w in col_widths]) + "+"
        
        # Create header
        header_row = "|" + "|".join([
            f" {str(h)[:w]:<{w}} " for h, w in zip(headers, col_widths)
        ]) + "|"
        
        # Create rows
        table_rows = []
        for row in rows:
            table_row = "|" + "|".join([
                f" {str(cell)[:w]:<{w}} " for cell, w in zip(row, col_widths)
            ]) + "|"
            table_rows.append(table_row)
        
        # Combine
        table = [separator, header_row, separator]
        table.extend(table_rows)
        table.append(separator)
        
        return "\n".join(table)
