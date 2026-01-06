"""
File Handler - Manage file uploads, downloads, and temporary files
"""

import asyncio
import logging
import aiofiles
import aiohttp
from pathlib import Path
from typing import Optional, Union, List
from datetime import datetime, timedelta
import discord

from config import TEMP_DIR, MAX_FILE_SIZE_MB

logger = logging.getLogger("office_bot.filehandler")

# =============================================================================
# FILE HANDLER CLASS
# =============================================================================

class FileHandler:
    """
    Handle file operations for Discord bot
    """
    
    def __init__(self):
        self.temp_dir = TEMP_DIR
        self.temp_dir.mkdir(exist_ok=True)
        
        # Allowed extensions
        self.allowed_extensions = {
            'excel': ['.xlsx', '.xls', '.xlsm'],
            'csv': ['.csv'],
            'text': ['.txt', '.md'],
            'image': ['.png', '.jpg', '.jpeg'],
        }
        
        logger.info("File Handler initialized")
    
    # =========================================================================
    # UPLOAD FROM DISCORD
    # =========================================================================
    
    async def download_attachment(
        self,
        attachment: discord.Attachment,
        subfolder: Optional[str] = None
    ) -> Path:
        """
        Download Discord attachment to temp directory
        """
        # Validate size
        if attachment.size > MAX_FILE_SIZE_MB * 1024 * 1024:
            raise ValueError(
                f"File terlalu besar! Max {MAX_FILE_SIZE_MB}MB, "
                f"file kamu {attachment.size / 1024 / 1024:.2f}MB"
            )
        
        # Validate extension
        file_ext = Path(attachment.filename).suffix.lower()
        if not self._is_allowed_extension(file_ext):
            raise ValueError(
                f"Tipe file '{file_ext}' tidak didukung. "
                f"Allowed: {self._get_allowed_extensions_str()}"
            )
        
        # Create subfolder if specified
        save_dir = self.temp_dir
        if subfolder:
            save_dir = self.temp_dir / subfolder
            save_dir.mkdir(exist_ok=True)
        
        # Generate unique filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_filename = self._sanitize_filename(attachment.filename)
        file_path = save_dir / f"{timestamp}_{safe_filename}"
        
        # Download
        try:
            await attachment.save(file_path)
            logger.info(f"Downloaded: {file_path}")
            return file_path
        except Exception as e:
            logger.error(f"Failed to download attachment: {e}")
            raise
    
    async def download_from_url(
        self,
        url: str,
        filename: Optional[str] = None
    ) -> Path:
        """
        Download file from URL
        """
        if filename is None:
            filename = f"download_{datetime.now().strftime('%Y%m%d_%H%M%S')}.tmp"
        
        file_path = self.temp_dir / filename
        
        try:
            async with aiohttp.ClientSession() as session:
                async with session.get(url) as response:
                    if response.status != 200:
                        raise Exception(f"Failed to download: HTTP {response.status}")
                    
                    # Check size
                    content_length = response.headers.get('content-length')
                    if content_length:
                        size_mb = int(content_length) / 1024 / 1024
                        if size_mb > MAX_FILE_SIZE_MB:
                            raise ValueError(f"File terlalu besar: {size_mb:.2f}MB")
                    
                    # Download
                    async with aiofiles.open(file_path, 'wb') as f:
                        await f.write(await response.read())
            
            logger.info(f"Downloaded from URL: {file_path}")
            return file_path
            
        except Exception as e:
            logger.error(f"Failed to download from URL: {e}")
            raise
    
    # =========================================================================
    # FILE VALIDATION
    # =========================================================================
    
    def _is_allowed_extension(self, ext: str) -> bool:
        """Check if file extension is allowed"""
        ext = ext.lower()
        for category, extensions in self.allowed_extensions.items():
            if ext in extensions:
                return True
        return False
    
    def _get_allowed_extensions_str(self) -> str:
        """Get string of allowed extensions"""
        all_exts = []
        for extensions in self.allowed_extensions.values():
            all_exts.extend(extensions)
        return ", ".join(all_exts)
    
    def get_file_category(self, file_path: Union[str, Path]) -> Optional[str]:
        """Get file category based on extension"""
        ext = Path(file_path).suffix.lower()
        for category, extensions in self.allowed_extensions.items():
            if ext in extensions:
                return category
        return None
    
    def _sanitize_filename(self, filename: str) -> str:
        """Sanitize filename for safe storage"""
        # Remove/replace unsafe characters
        unsafe_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|']
        safe_filename = filename
        for char in unsafe_chars:
            safe_filename = safe_filename.replace(char, '_')
        return safe_filename
    
    # =========================================================================
    # FILE INFO
    # =========================================================================
    
    def get_file_info(self, file_path: Union[str, Path]) -> dict:
        """Get file information"""
        file_path = Path(file_path)
        
        if not file_path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")
        
        stat = file_path.stat()
        
        return {
            "name": file_path.name,
            "size": stat.st_size,
            "size_mb": stat.st_size / 1024 / 1024,
            "extension": file_path.suffix,
            "category": self.get_file_category(file_path),
            "created": datetime.fromtimestamp(stat.st_ctime),
            "modified": datetime.fromtimestamp(stat.st_mtime),
        }
    
    # =========================================================================
    # CLEANUP
    # =========================================================================
    
    async def cleanup_old_files(self, max_age_hours: int = 24):
        """
        Delete temporary files older than max_age_hours
        """
        cutoff_time = datetime.now() - timedelta(hours=max_age_hours)
        deleted_count = 0
        
        for file_path in self.temp_dir.rglob("*"):
            if file_path.is_file():
                modified_time = datetime.fromtimestamp(file_path.stat().st_mtime)
                if modified_time < cutoff_time:
                    try:
                        file_path.unlink()
                        deleted_count += 1
                    except Exception as e:
                        logger.warning(f"Failed to delete {file_path}: {e}")
        
        logger.info(f"Cleaned up {deleted_count} old files")
        return deleted_count
    
    async def delete_file(self, file_path: Union[str, Path]):
        """Delete a specific file"""
        file_path = Path(file_path)
        try:
            if file_path.exists():
                file_path.unlink()
                logger.info(f"Deleted: {file_path}")
        except Exception as e:
            logger.error(f"Failed to delete {file_path}: {e}")
            raise
    
    async def cleanup_user_files(self, user_id: int):
        """Delete all files for a specific user"""
        pattern = f"*_{user_id}_*"
        deleted = 0
        
        for file_path in self.temp_dir.glob(pattern):
            try:
                file_path.unlink()
                deleted += 1
            except Exception as e:
                logger.warning(f"Failed to delete {file_path}: {e}")
        
        return deleted
    
    # =========================================================================
    # DISCORD HELPERS
    # =========================================================================
    
    def create_discord_file(
        self,
        file_path: Union[str, Path],
        filename: Optional[str] = None
    ) -> discord.File:
        """
        Create Discord File object for sending
        """
        file_path = Path(file_path)
        
        if not file_path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")
        
        if filename is None:
            filename = file_path.name
        
        return discord.File(str(file_path), filename=filename)
    
    async def send_file_to_channel(
        self,
        channel: discord.TextChannel,
        file_path: Union[str, Path],
        message: Optional[str] = None,
        embed: Optional[discord.Embed] = None
    ):
        """
        Send file to Discord channel
        """
        discord_file = self.create_discord_file(file_path)
        
        await channel.send(
            content=message,
            file=discord_file,
            embed=embed
        )
    
    # =========================================================================
    # UTILITIES
    # =========================================================================
    
    def get_temp_path(self, filename: str) -> Path:
        """Get path for temp file"""
        return self.temp_dir / filename
    
    def ensure_dir(self, dir_path: Union[str, Path]) -> Path:
        """Ensure directory exists"""
        dir_path = Path(dir_path)
        dir_path.mkdir(parents=True, exist_ok=True)
        return dir_path
