"""
Bot Class - Discord Office Assistant
"""

import discord
from discord.ext import commands
import logging
import aiosqlite
from pathlib import Path
from config import (
    DISCORD_TOKEN, BOT_PREFIX, OWNER_ID,
    DATA_DIR, COLORS, EMOJIS
)

logger = logging.getLogger("office_bot")

# =============================================================================
# BOT CLASS
# =============================================================================
class OfficeBot(commands.Bot):
    """Main bot class for Office Assistant"""
    
    def __init__(self):
        # Intents
        intents = discord.Intents.default()
        intents.message_content = True
        intents.members = True
        intents.guilds = True
        
        # Initialize bot
        super().__init__(
            command_prefix=commands.when_mentioned_or(BOT_PREFIX),
            intents=intents,
            help_command=None,  # Custom help command
            owner_id=OWNER_ID if OWNER_ID else None,
        )
        
        # Bot state
        self.db = None
        self.ai_engine = None
        self.excel_engine = None
        
    # =========================================================================
    # LIFECYCLE
    # =========================================================================
    async def setup_hook(self):
        """Called when bot is starting up"""
        logger.info("Running setup hook...")
        
        # Setup database
        await self.setup_database()
        
        # Load cogs
        await self.load_cogs()
        
        # Initialize engines
        await self.init_engines()
        
        logger.info("Setup hook completed!")
    
    async def start_bot(self):
        """Start the bot"""
        await self.start(DISCORD_TOKEN)
    
    async def close(self):
        """Cleanup when bot is closing"""
        logger.info("Closing bot...")
        
        if self.db:
            await self.db.close()
        
        await super().close()
    
    # =========================================================================
    # DATABASE
    # =========================================================================
    async def setup_database(self):
        """Setup SQLite database"""
        db_path = DATA_DIR / "bot.db"
        self.db = await aiosqlite.connect(db_path)
        
        # Create tables
        await self.db.execute("""
            CREATE TABLE IF NOT EXISTS users (
                user_id INTEGER PRIMARY KEY,
                username TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                usage_count INTEGER DEFAULT 0
            )
        """)
        
        await self.db.execute("""
            CREATE TABLE IF NOT EXISTS usage_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                command TEXT,
                module TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                success BOOLEAN
            )
        """)
        
        await self.db.commit()
        logger.info("Database initialized!")
    
    # =========================================================================
    # COGS
    # =========================================================================
    async def load_cogs(self):
        """Load all cogs"""
        cog_files = [
            "cogs.excel_cog",
            "cogs.finance_cog",
            "cogs.tax_cog",
            "cogs.invoice_cog",
            "cogs.analyst_cog",
            "cogs.writer_cog",
            "cogs.hr_cog",
        ]
        
        for cog in cog_files:
            try:
                await self.load_extension(cog)
                logger.info(f"Loaded cog: {cog}")
            except Exception as e:
                logger.error(f"Failed to load cog {cog}: {e}")
    
    # =========================================================================
    # ENGINES
    # =========================================================================
    async def init_engines(self):
        """Initialize AI and Excel engines"""
        try:
            from core.ai_engine import AIEngine
            from core.excel_engine import ExcelEngine
            
            self.ai_engine = AIEngine()
            self.excel_engine = ExcelEngine()
            
            logger.info("Engines initialized!")
        except Exception as e:
            logger.error(f"Failed to initialize engines: {e}")
    
    # =========================================================================
    # EVENTS
    # =========================================================================
    async def on_ready(self):
        """Called when bot is ready"""
        logger.info(f"Bot is ready!")
        logger.info(f"Logged in as: {self.user.name} ({self.user.id})")
        logger.info(f"Connected to {len(self.guilds)} guild(s)")
        
        # Set presence
        activity = discord.Activity(
            type=discord.ActivityType.listening,
            name=f"{BOT_PREFIX}help | Office Assistant"
        )
        await self.change_presence(activity=activity)
    
    async def on_command_error(self, ctx, error):
        """Global error handler"""
        if isinstance(error, commands.CommandNotFound):
            return
        
        if isinstance(error, commands.MissingPermissions):
            await ctx.send(f"{EMOJIS['error']} Kamu tidak punya permission untuk command ini!")
            return
        
        if isinstance(error, commands.CommandOnCooldown):
            await ctx.send(
                f"{EMOJIS['warning']} Command dalam cooldown. "
                f"Coba lagi dalam {error.retry_after:.1f} detik."
            )
            return
        
        # Log unexpected errors
        logger.error(f"Command error in {ctx.command}: {error}")
        
        embed = discord.Embed(
            title=f"{EMOJIS['error']} Terjadi Error",
            description=f"```{str(error)[:500]}```",
            color=COLORS["error"]
        )
        await ctx.send(embed=embed)
    
    # =========================================================================
    # HELPERS
    # =========================================================================
    async def log_usage(self, user_id: int, command: str, module: str, success: bool):
        """Log command usage to database"""
        try:
            await self.db.execute(
                "INSERT INTO usage_logs (user_id, command, module, success) VALUES (?, ?, ?, ?)",
                (user_id, command, module, success)
            )
            await self.db.execute(
                "UPDATE users SET usage_count = usage_count + 1 WHERE user_id = ?",
                (user_id,)
            )
            await self.db.commit()
        except Exception as e:
            logger.error(f"Failed to log usage: {e}")
