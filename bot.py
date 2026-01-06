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
        self.file_handler = None
        
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
        
        # Add basic commands
        self.add_basic_commands()
        
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
            from core.file_handler import FileHandler
            
            self.ai_engine = AIEngine()
            self.excel_engine = ExcelEngine()
            self.file_handler = FileHandler()
            
            logger.info("Engines initialized!")
        except Exception as e:
            logger.error(f"Failed to initialize engines: {e}")
    
    # =========================================================================
    # BASIC COMMANDS
    # =========================================================================
    def add_basic_commands(self):
        """Add basic commands directly to bot"""
        
        @self.command(name="ping")
        async def ping(ctx):
            """Check bot latency"""
            latency = round(self.latency * 1000)
            await ctx.send(f"🏓 Pong! Latency: {latency}ms")
        
        @self.command(name="help", aliases=["bantuan", "h"])
        async def help_command(ctx):
            """Show help message"""
            embed = discord.Embed(
                title="🤖 Office Assistant Bot - Help",
                description="Bot profesional untuk pekerjaan perkantoran Indonesia",
                color=COLORS["primary"]
            )
            
            embed.add_field(
                name="📊 Excel Commands",
                value="`!buat` - Buat Excel\n"
                      "`!perbaiki` - Perbaiki Excel\n"
                      "`!rumus <nama>` - Penjelasan rumus\n"
                      "`!template <type>` - Get template\n"
                      "`!daftarrumus` - List rumus",
                inline=True
            )
            
            embed.add_field(
                name="💰 Finance Commands",
                value="`!neraca` - Buat neraca\n"
                      "`!labarugi` - Laporan laba rugi\n"
                      "`!bep` - Break Even Point\n"
                      "`!roi` - Return on Investment\n"
                      "`!depresiasi` - Hitung depresiasi",
                inline=True
            )
            
            embed.add_field(
                name="🧾 Tax Commands",
                value="`!pph21 <gaji> <status>` - PPh 21\n"
                      "`!pph23 <jumlah> <jenis>` - PPh 23\n"
                      "`!ppn <dpp>` - PPN 11%\n"
                      "`!ptkp` - Tabel PTKP\n"
                      "`!infopajak` - Info pajak",
                inline=True
            )
            
            embed.add_field(
                name="📄 Invoice Commands",
                value="`!invoice <details>` - Buat invoice\n"
                      "`!quotation <details>` - Penawaran\n"
                      "`!po <details>` - Purchase Order",
                inline=True
            )
            
            embed.add_field(
                name="👥 HR Commands",
                value="`!gaji <gaji> <status>` - Hitung gaji\n"
                      "`!slipgaji` - Generate slip gaji\n"
                      "`!lembur <gaji> <jam>` - Uang lembur",
                inline=True
            )
            
            embed.add_field(
                name="📝 Writer Commands",
                value="`!email <topic>` - Tulis email\n"
                      "`!memo <topic>` - Buat memo\n"
                      "`!surat <purpose>` - Buat surat\n"
                      "`!proposal <topic>` - Proposal",
                inline=True
            )
            
            embed.add_field(
                name="📈 Analysis Commands",
                value="`!analyze` - Analisis data\n"
                      "`!summary` - Ringkasan statistik",
                inline=True
            )
            
            embed.add_field(
                name="🔧 Utility",
                value="`!ping` - Cek latency\n"
                      "`!help` - Bantuan ini",
                inline=True
            )
            
            embed.set_footer(text=f"Prefix: {BOT_PREFIX} | Powered by Gemini AI")
            
            await ctx.send(embed=embed)
        
        @self.command(name="status")
        async def status(ctx):
            """Check bot status"""
            embed = discord.Embed(
                title="📊 Bot Status",
                color=COLORS["success"]
            )
            
            # AI Status
            ai_status = "✅ Ready" if self.ai_engine and self.ai_engine.is_available() else "❌ Not Available"
            gemini_status = "✅" if self.ai_engine and self.ai_engine.gemini_available else "❌"
            groq_status = "✅" if self.ai_engine and self.ai_engine.groq_available else "❌"
            
            embed.add_field(
                name="🤖 AI Engine",
                value=f"Status: {ai_status}\n"
                      f"Gemini: {gemini_status}\n"
                      f"Groq: {groq_status}",
                inline=True
            )
            
            # Bot info
            embed.add_field(
                name="📡 Connection",
                value=f"Latency: {round(self.latency * 1000)}ms\n"
                      f"Guilds: {len(self.guilds)}\n"
                      f"Prefix: `{BOT_PREFIX}`",
                inline=True
            )
            
            # Cogs
            loaded_cogs = len(self.cogs)
            embed.add_field(
                name="⚙️ Modules",
                value=f"Loaded: {loaded_cogs} cogs",
                inline=True
            )
            
            await ctx.send(embed=embed)
    
    # =========================================================================
    # EVENTS
    # =========================================================================
    async def on_ready(self):
        """Called when bot is ready"""
        logger.info(f"Bot is ready!")
        logger.info(f"Logged in as: {self.user.name} ({self.user.id})")
        logger.info(f"Connected to {len(self.guilds)} guild(s)")
        logger.info(f"Prefix: {BOT_PREFIX}")
        
        # Set presence
        activity = discord.Activity(
            type=discord.ActivityType.listening,
            name=f"{BOT_PREFIX}help | Office Assistant"
        )
        await self.change_presence(activity=activity)
    
    async def on_message(self, message):
        """Called on every message"""
        # Ignore bot messages
        if message.author.bot:
            return
        
        # Debug: Log messages that start with prefix
        if message.content.startswith(BOT_PREFIX):
            logger.info(f"Command received: {message.content} from {message.author}")
        
        # Process commands
        await self.process_commands(message)
    
    async def on_command_error(self, ctx, error):
        """Global error handler"""
        if isinstance(error, commands.CommandNotFound):
            # Optionally send message for unknown commands
            # await ctx.send(f"❌ Command tidak ditemukan. Gunakan `{BOT_PREFIX}help` untuk bantuan.")
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
        
        if isinstance(error, commands.MissingRequiredArgument):
            await ctx.send(
                f"{EMOJIS['error']} Argument tidak lengkap!\n"
                f"Gunakan `{BOT_PREFIX}help` untuk melihat cara pakai."
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
                """INSERT INTO users (user_id, usage_count) VALUES (?, 1)
                   ON CONFLICT(user_id) DO UPDATE SET usage_count = usage_count + 1""",
                (user_id,)
            )
            await self.db.commit()
        except Exception as e:
            logger.error(f"Failed to log usage: {e}")
