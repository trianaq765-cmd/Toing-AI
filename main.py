"""
Discord Office Assistant Bot - Entry Point
Bot profesional untuk pekerjaan perkantoran Indonesia
"""

import asyncio
import logging
import sys
import os
from aiohttp import web
from bot import OfficeBot
from config import validate_config, LOG_LEVEL, DEBUG_MODE

# =============================================================================
# LOGGING SETUP
# =============================================================================
def setup_logging():
    """Setup logging configuration"""
    log_format = "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s"
    date_format = "%Y-%m-%d %H:%M:%S"
    
    logging.basicConfig(
        level=getattr(logging, LOG_LEVEL.upper()),
        format=log_format,
        datefmt=date_format,
        handlers=[
            logging.StreamHandler(sys.stdout)
        ]
    )
    
    # Reduce noise from libraries
    logging.getLogger("discord").setLevel(logging.WARNING)
    logging.getLogger("discord.http").setLevel(logging.WARNING)
    logging.getLogger("urllib3").setLevel(logging.WARNING)
    logging.getLogger("asyncio").setLevel(logging.WARNING)
    logging.getLogger("aiohttp").setLevel(logging.WARNING)
    
    return logging.getLogger("office_bot")

# =============================================================================
# HTTP SERVER (untuk Render Web Service)
# =============================================================================
async def health_check(request):
    """Health check endpoint"""
    return web.json_response({
        "status": "healthy",
        "service": "Discord Office Bot",
        "message": "Bot is running!"
    })

async def home(request):
    """Home endpoint"""
    return web.Response(
        text="🤖 Discord Office Assistant Bot is running!",
        content_type="text/plain"
    )

async def start_webserver(logger):
    """Start HTTP server for Render"""
    app = web.Application()
    app.router.add_get("/", home)
    app.router.add_get("/health", health_check)
    
    port = int(os.getenv("PORT", 10000))
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", port)
    await site.start()
    
    logger.info(f"HTTP server started on port {port}")
    return runner

# =============================================================================
# MAIN
# =============================================================================
async def main():
    """Main entry point"""
    logger = setup_logging()
    
    # Banner
    print("""
    ╔═══════════════════════════════════════════════════════════════╗
    ║                                                               ║
    ║     🤖 DISCORD OFFICE ASSISTANT BOT                          ║
    ║     📊 Excel Master | 💼 Multi-Profesi | 🇮🇩 Indonesia        ║
    ║                                                               ║
    ║     Powered by: Gemini AI + Groq (FREE)                      ║
    ║     Hosted on: Render (FREE)                                 ║
    ║                                                               ║
    ╚═══════════════════════════════════════════════════════════════╝
    """)
    
    # Validate configuration
    logger.info("Validating configuration...")
    if not validate_config():
        logger.error("Configuration validation failed!")
        sys.exit(1)
    
    logger.info("Configuration validated successfully!")
    
    # Start HTTP server (untuk Render)
    webserver = await start_webserver(logger)
    
    # Create and run bot
    bot = OfficeBot()
    
    try:
        logger.info("Starting Discord bot...")
        await bot.start_bot()
    except KeyboardInterrupt:
        logger.info("Received keyboard interrupt, shutting down...")
    except Exception as e:
        logger.error(f"Fatal error: {e}")
        raise
    finally:
        await webserver.cleanup()
        await bot.close()

# =============================================================================
# RUN
# =============================================================================
if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\n👋 Bot stopped!")
