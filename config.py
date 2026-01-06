"""
Konfigurasi Bot Discord Office Assistant
"""

import os
from dotenv import load_dotenv
from pathlib import Path

# Load environment variables
load_dotenv()

# =============================================================================
# BASE PATHS
# =============================================================================
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
TEMP_DIR = BASE_DIR / "temp"
TEMPLATES_DIR = BASE_DIR / "templates"

# Create directories if not exist
DATA_DIR.mkdir(exist_ok=True)
TEMP_DIR.mkdir(exist_ok=True)

# =============================================================================
# DISCORD SETTINGS
# =============================================================================
DISCORD_TOKEN = os.getenv("DISCORD_TOKEN")
BOT_PREFIX = os.getenv("BOT_PREFIX", "!")
OWNER_ID = int(os.getenv("OWNER_ID", "0"))

# =============================================================================
# AI API SETTINGS
# =============================================================================
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
GROQ_API_KEY = os.getenv("GROQ_API_KEY")

# AI Model Settings
GEMINI_MODEL = "gemini-1.5-flash"
GROQ_MODEL = "llama-3.1-70b-versatile"

# =============================================================================
# BOT SETTINGS
# =============================================================================
DEBUG_MODE = os.getenv("DEBUG_MODE", "False").lower() == "true"
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO")

# File limits
MAX_FILE_SIZE_MB = 25  # Discord limit
MAX_EXCEL_ROWS = 100000
MAX_EXCEL_COLS = 100

# Rate limiting
AI_RATE_LIMIT = 60  # requests per minute
COOLDOWN_SECONDS = 3

# =============================================================================
# INDONESIA SETTINGS
# =============================================================================
CURRENCY_LOCALE = "id_ID"
CURRENCY_SYMBOL = "Rp"
DATE_FORMAT = "%d/%m/%Y"
DATETIME_FORMAT = "%d/%m/%Y %H:%M"

# PTKP 2024 (Penghasilan Tidak Kena Pajak)
PTKP_2024 = {
    "TK/0": 54_000_000,   # Tidak Kawin, tanpa tanggungan
    "TK/1": 58_500_000,   # Tidak Kawin, 1 tanggungan
    "TK/2": 63_000_000,   # Tidak Kawin, 2 tanggungan
    "TK/3": 67_500_000,   # Tidak Kawin, 3 tanggungan
    "K/0": 58_500_000,    # Kawin, tanpa tanggungan
    "K/1": 63_000_000,    # Kawin, 1 tanggungan
    "K/2": 67_500_000,    # Kawin, 2 tanggungan
    "K/3": 72_000_000,    # Kawin, 3 tanggungan
    "K/I/0": 112_500_000, # Kawin, istri bekerja, tanpa tanggungan
    "K/I/1": 117_000_000, # Kawin, istri bekerja, 1 tanggungan
    "K/I/2": 121_500_000, # Kawin, istri bekerja, 2 tanggungan
    "K/I/3": 126_000_000, # Kawin, istri bekerja, 3 tanggungan
}

# PPh 21 Tarif Progresif 2024
PPH21_TARIF = [
    (60_000_000, 0.05),      # 5% untuk 0 - 60 juta
    (250_000_000, 0.15),     # 15% untuk 60 - 250 juta
    (500_000_000, 0.25),     # 25% untuk 250 - 500 juta
    (5_000_000_000, 0.30),   # 30% untuk 500 juta - 5 miliar
    (float('inf'), 0.35),    # 35% untuk > 5 miliar
]

# PPN Rate
PPN_RATE = 0.11  # 11%

# =============================================================================
# EMBED COLORS
# =============================================================================
COLORS = {
    "primary": 0x5865F2,    # Discord Blurple
    "success": 0x57F287,    # Green
    "warning": 0xFEE75C,    # Yellow
    "error": 0xED4245,      # Red
    "info": 0x5865F2,       # Blue
    "excel": 0x217346,      # Excel Green
    "finance": 0xF4B400,    # Gold
    "tax": 0xEA4335,        # Red
    "invoice": 0x4285F4,    # Blue
}

# =============================================================================
# EMOJIS
# =============================================================================
EMOJIS = {
    "success": "✅",
    "error": "❌",
    "warning": "⚠️",
    "loading": "⏳",
    "excel": "📊",
    "file": "📁",
    "money": "💰",
    "tax": "🧾",
    "invoice": "📄",
    "chart": "📈",
    "calculator": "🔢",
    "ai": "🤖",
    "check": "☑️",
}

# =============================================================================
# VALIDATION
# =============================================================================
def validate_config():
    """Validasi konfigurasi yang diperlukan"""
    errors = []
    
    if not DISCORD_TOKEN:
        errors.append("DISCORD_TOKEN tidak ditemukan di .env")
    
    if not GEMINI_API_KEY:
        errors.append("GEMINI_API_KEY tidak ditemukan di .env")
    
    if errors:
        for error in errors:
            print(f"❌ Config Error: {error}")
        return False
    
    return True
