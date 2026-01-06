"""
AI Engine - Integrasi Google Gemini & Groq (GRATIS)
Primary: Gemini 2.0 Flash
Backup: Groq Llama 3.3 70B
Updated: January 2025
"""

import asyncio
import logging
import time
from typing import Optional, Dict, Any, List
from dataclasses import dataclass
from enum import Enum

import google.generativeai as genai
from groq import AsyncGroq

from config import (
    GEMINI_API_KEY, GROQ_API_KEY,
    GEMINI_MODEL, GROQ_MODEL,
    AI_RATE_LIMIT
)

logger = logging.getLogger("office_bot.ai")

# =============================================================================
# ENUMS & DATA CLASSES
# =============================================================================

class AIProvider(Enum):
    """AI Provider options"""
    GEMINI = "gemini"
    GROQ = "groq"

class TaskType(Enum):
    """Task types for context"""
    EXCEL_CREATE = "excel_create"
    EXCEL_FIX = "excel_fix"
    EXCEL_ANALYZE = "excel_analyze"
    EXCEL_FORMULA = "excel_formula"
    FINANCE = "finance"
    TAX = "tax"
    INVOICE = "invoice"
    DATA_ANALYSIS = "data_analysis"
    COPYWRITING = "copywriting"
    HR = "hr"
    GENERAL = "general"

@dataclass
class AIResponse:
    """AI Response wrapper"""
    success: bool
    content: str
    provider: AIProvider
    tokens_used: int = 0
    error: Optional[str] = None

# =============================================================================
# SYSTEM PROMPTS
# =============================================================================

SYSTEM_PROMPTS = {
    TaskType.GENERAL: """Kamu adalah asisten perkantoran profesional Indonesia yang ahli dalam:
- Microsoft Excel dan rumus-rumusnya
- Keuangan dan akuntansi
- Perpajakan Indonesia (PPh, PPN)
- Pembuatan invoice dan dokumen bisnis
- Analisis data
- Copywriting bisnis
- HR dan payroll

Selalu jawab dalam Bahasa Indonesia yang profesional.
Format angka menggunakan format Indonesia (titik untuk ribuan, koma untuk desimal).
Contoh: Rp 1.500.000,00""",

    TaskType.EXCEL_CREATE: """Kamu adalah ahli Microsoft Excel Indonesia. Tugasmu membuat struktur Excel.

ATURAN OUTPUT:
1. Berikan respons dalam format JSON yang valid
2. Struktur: {"sheets": [{"name": "Sheet1", "headers": [...], "data": [[...]], "formulas": {...}, "formatting": {...}}]}
3. Gunakan format angka Indonesia (Rp 1.500.000)
4. Sertakan rumus Excel yang sesuai

PENTING: Output HARUS berupa JSON valid, tanpa markdown code block.

Contoh format yang benar:
{"sheets": [{"name": "Laporan", "headers": ["No", "Deskripsi", "Jumlah", "Harga", "Total"], "data": [[1, "Item A", 10, 50000, "=C2*D2"]], "formulas": {"E10": "=SUM(E2:E9)"}, "column_widths": {"A": 5, "B": 30, "C": 10, "D": 15, "E": 15}, "formatting": {"currency_columns": ["D", "E"]}}]}""",

    TaskType.EXCEL_FIX: """Kamu adalah ahli debugging Excel. Tugasmu memperbaiki file Excel.

ANALISIS:
1. Identifikasi semua error (#REF!, #VALUE!, #NAME?, #DIV/0!, #N/A, #NULL!, circular reference)
2. Jelaskan penyebab error
3. Berikan solusi perbaikan

FORMAT OUTPUT (JSON):
{"errors_found": [{"cell": "D5", "error_type": "#REF!", "cause": "...", "fix": "..."}], "fixed_formulas": {"D5": "=VLOOKUP(A5,B:D,3,FALSE)"}, "recommendations": ["..."]}""",

    TaskType.EXCEL_FORMULA: """Kamu adalah master rumus Excel dengan pengetahuan 400+ formula.

TUGASMU:
1. Jelaskan rumus yang diminta
2. Berikan syntax yang benar
3. Berikan contoh penggunaan
4. Berikan tips dan variasi

KATEGORI RUMUS:
- Lookup: VLOOKUP, HLOOKUP, XLOOKUP, INDEX, MATCH
- Math: SUM, SUMIF, SUMIFS, SUMPRODUCT, ROUND
- Statistical: AVERAGE, COUNT, COUNTIF, MAX, MIN
- Text: CONCATENATE, TEXTJOIN, LEFT, RIGHT, MID
- Date: DATE, YEAR, MONTH, DAY, NETWORKDAYS
- Logical: IF, IFS, AND, OR, IFERROR
- Financial: PMT, FV, PV, NPV, IRR

Berikan contoh dengan konteks Indonesia (Rupiah, tanggal DD/MM/YYYY).""",

    TaskType.EXCEL_ANALYZE: """Kamu adalah data analyst profesional.

ANALISIS:
1. Summary statistics
2. Trend analysis
3. Key insights
4. Recommendations

FORMAT OUTPUT (JSON):
{"summary": {...}, "statistics": {...}, "insights": [...], "recommendations": [...]}""",

    TaskType.FINANCE: """Kamu adalah akuntan profesional Indonesia.

KEAHLIAN:
1. Laporan Keuangan (Neraca, Laba Rugi, Cash Flow)
2. Rasio Keuangan
3. Budgeting & Forecasting
4. Break Even Point
5. Depreciation

Gunakan standar PSAK Indonesia. Format Rupiah: Rp 1.500.000""",

    TaskType.TAX: """Kamu adalah konsultan pajak Indonesia.

PERPAJAKAN:
1. PPh 21 - PTKP 2024: TK/0=54jt, K/0=58.5jt, dst
2. Tarif progresif: 5%(s.d 60jt), 15%(60-250jt), 25%(250-500jt), 30%(500jt-5M), 35%(>5M)
3. PPh 23, PPh 25/29, PPh Final
4. PPN 11%

Hitung akurat dan berikan dasar hukum.""",

    TaskType.INVOICE: """Kamu adalah ahli dokumen bisnis Indonesia.

DOKUMEN:
1. Invoice / Faktur
2. Quotation
3. Purchase Order
4. Kwitansi

FORMAT INVOICE:
- Nomor: INV/2024/01/001
- Tanggal: DD/MM/YYYY
- PPN 11%
- Terbilang dalam Bahasa Indonesia

Output dalam format JSON untuk dikonversi ke Excel.""",

    TaskType.DATA_ANALYSIS: """Kamu adalah data analyst.

ANALISIS:
1. Descriptive Analytics
2. Diagnostic Analytics
3. Insights actionable

Berikan insight bisnis yang relevan untuk Indonesia.""",

    TaskType.COPYWRITING: """Kamu adalah copywriter profesional Indonesia.

DOKUMEN:
1. Email bisnis formal
2. Proposal
3. Memo internal
4. Surat resmi
5. Notulensi

Gunakan Bahasa Indonesia baku dan profesional.""",

    TaskType.HR: """Kamu adalah HR Manager Indonesia.

KEAHLIAN:
1. Payroll & Slip Gaji
2. BPJS: Kesehatan 1%, TK 2%
3. PPh 21
4. Lembur: jam 1 = 1.5x, jam 2+ = 2x
5. Cuti tahunan: 12 hari

Sesuai UU Ketenagakerjaan Indonesia."""
}

# =============================================================================
# AI ENGINE CLASS
# =============================================================================

class AIEngine:
    """
    AI Engine dengan dual provider (Gemini + Groq)
    Auto-fallback jika primary gagal
    """
    
    def __init__(self):
        self._setup_gemini()
        self._setup_groq()
        
        # Rate limiting
        self._request_timestamps: List[float] = []
        self._rate_limit = AI_RATE_LIMIT
        
        # Stats
        self.stats = {
            "gemini_requests": 0,
            "groq_requests": 0,
            "gemini_errors": 0,
            "groq_errors": 0,
            "fallbacks": 0
        }
        
        # Expose TaskType for external use
        self.TaskType = TaskType
        
        logger.info("AI Engine initialized with Gemini + Groq")
    
    def _setup_gemini(self):
        """Setup Google Gemini"""
        self.gemini_available = False
        self.gemini_model = None
        
        if not GEMINI_API_KEY:
            logger.warning("GEMINI_API_KEY not set, Gemini will be disabled")
            return
        
        try:
            genai.configure(api_key=GEMINI_API_KEY)
            
            # Try multiple model names (fallback)
            model_names = [
                "gemini-2.0-flash-exp",
                "gemini-1.5-flash-latest", 
                "gemini-1.5-flash",
                "gemini-1.5-pro-latest",
                "gemini-pro"
            ]
            
            for model_name in model_names:
                try:
                    self.gemini_model = genai.GenerativeModel(
                        model_name=model_name,
                        generation_config={
                            "temperature": 0.7,
                            "top_p": 0.95,
                            "top_k": 40,
                            "max_output_tokens": 8192,
                        }
                    )
                    # Test model
                    self.gemini_available = True
                    logger.info(f"Gemini initialized: {model_name}")
                    break
                except Exception as e:
                    logger.warning(f"Model {model_name} not available: {e}")
                    continue
            
            if not self.gemini_available:
                logger.error("No Gemini model available")
                
        except Exception as e:
            logger.error(f"Failed to initialize Gemini: {e}")
    
    def _setup_groq(self):
        """Setup Groq"""
        self.groq_available = False
        self.groq_client = None
        self.groq_model = GROQ_MODEL
        
        if not GROQ_API_KEY:
            logger.warning("GROQ_API_KEY not set, Groq will be disabled")
            return
        
        try:
            self.groq_client = AsyncGroq(api_key=GROQ_API_KEY)
            
            # Available Groq models (2025)
            self.groq_models = [
                "llama-3.3-70b-versatile",
                "llama-3.1-8b-instant",
                "llama3-groq-70b-8192-tool-use-preview",
                "mixtral-8x7b-32768",
                "gemma2-9b-it"
            ]
            
            self.groq_model = self.groq_models[0]  # Primary
            self.groq_available = True
            logger.info(f"Groq initialized: {self.groq_model}")
            
        except Exception as e:
            logger.error(f"Failed to initialize Groq: {e}")
    
    # =========================================================================
    # RATE LIMITING
    # =========================================================================
    
    async def _check_rate_limit(self) -> bool:
        """Check if we're within rate limits"""
        now = time.time()
        self._request_timestamps = [
            ts for ts in self._request_timestamps 
            if now - ts < 60
        ]
        
        if len(self._request_timestamps) >= self._rate_limit:
            return False
        
        self._request_timestamps.append(now)
        return True
    
    async def _wait_for_rate_limit(self):
        """Wait until rate limit allows"""
        while not await self._check_rate_limit():
            await asyncio.sleep(1)
    
    # =========================================================================
    # GEMINI
    # =========================================================================
    
    async def _query_gemini(
        self, 
        prompt: str, 
        system_prompt: str,
        max_retries: int = 3
    ) -> AIResponse:
        """Query Gemini API"""
        if not self.gemini_available or self.gemini_model is None:
            return AIResponse(
                success=False,
                content="",
                provider=AIProvider.GEMINI,
                error="Gemini not available"
            )
        
        full_prompt = f"{system_prompt}\n\n---\n\nUser Request:\n{prompt}"
        
        for attempt in range(max_retries):
            try:
                await self._wait_for_rate_limit()
                
                loop = asyncio.get_event_loop()
                response = await loop.run_in_executor(
                    None,
                    lambda: self.gemini_model.generate_content(full_prompt)
                )
                
                self.stats["gemini_requests"] += 1
                
                # Extract text from response
                response_text = ""
                if hasattr(response, 'text'):
                    response_text = response.text
                elif hasattr(response, 'parts'):
                    response_text = "".join([part.text for part in response.parts if hasattr(part, 'text')])
                
                return AIResponse(
                    success=True,
                    content=response_text,
                    provider=AIProvider.GEMINI,
                    tokens_used=0
                )
                
            except Exception as e:
                logger.warning(f"Gemini attempt {attempt + 1} failed: {e}")
                self.stats["gemini_errors"] += 1
                
                if attempt < max_retries - 1:
                    await asyncio.sleep(2 ** attempt)
        
        return AIResponse(
            success=False,
            content="",
            provider=AIProvider.GEMINI,
            error="Max retries exceeded"
        )
    
    # =========================================================================
    # GROQ
    # =========================================================================
    
    async def _query_groq(
        self,
        prompt: str,
        system_prompt: str,
        max_retries: int = 3
    ) -> AIResponse:
        """Query Groq API with model fallback"""
        if not self.groq_available or self.groq_client is None:
            return AIResponse(
                success=False,
                content="",
                provider=AIProvider.GROQ,
                error="Groq not available"
            )
        
        for attempt in range(max_retries):
            for model in self.groq_models:
                try:
                    await self._wait_for_rate_limit()
                    
                    response = await self.groq_client.chat.completions.create(
                        model=model,
                        messages=[
                            {"role": "system", "content": system_prompt},
                            {"role": "user", "content": prompt}
                        ],
                        temperature=0.7,
                        max_tokens=8192,
                    )
                    
                    self.stats["groq_requests"] += 1
                    
                    return AIResponse(
                        success=True,
                        content=response.choices[0].message.content,
                        provider=AIProvider.GROQ,
                        tokens_used=response.usage.total_tokens if response.usage else 0
                    )
                    
                except Exception as e:
                    error_msg = str(e)
                    
                    # If model deprecated, try next model
                    if "decommissioned" in error_msg.lower() or "not found" in error_msg.lower():
                        logger.warning(f"Groq model {model} not available, trying next...")
                        continue
                    
                    logger.warning(f"Groq attempt with {model} failed: {e}")
                    self.stats["groq_errors"] += 1
                    break
            
            if attempt < max_retries - 1:
                await asyncio.sleep(2 ** attempt)
        
        return AIResponse(
            success=False,
            content="",
            provider=AIProvider.GROQ,
            error="All Groq models failed"
        )
    
    # =========================================================================
    # MAIN QUERY METHOD
    # =========================================================================
    
    async def query(
        self,
        prompt: str,
        task_type: TaskType = TaskType.GENERAL,
        prefer_provider: Optional[AIProvider] = None,
        custom_system_prompt: Optional[str] = None
    ) -> AIResponse:
        """
        Main method untuk query AI
        Auto-fallback dari Gemini ke Groq jika gagal
        """
        system_prompt = custom_system_prompt or SYSTEM_PROMPTS.get(
            task_type, 
            SYSTEM_PROMPTS[TaskType.GENERAL]
        )
        
        # Determine primary provider
        if prefer_provider == AIProvider.GROQ and self.groq_available:
            primary_query = self._query_groq
            fallback_query = self._query_gemini
        else:
            primary_query = self._query_gemini
            fallback_query = self._query_groq
        
        # Try primary
        response = await primary_query(prompt, system_prompt)
        
        if response.success:
            return response
        
        # Fallback
        logger.info(f"Falling back from {response.provider.value} to other provider")
        self.stats["fallbacks"] += 1
        
        fallback_response = await fallback_query(prompt, system_prompt)
        
        if fallback_response.success:
            return fallback_response
        
        # Both failed
        return AIResponse(
            success=False,
            content="",
            provider=AIProvider.GEMINI,
            error=f"All providers failed. Primary: {response.error}, Fallback: {fallback_response.error}"
        )
    
    # =========================================================================
    # SPECIALIZED METHODS
    # =========================================================================
    
    async def create_excel_structure(self, description: str) -> AIResponse:
        """Generate Excel structure from description"""
        prompt = f"""Buatkan struktur Excel untuk:

{description}

PENTING: Output HARUS berupa JSON valid saja, tanpa markdown, tanpa code block, tanpa penjelasan."""
        
        return await self.query(prompt, TaskType.EXCEL_CREATE)
    
    async def fix_excel_errors(self, errors_description: str) -> AIResponse:
        """Analyze and fix Excel errors"""
        prompt = f"""Analisis dan perbaiki error Excel berikut:

{errors_description}

Berikan solusi perbaikan."""
        
        return await self.query(prompt, TaskType.EXCEL_FIX)
    
    async def explain_formula(self, formula: str) -> AIResponse:
        """Explain Excel formula"""
        prompt = f"""Jelaskan rumus Excel berikut secara detail:

{formula}

Sertakan:
1. Syntax lengkap
2. Penjelasan parameter
3. Contoh penggunaan Indonesia
4. Tips penggunaan"""
        
        return await self.query(prompt, TaskType.EXCEL_FORMULA)
    
    async def analyze_data(self, data_description: str) -> AIResponse:
        """Analyze data and provide insights"""
        prompt = f"""Analisis data berikut:

{data_description}

Berikan summary, insights, dan recommendations."""
        
        return await self.query(prompt, TaskType.DATA_ANALYSIS)
    
    async def calculate_tax(self, tax_query: str) -> AIResponse:
        """Calculate Indonesian tax"""
        return await self.query(tax_query, TaskType.TAX)
    
    async def generate_invoice(self, invoice_details: str) -> AIResponse:
        """Generate invoice structure"""
        prompt = f"""Buatkan struktur invoice untuk:

{invoice_details}

Output dalam format JSON valid untuk dikonversi ke Excel."""
        
        return await self.query(prompt, TaskType.INVOICE)
    
    async def generate_document(self, doc_request: str) -> AIResponse:
        """Generate business document"""
        return await self.query(doc_request, TaskType.COPYWRITING)
    
    async def hr_calculation(self, hr_query: str) -> AIResponse:
        """HR and payroll calculations"""
        return await self.query(hr_query, TaskType.HR)
    
    async def finance_analysis(self, finance_query: str) -> AIResponse:
        """Financial analysis and calculations"""
        return await self.query(finance_query, TaskType.FINANCE)
    
    # =========================================================================
    # UTILITIES
    # =========================================================================
    
    def get_stats(self) -> Dict[str, Any]:
        """Get AI usage statistics"""
        return {
            **self.stats,
            "gemini_available": self.gemini_available,
            "groq_available": self.groq_available,
            "current_rate": len(self._request_timestamps)
        }
    
    def reset_stats(self):
        """Reset statistics"""
        self.stats = {
            "gemini_requests": 0,
            "groq_requests": 0,
            "gemini_errors": 0,
            "groq_errors": 0,
            "fallbacks": 0
        }
    
    def is_available(self) -> bool:
        """Check if at least one AI provider is available"""
        return self.gemini_available or self.groq_available
