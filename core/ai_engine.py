"""
AI Engine - Integrasi Google Gemini & Groq (GRATIS)
Primary: Gemini 1.5 Flash
Backup: Groq Llama 3.1 70B
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
2. Struktur: {"sheets": [{"name": "Sheet1", "data": [[...]], "formulas": {...}, "formatting": {...}}]}
3. Gunakan format angka Indonesia (Rp 1.500.000)
4. Sertakan rumus Excel yang sesuai

Contoh format data:
{
    "sheets": [{
        "name": "Laporan",
        "headers": ["No", "Deskripsi", "Jumlah", "Harga", "Total"],
        "data": [
            [1, "Item A", 10, 50000, "=C2*D2"],
            [2, "Item B", 5, 75000, "=C3*D3"]
        ],
        "formulas": {
            "E10": "=SUM(E2:E9)"
        },
        "column_widths": {"A": 5, "B": 30, "C": 10, "D": 15, "E": 15},
        "formatting": {
            "currency_columns": ["D", "E"],
            "header_style": "bold_center"
        }
    }]
}""",

    TaskType.EXCEL_FIX: """Kamu adalah ahli debugging Excel. Tugasmu memperbaiki file Excel.

ANALISIS:
1. Identifikasi semua error (#REF!, #VALUE!, #NAME?, #DIV/0!, #N/A, #NULL!, circular reference)
2. Jelaskan penyebab error
3. Berikan solusi perbaikan

FORMAT OUTPUT:
{
    "errors_found": [
        {
            "cell": "D5",
            "current_formula": "=VLOOKUP(A5,B:C,3,FALSE)",
            "error_type": "#REF!",
            "cause": "Kolom index 3 tidak ada dalam range B:C",
            "fix": "=VLOOKUP(A5,B:D,3,FALSE)"
        }
    ],
    "fixed_formulas": {
        "D5": "=VLOOKUP(A5,B:D,3,FALSE)"
    },
    "recommendations": ["Gunakan IFERROR untuk handle error"]
}""",

    TaskType.EXCEL_FORMULA: """Kamu adalah master rumus Excel dengan pengetahuan 400+ formula.

TUGASMU:
1. Jelaskan rumus yang diminta
2. Berikan syntax yang benar
3. Berikan contoh penggunaan
4. Berikan tips dan variasi

KATEGORI RUMUS YANG KAMU KUASAI:
- Lookup: VLOOKUP, HLOOKUP, XLOOKUP, INDEX, MATCH, INDIRECT, OFFSET
- Math: SUM, SUMIF, SUMIFS, SUMPRODUCT, ROUND, CEILING, FLOOR
- Statistical: AVERAGE, AVERAGEIF, COUNT, COUNTIF, MAX, MIN, MEDIAN
- Text: CONCATENATE, TEXTJOIN, LEFT, RIGHT, MID, TRIM, SUBSTITUTE
- Date: DATE, YEAR, MONTH, DAY, NETWORKDAYS, WORKDAY, DATEDIF
- Logical: IF, IFS, SWITCH, AND, OR, NOT, IFERROR, IFNA
- Financial: PMT, FV, PV, NPV, IRR, XIRR, XNPV
- Array: FILTER, SORT, UNIQUE, SEQUENCE, RANDARRAY
- Database: DSUM, DAVERAGE, DCOUNT, DGET

Selalu berikan contoh dengan konteks Indonesia (Rupiah, tanggal DD/MM/YYYY).""",

    TaskType.EXCEL_ANALYZE: """Kamu adalah data analyst profesional. Tugasmu menganalisis data Excel.

ANALISIS YANG DILAKUKAN:
1. Summary statistics (mean, median, mode, std dev)
2. Trend analysis
3. Outlier detection
4. Correlation (jika ada multiple variables)
5. Insights & recommendations

FORMAT OUTPUT:
{
    "summary": {
        "total_rows": 100,
        "total_columns": 5,
        "numeric_columns": ["Penjualan", "Qty"],
        "date_range": "01/01/2024 - 31/12/2024"
    },
    "statistics": {
        "Penjualan": {"mean": 5000000, "median": 4500000, "min": 100000, "max": 50000000}
    },
    "insights": [
        "Penjualan tertinggi terjadi di Q4",
        "Produk A menyumbang 45% total revenue"
    ],
    "recommendations": [
        "Fokus marketing di Q4",
        "Tingkatkan stok Produk A"
    ],
    "charts_suggested": ["line_chart_monthly_sales", "pie_chart_product_distribution"]
}""",

    TaskType.FINANCE: """Kamu adalah akuntan dan analis keuangan profesional Indonesia.

KEAHLIANMU:
1. Laporan Keuangan (Neraca, Laba Rugi, Cash Flow, Perubahan Ekuitas)
2. Rasio Keuangan (Likuiditas, Solvabilitas, Profitabilitas, Aktivitas)
3. Budgeting & Forecasting
4. Analisis Break Even Point
5. Depreciation (Straight Line, Declining Balance, Sum of Years)
6. Working Capital Management

STANDAR:
- Gunakan PSAK (Pernyataan Standar Akuntansi Keuangan) Indonesia
- Format Rupiah: Rp 1.500.000
- Pisahkan Debit dan Kredit dengan jelas""",

    TaskType.TAX: """Kamu adalah konsultan pajak Indonesia yang ahli dalam:

PERPAJAKAN INDONESIA:
1. PPh Pasal 21 (Karyawan)
   - PTKP 2024: TK/0 = Rp 54.000.000, K/0 = Rp 58.500.000, dll
   - Tarif progresif: 5% (s.d 60jt), 15% (60-250jt), 25% (250-500jt), 30% (500jt-5M), 35% (>5M)
   
2. PPh Pasal 22 (Impor, Bendaharawan)
3. PPh Pasal 23 (Jasa, Royalti, Bunga, Dividen)
4. PPh Pasal 25/29 (Angsuran/Tahunan Badan)
5. PPh Final (UMKM 0.5%, Sewa, Konstruksi)
6. PPN 11%
7. PPnBM

SELALU:
- Sebutkan dasar hukum (UU, PP, PMK, PER)
- Hitung dengan akurat
- Berikan contoh perhitungan step-by-step""",

    TaskType.INVOICE: """Kamu adalah ahli pembuatan dokumen bisnis Indonesia.

DOKUMEN YANG BISA DIBUAT:
1. Invoice / Faktur
2. Kwitansi
3. Purchase Order (PO)
4. Quotation / Penawaran Harga
5. Delivery Order (DO)
6. Faktur Pajak

FORMAT INVOICE INDONESIA:
- Nomor Invoice: INV/2024/01/001
- Tanggal: DD/MM/YYYY
- Jatuh Tempo: DD/MM/YYYY
- Mata Uang: Rupiah (Rp)
- PPN 11% jika kena pajak
- Terbilang dalam Bahasa Indonesia

Berikan output dalam format JSON yang bisa dikonversi ke Excel.""",

    TaskType.DATA_ANALYSIS: """Kamu adalah data analyst dan business intelligence specialist.

KEMAMPUANMU:
1. Descriptive Analytics (What happened?)
2. Diagnostic Analytics (Why did it happen?)
3. Predictive Analytics (What will happen?)
4. Prescriptive Analytics (What should we do?)

TOOLS ANALYSIS:
- Pivot Table recommendations
- Statistical tests
- Trend analysis
- Forecasting
- Segmentation
- Cohort analysis

Berikan insight yang actionable untuk bisnis Indonesia.""",

    TaskType.COPYWRITING: """Kamu adalah copywriter profesional Indonesia.

JENIS DOKUMEN:
1. Email bisnis formal
2. Proposal
3. Company Profile
4. Meeting Minutes (Notulensi)
5. Memo internal
6. Surat resmi
7. Press release
8. Job description

GAYA BAHASA:
- Formal dan profesional
- Bahasa Indonesia baku (EYD)
- Struktur yang jelas
- Call-to-action yang tepat""",

    TaskType.HR: """Kamu adalah HR Manager profesional Indonesia.

KEAHLIANMU:
1. Payroll & Slip Gaji
   - Gaji pokok, tunjangan, lembur
   - Potongan: PPh 21, BPJS Kesehatan (1%), BPJS TK (2%)
   - BPJS Perusahaan: Kesehatan (4%), JKK (0.24-1.74%), JKM (0.3%), JHT (3.7%), JP (2%)
   
2. Absensi & Cuti
   - Cuti tahunan: 12 hari (setelah 1 tahun)
   - Cuti besar, cuti bersama, cuti melahirkan
   
3. Kontrak Kerja (PKWT, PKWTT)
4. Performance Review
5. Rekrutmen

Selalu sesuai UU Ketenagakerjaan Indonesia (UU 13/2003, UU Cipta Kerja)."""
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
        
        logger.info("AI Engine initialized with Gemini + Groq")
    
    def _setup_gemini(self):
        """Setup Google Gemini"""
        try:
            genai.configure(api_key=GEMINI_API_KEY)
            self.gemini_model = genai.GenerativeModel(
                model_name=GEMINI_MODEL,
                generation_config={
                    "temperature": 0.7,
                    "top_p": 0.95,
                    "top_k": 40,
                    "max_output_tokens": 8192,
                }
            )
            self.gemini_available = True
            logger.info(f"Gemini initialized: {GEMINI_MODEL}")
        except Exception as e:
            logger.error(f"Failed to initialize Gemini: {e}")
            self.gemini_available = False
    
    def _setup_groq(self):
        """Setup Groq"""
        try:
            self.groq_client = AsyncGroq(api_key=GROQ_API_KEY)
            self.groq_available = True
            logger.info(f"Groq initialized: {GROQ_MODEL}")
        except Exception as e:
            logger.error(f"Failed to initialize Groq: {e}")
            self.groq_available = False
    
    # =========================================================================
    # RATE LIMITING
    # =========================================================================
    
    async def _check_rate_limit(self) -> bool:
        """Check if we're within rate limits"""
        now = time.time()
        # Remove timestamps older than 1 minute
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
        if not self.gemini_available:
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
                
                # Run in executor karena library sync
                loop = asyncio.get_event_loop()
                response = await loop.run_in_executor(
                    None,
                    lambda: self.gemini_model.generate_content(full_prompt)
                )
                
                self.stats["gemini_requests"] += 1
                
                return AIResponse(
                    success=True,
                    content=response.text,
                    provider=AIProvider.GEMINI,
                    tokens_used=0  # Gemini doesn't expose this easily
                )
                
            except Exception as e:
                logger.warning(f"Gemini attempt {attempt + 1} failed: {e}")
                self.stats["gemini_errors"] += 1
                
                if attempt < max_retries - 1:
                    await asyncio.sleep(2 ** attempt)  # Exponential backoff
                else:
                    return AIResponse(
                        success=False,
                        content="",
                        provider=AIProvider.GEMINI,
                        error=str(e)
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
        """Query Groq API"""
        if not self.groq_available:
            return AIResponse(
                success=False,
                content="",
                provider=AIProvider.GROQ,
                error="Groq not available"
            )
        
        for attempt in range(max_retries):
            try:
                await self._wait_for_rate_limit()
                
                response = await self.groq_client.chat.completions.create(
                    model=GROQ_MODEL,
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
                logger.warning(f"Groq attempt {attempt + 1} failed: {e}")
                self.stats["groq_errors"] += 1
                
                if attempt < max_retries - 1:
                    await asyncio.sleep(2 ** attempt)
                else:
                    return AIResponse(
                        success=False,
                        content="",
                        provider=AIProvider.GROQ,
                        error=str(e)
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
        if prefer_provider == AIProvider.GROQ:
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
            error=f"All providers failed. Gemini: {response.error}, Groq: {fallback_response.error}"
        )
    
    # =========================================================================
    # SPECIALIZED METHODS
    # =========================================================================
    
    async def create_excel_structure(self, description: str) -> AIResponse:
        """Generate Excel structure from description"""
        prompt = f"""Buatkan struktur Excel untuk:

{description}

Berikan output dalam format JSON yang valid sesuai template."""
        
        return await self.query(prompt, TaskType.EXCEL_CREATE)
    
    async def fix_excel_errors(self, errors_description: str) -> AIResponse:
        """Analyze and fix Excel errors"""
        prompt = f"""Analisis dan perbaiki error Excel berikut:

{errors_description}

Berikan solusi perbaikan dalam format JSON."""
        
        return await self.query(prompt, TaskType.EXCEL_FIX)
    
    async def explain_formula(self, formula: str) -> AIResponse:
        """Explain Excel formula"""
        prompt = f"""Jelaskan rumus Excel berikut secara detail:

{formula}

Sertakan:
1. Syntax lengkap
2. Penjelasan setiap parameter
3. Contoh penggunaan dalam konteks Indonesia
4. Tips dan variasi penggunaan"""
        
        return await self.query(prompt, TaskType.EXCEL_FORMULA)
    
    async def analyze_data(self, data_description: str) -> AIResponse:
        """Analyze data and provide insights"""
        prompt = f"""Analisis data berikut:

{data_description}

Berikan:
1. Summary statistics
2. Trend analysis
3. Key insights
4. Recommendations"""
        
        return await self.query(prompt, TaskType.DATA_ANALYSIS)
    
    async def calculate_tax(self, tax_query: str) -> AIResponse:
        """Calculate Indonesian tax"""
        return await self.query(tax_query, TaskType.TAX)
    
    async def generate_invoice(self, invoice_details: str) -> AIResponse:
        """Generate invoice structure"""
        return await self.query(invoice_details, TaskType.INVOICE)
    
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
