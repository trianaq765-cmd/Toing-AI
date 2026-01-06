"""
PPh 21 Calculator - Detailed Indonesian Income Tax Calculator
"""

import logging
from typing import Dict, Any, Optional, List
from dataclasses import dataclass
from datetime import datetime

from config import PTKP_2024, PPH21_TARIF

logger = logging.getLogger("office_bot.pph21")

# =============================================================================
# DATA CLASSES
# =============================================================================

@dataclass
class EmployeeData:
    """Employee data for PPh 21 calculation"""
    nik: str
    name: str
    gaji_pokok: float
    tunjangan_tetap: float = 0
    tunjangan_tidak_tetap: float = 0
    bonus: float = 0
    overtime: float = 0
    ptkp_status: str = "TK/0"
    bpjs_kesehatan_employee: float = 0  # 1%
    bpjs_tk_employee: float = 0  # 2%
    iuran_pensiun: float = 0

@dataclass
class PPh21Result:
    """PPh 21 calculation result"""
    bruto_setahun: float
    pengurang_setahun: float
    netto_setahun: float
    ptkp: float
    pkp: float
    pph21_setahun: float
    pph21_sebulan: float
    tax_breakdown: List[Dict[str, Any]]
    effective_rate: float

# =============================================================================
# PPH 21 CALCULATOR
# =============================================================================

class PPh21Calculator:
    """
    Indonesian Income Tax (PPh 21) Calculator
    Based on UU HPP 2022
    """
    
    def __init__(self):
        self.ptkp_rates = PTKP_2024
        self.tax_brackets = PPH21_TARIF
        logger.info("PPh 21 Calculator initialized")
    
    # =========================================================================
    # MAIN CALCULATION
    # =========================================================================
    
    def calculate(self, employee: EmployeeData) -> PPh21Result:
        """
        Calculate PPh 21 for employee
        
        Method: Gross Up (company bears the tax) or Nett (employee bears)
        This implementation uses Nett method
        """
        
        # 1. Calculate annual gross income
        bruto_setahun = self._calculate_annual_gross(employee)
        
        # 2. Calculate deductions
        pengurang_setahun = self._calculate_deductions(employee, bruto_setahun)
        
        # 3. Net income
        netto_setahun = bruto_setahun - pengurang_setahun
        
        # 4. Get PTKP
        ptkp = self.ptkp_rates.get(employee.ptkp_status, 54_000_000)
        
        # 5. PKP (Penghasilan Kena Pajak)
        pkp = max(0, netto_setahun - ptkp)
        
        # 6. Calculate PPh 21
        pph21_setahun, breakdown = self._calculate_progressive_tax(pkp)
        
        # 7. Monthly PPh 21
        pph21_sebulan = pph21_setahun / 12
        
        # 8. Effective rate
        effective_rate = (pph21_setahun / bruto_setahun * 100) if bruto_setahun > 0 else 0
        
        return PPh21Result(
            bruto_setahun=bruto_setahun,
            pengurang_setahun=pengurang_setahun,
            netto_setahun=netto_setahun,
            ptkp=ptkp,
            pkp=pkp,
            pph21_setahun=pph21_setahun,
            pph21_sebulan=pph21_sebulan,
            tax_breakdown=breakdown,
            effective_rate=effective_rate
        )
    
    def calculate_simple(
        self, 
        gaji_setahun: float, 
        ptkp_status: str = "TK/0"
    ) -> PPh21Result:
        """Simplified calculation with just annual salary"""
        
        employee = EmployeeData(
            nik="",
            name="",
            gaji_pokok=gaji_setahun / 12,
            ptkp_status=ptkp_status
        )
        
        return self.calculate(employee)
    
    # =========================================================================
    # CALCULATION HELPERS
    # =========================================================================
    
    def _calculate_annual_gross(self, employee: EmployeeData) -> float:
        """Calculate annual gross income"""
        
        # Monthly components
        monthly_regular = (
            employee.gaji_pokok +
            employee.tunjangan_tetap
        )
        
        # Annual from monthly
        annual_from_monthly = monthly_regular * 12
        
        # Add irregular income (bonus, overtime, etc)
        irregular = (
            employee.tunjangan_tidak_tetap +
            employee.bonus +
            employee.overtime
        )
        
        return annual_from_monthly + irregular
    
    def _calculate_deductions(self, employee: EmployeeData, bruto: float) -> float:
        """
        Calculate allowable deductions (pengurang)
        
        Deductions:
        1. Biaya Jabatan: 5% dari bruto, max Rp 6.000.000/tahun
        2. Iuran Pensiun
        3. BPJS contributions (employee portion)
        """
        
        # 1. Biaya Jabatan (5% max 6jt/year)
        biaya_jabatan = min(bruto * 0.05, 6_000_000)
        
        # 2. Pension contributions
        iuran_pensiun_setahun = employee.iuran_pensiun * 12
        
        # 3. BPJS employee contributions (already deducted from gross)
        # These are already in the employee data
        bpjs_employee_setahun = (
            employee.bpjs_kesehatan_employee * 12 +
            employee.bpjs_tk_employee * 12
        )
        
        total_pengurang = (
            biaya_jabatan +
            iuran_pensiun_setahun +
            bpjs_employee_setahun
        )
        
        return total_pengurang
    
    def _calculate_progressive_tax(self, pkp: float) -> tuple[float, List[Dict[str, Any]]]:
        """
        Calculate tax using progressive rates
        Returns: (total_tax, breakdown)
        """
        
        if pkp <= 0:
            return 0, []
        
        total_tax = 0
        breakdown = []
        remaining = pkp
        previous_bracket = 0
        
        for i, (bracket, rate) in enumerate(self.tax_brackets):
            # Calculate taxable amount in this bracket
            if bracket == float('inf'):
                taxable_in_bracket = remaining
            else:
                taxable_in_bracket = min(remaining, bracket - previous_bracket)
            
            if taxable_in_bracket > 0:
                tax_in_bracket = taxable_in_bracket * rate
                total_tax += tax_in_bracket
                
                # Add to breakdown
                breakdown.append({
                    "layer": i + 1,
                    "bracket_min": previous_bracket,
                    "bracket_max": bracket,
                    "taxable": taxable_in_bracket,
                    "rate": rate,
                    "tax": tax_in_bracket
                })
                
                remaining -= taxable_in_bracket
            
            if remaining <= 0:
                break
            
            previous_bracket = bracket
        
        return total_tax, breakdown
    
    # =========================================================================
    # BPJS CALCULATIONS
    # =========================================================================
    
    def calculate_bpjs_kesehatan(self, gaji: float) -> Dict[str, float]:
        """
        Calculate BPJS Kesehatan
        
        Employee: 1% of salary (max dari upah Rp 12jt)
        Employer: 4% of salary (max dari upah Rp 12jt)
        """
        max_salary = 12_000_000
        base_salary = min(gaji, max_salary)
        
        return {
            "employee": base_salary * 0.01,
            "employer": base_salary * 0.04,
            "total": base_salary * 0.05
        }
    
    def calculate_bpjs_ketenagakerjaan(self, gaji: float) -> Dict[str, float]:
        """
        Calculate BPJS Ketenagakerjaan (BPJS TK)
        
        Components:
        - JKK (Kecelakaan Kerja): 0.24% - 1.74% (employer) - assume 0.54%
        - JKM (Kematian): 0.30% (employer)
        - JHT (Hari Tua): 3.7% (employer) + 2% (employee)
        - JP (Pensiun): 2% (employer) + 1% (employee) - for salary > 10jt
        """
        
        # JHT: max dari upah Rp 9.5jt (updated 2024)
        jht_base = min(gaji, 9_500_000)
        
        # JP: max dari upah Rp 10jt
        jp_base = min(gaji, 10_000_000) if gaji >= 10_000_000 else 0
        
        employee_total = (
            jht_base * 0.02 +  # JHT employee
            (jp_base * 0.01 if jp_base > 0 else 0)  # JP employee
        )
        
        employer_total = (
            gaji * 0.0054 +    # JKK (assume 0.54%)
            gaji * 0.003 +     # JKM
            jht_base * 0.037 + # JHT employer
            (jp_base * 0.02 if jp_base > 0 else 0)  # JP employer
        )
        
        return {
            "employee": employee_total,
            "employer": employer_total,
            "total": employee_total + employer_total,
            "components": {
                "jkk": gaji * 0.0054,
                "jkm": gaji * 0.003,
                "jht_employee": jht_base * 0.02,
                "jht_employer": jht_base * 0.037,
                "jp_employee": jp_base * 0.01 if jp_base > 0 else 0,
                "jp_employer": jp_base * 0.02 if jp_base > 0 else 0
            }
        }
    
    # =========================================================================
    # UTILITIES
    # =========================================================================
    
    def get_ptkp(self, status: str) -> float:
        """Get PTKP amount by status"""
        return self.ptkp_rates.get(status.upper(), 54_000_000)
    
    def validate_ptkp_status(self, status: str) -> bool:
        """Validate PTKP status"""
        return status.upper() in self.ptkp_rates
    
    def get_all_ptkp_statuses(self) -> Dict[str, float]:
        """Get all PTKP statuses and amounts"""
        return self.ptkp_rates.copy()
    
    def estimate_monthly_pph21(self, gaji_bulanan: float, ptkp_status: str = "TK/0") -> float:
        """Quick estimate monthly PPh 21 from monthly salary"""
        gaji_tahunan = gaji_bulanan * 12
        result = self.calculate_simple(gaji_tahunan, ptkp_status)
        return result.pph21_sebulan
    
    def calculate_take_home_pay(self, employee: EmployeeData) -> Dict[str, float]:
        """
        Calculate take home pay (gaji bersih)
        
        Gross - BPJS employee - PPh 21 = Take Home Pay
        """
        
        # Monthly gross
        gross_monthly = (
            employee.gaji_pokok +
            employee.tunjangan_tetap
        )
        
        # Calculate PPh 21
        pph21_result = self.calculate(employee)
        
        # BPJS
        bpjs_kes = self.calculate_bpjs_kesehatan(gross_monthly)
        bpjs_tk = self.calculate_bpjs_ketenagakerjaan(gross_monthly)
        
        # Total deductions
        total_deductions = (
            bpjs_kes["employee"] +
            bpjs_tk["employee"] +
            pph21_result.pph21_sebulan
        )
        
        # Take home pay
        take_home = gross_monthly - total_deductions
        
        return {
            "gross": gross_monthly,
            "bpjs_kesehatan": bpjs_kes["employee"],
            "bpjs_tk": bpjs_tk["employee"],
            "pph21": pph21_result.pph21_sebulan,
            "total_deductions": total_deductions,
            "take_home_pay": take_home
        }
    
    def format_result_text(self, result: PPh21Result) -> str:
        """Format result as readable text"""
        
        text = "PERHITUNGAN PPh 21\n"
        text += "=" * 50 + "\n\n"
        
        text += f"Penghasilan Bruto Setahun: Rp {result.bruto_setahun:,.0f}\n"
        text += f"Pengurang:                  Rp {result.pengurang_setahun:,.0f}\n"
        text += f"Penghasilan Netto Setahun:  Rp {result.netto_setahun:,.0f}\n"
        text += f"PTKP:                       Rp {result.ptkp:,.0f}\n"
        text += "-" * 50 + "\n"
        text += f"PKP (Penghasilan Kena Pajak): Rp {result.pkp:,.0f}\n\n"
        
        text += "RINCIAN PAJAK PROGRESIF:\n"
        for layer in result.tax_breakdown:
            text += f"  Layer {layer['layer']}: "
            text += f"Rp {layer['taxable']:,.0f} x {layer['rate']*100}% = "
            text += f"Rp {layer['tax']:,.0f}\n"
        
        text += "\n" + "=" * 50 + "\n"
        text += f"PPh 21 SETAHUN:  Rp {result.pph21_setahun:,.0f}\n"
        text += f"PPh 21 PER BULAN: Rp {result.pph21_sebulan:,.0f}\n"
        text += f"Effective Rate:   {result.effective_rate:.2f}%\n"
        
        return text

# =============================================================================
# SINGLETON
# =============================================================================

_pph21_calculator = None

def get_pph21_calculator() -> PPh21Calculator:
    """Get singleton instance"""
    global _pph21_calculator
    if _pph21_calculator is None:
        _pph21_calculator = PPh21Calculator()
    return _pph21_calculator
