"""
Payroll Module - Salary and benefits calculations
"""

import logging
from typing import Dict, Any, List, Optional
from dataclasses import dataclass

from config import PTKP_2024, PPH21_TARIF

logger = logging.getLogger("office_bot.hr.payroll")

@dataclass
class Employee:
    """Employee data"""
    nik: str
    name: str
    gaji_pokok: float
    tunjangan_tetap: float = 0
    tunjangan_tidak_tetap: float = 0
    ptkp_status: str = "TK/0"

@dataclass
class PayrollResult:
    """Payroll calculation result"""
    gross_salary: float
    bpjs_kesehatan: float
    bpjs_tk: float
    pph21: float
    total_deductions: float
    net_salary: float

class PayrollCalculator:
    """
    Payroll calculator with Indonesian regulations
    """
    
    def __init__(self):
        logger.info("Payroll Calculator initialized")
    
    def calculate_bpjs_kesehatan(self, gaji: float) -> Dict[str, float]:
        """
        BPJS Kesehatan calculation
        - Employee: 1%
        - Employer: 4%
        - Max base: Rp 12.000.000
        """
        max_base = 12_000_000
        base = min(gaji, max_base)
        
        return {
            "base": base,
            "employee": base * 0.01,
            "employer": base * 0.04,
            "total": base * 0.05
        }
    
    def calculate_bpjs_ketenagakerjaan(self, gaji: float) -> Dict[str, float]:
        """
        BPJS Ketenagakerjaan calculation
        
        JHT: Employee 2%, Employer 3.7%
        JP: Employee 1%, Employer 2% (for gaji > UMP)
        JKK: Employer 0.24% - 1.74%
        JKM: Employer 0.3%
        """
        jht_base = min(gaji, 9_500_000)
        jp_base = min(gaji, 10_000_000)
        
        jht_employee = jht_base * 0.02
        jht_employer = jht_base * 0.037
        
        jp_employee = jp_base * 0.01
        jp_employer = jp_base * 0.02
        
        jkk = gaji * 0.0054  # Assume middle rate
        jkm = gaji * 0.003
        
        return {
            "jht_employee": jht_employee,
            "jht_employer": jht_employer,
            "jp_employee": jp_employee,
            "jp_employer": jp_employer,
            "jkk": jkk,
            "jkm": jkm,
            "total_employee": jht_employee + jp_employee,
            "total_employer": jht_employer + jp_employer + jkk + jkm
        }
    
    def calculate_pph21_monthly(
        self, 
        gaji_bulanan: float,
        ptkp_status: str = "TK/0"
    ) -> float:
        """
        Calculate monthly PPh 21
        """
        # Annual gross
        gaji_tahunan = gaji_bulanan * 12
        
        # Deductions (biaya jabatan max 6jt/year)
        biaya_jabatan = min(gaji_tahunan * 0.05, 6_000_000)
        
        # Net annual
        netto = gaji_tahunan - biaya_jabatan
        
        # PTKP
        ptkp = PTKP_2024.get(ptkp_status.upper(), 54_000_000)
        
        # PKP
        pkp = max(0, netto - ptkp)
        
        # Progressive tax
        pph21_tahunan = self._calculate_progressive_tax(pkp)
        
        return pph21_tahunan / 12
    
    def _calculate_progressive_tax(self, pkp: float) -> float:
        """Calculate progressive tax"""
        if pkp <= 0:
            return 0
        
        tax = 0
        remaining = pkp
        prev_bracket = 0
        
        for bracket, rate in PPH21_TARIF:
            taxable = min(remaining, bracket - prev_bracket)
            if taxable > 0:
                tax += taxable * rate
                remaining -= taxable
            if remaining <= 0:
                break
            prev_bracket = bracket
        
        return tax
    
    def calculate_payroll(self, employee: Employee) -> PayrollResult:
        """
        Calculate complete payroll for employee
        """
        # Gross salary
        gross = employee.gaji_pokok + employee.tunjangan_tetap + employee.tunjangan_tidak_tetap
        
        # BPJS Kesehatan
        bpjs_kes = self.calculate_bpjs_kesehatan(employee.gaji_pokok)
        
        # BPJS TK
        bpjs_tk = self.calculate_bpjs_ketenagakerjaan(employee.gaji_pokok)
        
        # PPh 21
        pph21 = self.calculate_pph21_monthly(gross, employee.ptkp_status)
        
        # Total deductions
        total_deductions = bpjs_kes["employee"] + bpjs_tk["total_employee"] + pph21
        
        # Net salary
        net = gross - total_deductions
        
        return PayrollResult(
            gross_salary=gross,
            bpjs_kesehatan=bpjs_kes["employee"],
            bpjs_tk=bpjs_tk["total_employee"],
            pph21=pph21,
            total_deductions=total_deductions,
            net_salary=net
        )
    
    def calculate_overtime(
        self, 
        gaji_pokok: float, 
        jam_lembur: float,
        hari_libur: bool = False
    ) -> float:
        """
        Calculate overtime pay
        
        Normal day:
        - Hour 1: 1.5x hourly rate
        - Hour 2+: 2x hourly rate
        
        Holiday/weekend:
        - Hour 1-7: 2x hourly rate
        - Hour 8: 3x hourly rate
        - Hour 9+: 4x hourly rate
        """
        # Hourly rate = 1/173 x monthly salary
        hourly_rate = gaji_pokok / 173
        
        if hari_libur:
            if jam_lembur <= 7:
                return jam_lembur * hourly_rate * 2
            elif jam_lembur == 8:
                return (7 * hourly_rate * 2) + (hourly_rate * 3)
            else:
                return (7 * hourly_rate * 2) + (hourly_rate * 3) + ((jam_lembur - 8) * hourly_rate * 4)
        else:
            if jam_lembur <= 1:
                return jam_lembur * hourly_rate * 1.5
            else:
                return (hourly_rate * 1.5) + ((jam_lembur - 1) * hourly_rate * 2)
