"""
Payroll Template - Template slip gaji Indonesia
Include BPJS & PPh 21
"""

from datetime import datetime
from typing import Dict, Any, Optional, List

class PayrollTemplate:
    """
    Payroll/Slip Gaji template dengan format Indonesia
    Include BPJS Kesehatan, BPJS TK, dan PPh 21
    """
    
    @staticmethod
    def get_structure(
        company_name: str = "PT. NAMA PERUSAHAAN",
        period: Optional[str] = None,
        employees: Optional[List[Dict]] = None
    ) -> Dict[str, Any]:
        """
        Generate payroll structure
        
        Args:
            company_name: Nama perusahaan
            period: Periode gaji (default: bulan ini)
            employees: List of employee data
        """
        
        if period is None:
            now = datetime.now()
            period = f"{now.strftime('%B')} {now.year}"
        
        if employees is None:
            employees = []
        
        # Build header
        headers = [
            "No",
            "NIK",
            "Nama Karyawan",
            "Jabatan",
            "Gaji Pokok",
            "Tunj. Tetap",
            "Tunj. Makan",
            "Tunj. Transport",
            "Total Pendapatan",
            "BPJS Kes (1%)",
            "BPJS TK (2%)",
            "PPh 21",
            "Total Potongan",
            "Gaji Bersih"
        ]
        
        # Build employee rows with formulas
        data_rows = []
        for i, emp in enumerate(employees, 1):
            row_num = i + 5  # Data starts at row 6
            data_rows.append([
                i,
                emp.get("nik", ""),
                emp.get("nama", ""),
                emp.get("jabatan", ""),
                emp.get("gaji_pokok", 0),
                emp.get("tunj_tetap", 0),
                emp.get("tunj_makan", 0),
                emp.get("tunj_transport", 0),
                f"=SUM(E{row_num}:H{row_num})",      # Total Pendapatan
                f"=E{row_num}*0.01",                  # BPJS Kes 1%
                f"=E{row_num}*0.02",                  # BPJS TK 2%
                emp.get("pph21", 0),                  # PPh 21 (calculated separately)
                f"=SUM(J{row_num}:L{row_num})",      # Total Potongan
                f"=I{row_num}-M{row_num}"            # Gaji Bersih
            ])
        
        # Add empty rows for template
        if not employees:
            for i in range(1, 11):  # 10 empty rows
                row_num = i + 5
                data_rows.append([
                    i,
                    "",
                    "",
                    "",
                    0,
                    0,
                    0,
                    0,
                    f"=SUM(E{row_num}:H{row_num})",
                    f"=E{row_num}*0.01",
                    f"=E{row_num}*0.02",
                    0,
                    f"=SUM(J{row_num}:L{row_num})",
                    f"=I{row_num}-M{row_num}"
                ])
        
        last_row = 5 + len(data_rows)
        
        structure = {
            "sheets": [{
                "name": "Payroll",
                "data": [
                    # Header
                    [company_name, "", "", "", "", "", "", "", "", "", "", "", "", ""],
                    [f"DAFTAR GAJI KARYAWAN - {period.upper()}", "", "", "", "", "", "", "", "", "", "", "", "", ""],
                    ["", "", "", "", "", "", "", "", "", "", "", "", "", ""],
                    ["", "", "", "", "", "", "", "", "", "", "", "", "", ""],
                    # Table Header
                    headers,
                ] + data_rows + [
                    # Summary
                    ["", "", "", "", "", "", "", "", "", "", "", "", "", ""],
                    [
                        "", "", "", "TOTAL:",
                        f"=SUM(E6:E{last_row})",
                        f"=SUM(F6:F{last_row})",
                        f"=SUM(G6:G{last_row})",
                        f"=SUM(H6:H{last_row})",
                        f"=SUM(I6:I{last_row})",
                        f"=SUM(J6:J{last_row})",
                        f"=SUM(K6:K{last_row})",
                        f"=SUM(L6:L{last_row})",
                        f"=SUM(M6:M{last_row})",
                        f"=SUM(N6:N{last_row})"
                    ],
                ],
                "column_widths": {
                    "A": 5,
                    "B": 15,
                    "C": 25,
                    "D": 20,
                    "E": 15,
                    "F": 12,
                    "G": 12,
                    "H": 12,
                    "I": 15,
                    "J": 12,
                    "K": 12,
                    "L": 12,
                    "M": 15,
                    "N": 15
                },
                "formatting": {
                    "currency_columns": ["E", "F", "G", "H", "I", "J", "K", "L", "M", "N"],
                    "header_row": 5,
                    "auto_fit": False
                }
            }]
        }
        
        return structure
    
    @staticmethod
    def get_slip_gaji_structure(
        employee_name: str = "NAMA KARYAWAN",
        employee_nik: str = "000000",
        period: str = "",
        gaji_pokok: float = 0,
        tunjangan: Dict[str, float] = None,
        potongan: Dict[str, float] = None
    ) -> Dict[str, Any]:
        """Generate individual slip gaji"""
        
        if period == "":
            now = datetime.now()
            period = f"{now.strftime('%B')} {now.year}"
        
        if tunjangan is None:
            tunjangan = {"Tunjangan Tetap": 0, "Tunjangan Makan": 0, "Tunjangan Transport": 0}
        
        if potongan is None:
            potongan = {"BPJS Kesehatan": 0, "BPJS TK": 0, "PPh 21": 0}
        
        # Build tunjangan rows
        tunjangan_rows = []
        for name, value in tunjangan.items():
            tunjangan_rows.append(["", name, value, "", "", ""])
        
        # Build potongan rows
        potongan_rows = []
        for name, value in potongan.items():
            potongan_rows.append(["", "", "", "", name, value])
        
        structure = {
            "sheets": [{
                "name": "Slip Gaji",
                "data": [
                    ["SLIP GAJI", "", "", "", "", ""],
                    ["", "", "", "", "", ""],
                    [f"Nama: {employee_name}", "", "", f"NIK: {employee_nik}", "", ""],
                    [f"Periode: {period}", "", "", "", "", ""],
                    ["", "", "", "", "", ""],
                    ["PENDAPATAN", "", "", "", "POTONGAN", ""],
                    ["", "Gaji Pokok", gaji_pokok, "", "", ""],
                ] + [
                    # Merge tunjangan and potongan rows
                    [
                        tunjangan_rows[i][0] if i < len(tunjangan_rows) else "",
                        tunjangan_rows[i][1] if i < len(tunjangan_rows) else "",
                        tunjangan_rows[i][2] if i < len(tunjangan_rows) else "",
                        "",
                        potongan_rows[i][4] if i < len(potongan_rows) else "",
                        potongan_rows[i][5] if i < len(potongan_rows) else "",
                    ]
                    for i in range(max(len(tunjangan_rows), len(potongan_rows)))
                ] + [
                    ["", "", "", "", "", ""],
                    ["", "Total Pendapatan", f"=SUM(C7:C{7+len(tunjangan_rows)})", "", "Total Potongan", f"=SUM(F7:F{7+len(potongan_rows)})"],
                    ["", "", "", "", "", ""],
                    ["", "", "GAJI BERSIH", f"=C{8+max(len(tunjangan_rows), len(potongan_rows))}-F{8+max(len(tunjangan_rows), len(potongan_rows))}", "", ""],
                ],
                "column_widths": {
                    "A": 3,
                    "B": 25,
                    "C": 18,
                    "D": 5,
                    "E": 25,
                    "F": 18
                },
                "formatting": {
                    "currency_columns": ["C", "F"]
                }
            }]
        }
        
        return structure
