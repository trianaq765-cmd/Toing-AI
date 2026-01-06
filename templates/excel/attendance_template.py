"""
Attendance Template - Template absensi karyawan
"""

from datetime import datetime, timedelta
from typing import Dict, Any, Optional, List
import calendar

class AttendanceTemplate:
    """
    Attendance/Absensi template untuk tracking kehadiran karyawan
    """
    
    @staticmethod
    def get_structure(
        month: Optional[int] = None,
        year: Optional[int] = None,
        employees: Optional[List[Dict]] = None
    ) -> Dict[str, Any]:
        """
        Generate attendance sheet structure
        
        Args:
            month: Bulan (1-12), default bulan ini
            year: Tahun, default tahun ini
            employees: List of employee data [{nik, nama, jabatan}]
        """
        
        now = datetime.now()
        if month is None:
            month = now.month
        if year is None:
            year = now.year
        
        # Get number of days in month
        num_days = calendar.monthrange(year, month)[1]
        
        # Month name
        month_name = calendar.month_name[month]
        
        if employees is None:
            employees = []
        
        # Build header row with dates
        header_row = ["No", "NIK", "Nama", "Jabatan"]
        for day in range(1, num_days + 1):
            header_row.append(str(day))
        header_row.extend(["Hadir", "Izin", "Sakit", "Alpha", "Cuti"])
        
        # Build employee rows
        employee_rows = []
        for i, emp in enumerate(employees, 1):
            row = [
                i,
                emp.get("nik", ""),
                emp.get("nama", ""),
                emp.get("jabatan", ""),
            ]
            # Add empty cells for each day
            for day in range(1, num_days + 1):
                row.append("")
            
            # Summary formulas
            row_num = i + 5
            first_day_col = 5  # Column E
            last_day_col = 4 + num_days
            
            # Count formulas
            hadir_formula = f'=COUNTIF(E{row_num}:{chr(64+last_day_col)}{row_num},"H")'
            izin_formula = f'=COUNTIF(E{row_num}:{chr(64+last_day_col)}{row_num},"I")'
            sakit_formula = f'=COUNTIF(E{row_num}:{chr(64+last_day_col)}{row_num},"S")'
            alpha_formula = f'=COUNTIF(E{row_num}:{chr(64+last_day_col)}{row_num},"A")'
            cuti_formula = f'=COUNTIF(E{row_num}:{chr(64+last_day_col)}{row_num},"C")'
            
            row.extend([hadir_formula, izin_formula, sakit_formula, alpha_formula, cuti_formula])
            employee_rows.append(row)
        
        # Add empty rows for template
        if not employees:
            for i in range(1, 21):  # 20 empty rows
                row = [i, "", "", ""]
                for day in range(1, num_days + 1):
                    row.append("")
                row_num = i + 5
                last_day_col = 4 + num_days
                row.extend([
                    f'=COUNTIF(E{row_num}:{chr(64+last_day_col)}{row_num},"H")',
                    f'=COUNTIF(E{row_num}:{chr(64+last_day_col)}{row_num},"I")',
                    f'=COUNTIF(E{row_num}:{chr(64+last_day_col)}{row_num},"S")',
                    f'=COUNTIF(E{row_num}:{chr(64+last_day_col)}{row_num},"A")',
                    f'=COUNTIF(E{row_num}:{chr(64+last_day_col)}{row_num},"C")'
                ])
                employee_rows.append(row)
        
        # Column widths
        col_widths = {
            "A": 5,
            "B": 15,
            "C": 25,
            "D": 20,
        }
        # Date columns
        for i, day in enumerate(range(1, num_days + 1), 5):
            col_widths[chr(64 + i)] = 4
        
        structure = {
            "sheets": [{
                "name": "Absensi",
                "data": [
                    ["DAFTAR HADIR KARYAWAN"] + [""] * (len(header_row) - 1),
                    [f"{month_name} {year}"] + [""] * (len(header_row) - 1),
                    [""] * len(header_row),
                    ["Keterangan: H=Hadir, I=Izin, S=Sakit, A=Alpha, C=Cuti"] + [""] * (len(header_row) - 1),
                    header_row,
                ] + employee_rows,
                "column_widths": col_widths,
                "formatting": {
                    "header_row": 5
                }
            }]
        }
        
        return structure
    
    @staticmethod
    def get_daily_attendance_structure(
        date: Optional[datetime] = None,
    ) -> Dict[str, Any]:
        """Generate daily attendance sheet"""
        
        if date is None:
            date = datetime.now()
        
        structure = {
            "sheets": [{
                "name": "Absensi Harian",
                "data": [
                    ["DAFTAR HADIR HARIAN", "", "", "", "", ""],
                    [f"Tanggal: {date.strftime('%d/%m/%Y')}", "", "", "", "", ""],
                    ["", "", "", "", "", ""],
                    ["No", "NIK", "Nama", "Jam Masuk", "Jam Keluar", "Keterangan"],
                ] + [
                    [i, "", "", "", "", ""] for i in range(1, 51)
                ],
                "column_widths": {
                    "A": 5,
                    "B": 15,
                    "C": 30,
                    "D": 12,
                    "E": 12,
                    "F": 25
                },
                "formatting": {
                    "header_row": 4
                }
            }]
        }
        
        return structure
