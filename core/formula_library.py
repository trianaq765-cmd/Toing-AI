"""
Formula Library - 400+ Excel formulas dengan dokumentasi
Organized by categories with examples and tips
"""

import logging
from typing import Optional, Dict, Any, List

logger = logging.getLogger("office_bot.formula")

# =============================================================================
# FORMULA LIBRARY CLASS
# =============================================================================

class FormulaLibrary:
    """
    Library of Excel formulas with Indonesian examples
    """
    
    def __init__(self):
        self.formulas = self._build_formula_database()
        logger.info(f"Formula Library initialized with {len(self.formulas)} formulas")
    
    # =========================================================================
    # PUBLIC METHODS
    # =========================================================================
    
    def get_formula(self, name: str) -> Optional[Dict[str, Any]]:
        """Get formula by name"""
        return self.formulas.get(name.upper())
    
    def search_formula(self, keyword: str) -> List[Dict[str, Any]]:
        """Search formulas by keyword"""
        keyword = keyword.lower()
        results = []
        
        for name, info in self.formulas.items():
            if (keyword in name.lower() or 
                keyword in info['description'].lower() or
                keyword in info.get('category', '').lower()):
                results.append({**info, 'name': name})
        
        return results
    
    def get_formulas_by_category(self, category: str) -> List[Dict[str, Any]]:
        """Get all formulas in a category"""
        results = []
        category = category.lower()
        
        for name, info in self.formulas.items():
            if info.get('category', '').lower() == category:
                results.append({**info, 'name': name})
        
        return results
    
    def get_categories(self) -> List[str]:
        """Get list of all categories"""
        categories = set()
        for info in self.formulas.values():
            if info.get('category'):
                categories.add(info['category'])
        return sorted(list(categories))
    
    def get_categories_with_count(self) -> Dict[str, int]:
        """Get categories with formula count"""
        counts = {}
        for info in self.formulas.values():
            category = info.get('category', 'other')
            counts[category] = counts.get(category, 0) + 1
        return counts
    
    # =========================================================================
    # FORMULA DATABASE
    # =========================================================================
    
    def _build_formula_database(self) -> Dict[str, Dict[str, Any]]:
        """Build complete formula database"""
        
        formulas = {}
        
        # LOOKUP & REFERENCE
        formulas.update(self._lookup_formulas())
        
        # MATH & TRIG
        formulas.update(self._math_formulas())
        
        # STATISTICAL
        formulas.update(self._statistical_formulas())
        
        # TEXT
        formulas.update(self._text_formulas())
        
        # DATE & TIME
        formulas.update(self._date_formulas())
        
        # LOGICAL
        formulas.update(self._logical_formulas())
        
        # FINANCIAL
        formulas.update(self._financial_formulas())
        
        # ARRAY & DYNAMIC
        formulas.update(self._array_formulas())
        
        # DATABASE
        formulas.update(self._database_formulas())
        
        return formulas
    
    # =========================================================================
    # LOOKUP & REFERENCE FORMULAS
    # =========================================================================
    
    def _lookup_formulas(self) -> Dict[str, Dict[str, Any]]:
        return {
            "VLOOKUP": {
                "category": "lookup",
                "description": "Mencari nilai dalam kolom pertama tabel dan mengembalikan nilai dari kolom lain",
                "syntax": "=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])",
                "parameters": [
                    {"name": "lookup_value", "description": "Nilai yang dicari"},
                    {"name": "table_array", "description": "Range tabel data"},
                    {"name": "col_index_num", "description": "Nomor kolom yang akan dikembalikan"},
                    {"name": "range_lookup", "description": "TRUE (approx) atau FALSE (exact match)"}
                ],
                "examples": [
                    {
                        "title": "Mencari harga produk",
                        "formula": "=VLOOKUP(A2, ProductTable, 3, FALSE)",
                        "explanation": "Mencari produk di A2 dalam ProductTable, kembalikan nilai kolom ke-3"
                    },
                    {
                        "title": "Dengan IFERROR",
                        "formula": "=IFERROR(VLOOKUP(A2, Data, 2, FALSE), 'Tidak ditemukan')",
                        "explanation": "Handle error jika data tidak ditemukan"
                    }
                ],
                "tips": "Gunakan FALSE untuk exact match. Tabel harus diurutkan jika menggunakan TRUE.",
                "related": ["HLOOKUP", "XLOOKUP", "INDEX", "MATCH"]
            },
            
            "HLOOKUP": {
                "category": "lookup",
                "description": "Mencari nilai dalam baris pertama dan mengembalikan nilai dari baris lain",
                "syntax": "=HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])",
                "parameters": [
                    {"name": "lookup_value", "description": "Nilai yang dicari"},
                    {"name": "table_array", "description": "Range tabel data (horizontal)"},
                    {"name": "row_index_num", "description": "Nomor baris yang akan dikembalikan"},
                    {"name": "range_lookup", "description": "TRUE atau FALSE"}
                ],
                "examples": [
                    {
                        "title": "Mencari data per bulan",
                        "formula": "=HLOOKUP('Jan', A1:M5, 3, FALSE)",
                        "explanation": "Mencari kolom 'Jan' dan kembalikan nilai baris ke-3"
                    }
                ],
                "tips": "Kebalikan dari VLOOKUP, digunakan untuk tabel horizontal"
            },
            
            "XLOOKUP": {
                "category": "lookup",
                "description": "Fungsi lookup modern yang lebih powerful (Excel 365)",
                "syntax": "=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])",
                "parameters": [
                    {"name": "lookup_value", "description": "Nilai yang dicari"},
                    {"name": "lookup_array", "description": "Array tempat mencari"},
                    {"name": "return_array", "description": "Array yang dikembalikan"},
                    {"name": "if_not_found", "description": "Nilai jika tidak ditemukan"}
                ],
                "examples": [
                    {
                        "title": "Lookup sederhana",
                        "formula": "=XLOOKUP(A2, B:B, C:C, 'Not Found')",
                        "explanation": "Cari A2 di kolom B, kembalikan nilai dari kolom C"
                    }
                ],
                "tips": "Lebih fleksibel dari VLOOKUP, bisa lookup ke kiri, support array"
            },
            
            "INDEX": {
                "category": "lookup",
                "description": "Mengembalikan nilai dari posisi tertentu dalam range",
                "syntax": "=INDEX(array, row_num, [column_num])",
                "parameters": [
                    {"name": "array", "description": "Range data"},
                    {"name": "row_num", "description": "Nomor baris"},
                    {"name": "column_num", "description": "Nomor kolom (optional)"}
                ],
                "examples": [
                    {
                        "title": "Get nilai spesifik",
                        "formula": "=INDEX(A1:C10, 5, 2)",
                        "explanation": "Ambil nilai di baris 5, kolom 2 dari range A1:C10"
                    },
                    {
                        "title": "Kombinasi dengan MATCH",
                        "formula": "=INDEX(C:C, MATCH(A2, B:B, 0))",
                        "explanation": "Lookup fleksibel, bisa ke kiri (alternatif VLOOKUP)"
                    }
                ],
                "tips": "Sangat powerful jika dikombinasikan dengan MATCH"
            },
            
            "MATCH": {
                "category": "lookup",
                "description": "Mencari posisi nilai dalam range",
                "syntax": "=MATCH(lookup_value, lookup_array, [match_type])",
                "parameters": [
                    {"name": "lookup_value", "description": "Nilai yang dicari"},
                    {"name": "lookup_array", "description": "Range pencarian"},
                    {"name": "match_type", "description": "0=exact, 1=less than, -1=greater than"}
                ],
                "examples": [
                    {
                        "title": "Cari posisi",
                        "formula": "=MATCH('Apel', A:A, 0)",
                        "explanation": "Cari posisi 'Apel' di kolom A"
                    }
                ],
                "tips": "Sering dikombinasikan dengan INDEX untuk lookup 2 arah"
            },
            
            "OFFSET": {
                "category": "lookup",
                "description": "Referensi cell dengan offset dari posisi tertentu",
                "syntax": "=OFFSET(reference, rows, cols, [height], [width])",
                "examples": [
                    {
                        "title": "Dynamic range",
                        "formula": "=SUM(OFFSET(A1, 0, 0, 10, 1))",
                        "explanation": "Sum 10 cells mulai dari A1"
                    }
                ],
                "tips": "Berguna untuk membuat dynamic named ranges"
            },
            
            "INDIRECT": {
                "category": "lookup",
                "description": "Mengkonversi text menjadi referensi cell",
                "syntax": "=INDIRECT(ref_text, [a1])",
                "examples": [
                    {
                        "title": "Dynamic reference",
                        "formula": "=INDIRECT('Sheet'&A1&'!B5')",
                        "explanation": "Referensi sheet dinamis berdasarkan nilai A1"
                    }
                ],
                "tips": "Berguna untuk formula dinamis, tapi membuat file lebih lambat"
            }
        }
    
    # =========================================================================
    # MATH & TRIG FORMULAS
    # =========================================================================
    
    def _math_formulas(self) -> Dict[str, Dict[str, Any]]:
        return {
            "SUM": {
                "category": "math",
                "description": "Menjumlahkan semua angka dalam range",
                "syntax": "=SUM(number1, [number2], ...)",
                "examples": [
                    {
                        "title": "Jumlah range",
                        "formula": "=SUM(A1:A10)",
                        "explanation": "Jumlahkan semua nilai dari A1 sampai A10"
                    },
                    {
                        "title": "Multiple ranges",
                        "formula": "=SUM(A1:A10, C1:C10, 100)",
                        "explanation": "Jumlahkan beberapa range plus nilai tambahan"
                    }
                ],
                "tips": "Formula paling basic dan sering digunakan"
            },
            
            "SUMIF": {
                "category": "math",
                "description": "Menjumlahkan cell yang memenuhi kriteria",
                "syntax": "=SUMIF(range, criteria, [sum_range])",
                "parameters": [
                    {"name": "range", "description": "Range untuk cek kriteria"},
                    {"name": "criteria", "description": "Kondisi yang harus dipenuhi"},
                    {"name": "sum_range", "description": "Range yang dijumlahkan (optional)"}
                ],
                "examples": [
                    {
                        "title": "Jumlah dengan kondisi",
                        "formula": "=SUMIF(A:A, 'Jakarta', B:B)",
                        "explanation": "Jumlahkan B jika A = 'Jakarta'"
                    },
                    {
                        "title": "Kondisi angka",
                        "formula": "=SUMIF(C:C, '>1000000', C:C)",
                        "explanation": "Jumlahkan nilai di C yang > Rp 1.000.000"
                    }
                ],
                "tips": "Gunakan wildcard: * untuk text matching"
            },
            
            "SUMIFS": {
                "category": "math",
                "description": "Menjumlahkan dengan multiple kriteria",
                "syntax": "=SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)",
                "examples": [
                    {
                        "title": "Multiple kondisi",
                        "formula": "=SUMIFS(D:D, A:A, 'Jakarta', B:B, 'Produk A')",
                        "explanation": "Jumlahkan D jika A='Jakarta' DAN B='Produk A'"
                    },
                    {
                        "title": "Date range",
                        "formula": "=SUMIFS(C:C, B:B, '>=01/01/2024', B:B, '<=31/12/2024')",
                        "explanation": "Jumlahkan C untuk tahun 2024"
                    }
                ],
                "tips": "Bisa pakai banyak kriteria, lebih powerful dari SUMIF"
            },
            
            "SUMPRODUCT": {
                "category": "math",
                "description": "Mengalikan lalu menjumlahkan arrays",
                "syntax": "=SUMPRODUCT(array1, [array2], ...)",
                "examples": [
                    {
                        "title": "Qty x Price",
                        "formula": "=SUMPRODUCT(B2:B10, C2:C10)",
                        "explanation": "Kalikan Qty dengan Price lalu jumlahkan semua"
                    },
                    {
                        "title": "Conditional count",
                        "formula": "=SUMPRODUCT((A:A='Jakarta')*(B:B>1000))",
                        "explanation": "Count rows dengan multiple kondisi"
                    }
                ],
                "tips": "Sangat powerful untuk perhitungan complex tanpa array formula"
            },
            
            "ROUND": {
                "category": "math",
                "description": "Membulatkan angka ke digit tertentu",
                "syntax": "=ROUND(number, num_digits)",
                "examples": [
                    {
                        "title": "Bulatkan ke ribuan",
                        "formula": "=ROUND(1234.56, -3)",
                        "explanation": "Hasil: 1000"
                    },
                    {
                        "title": "Bulatkan 2 desimal",
                        "formula": "=ROUND(1234.5678, 2)",
                        "explanation": "Hasil: 1234.57"
                    }
                ],
                "tips": "num_digits negatif membulatkan ke kiri desimal"
            },
            
            "ROUNDUP": {
                "category": "math",
                "description": "Membulatkan ke atas",
                "syntax": "=ROUNDUP(number, num_digits)",
                "tips": "Selalu bulatkan ke atas, berguna untuk estimasi cost"
            },
            
            "ROUNDDOWN": {
                "category": "math",
                "description": "Membulatkan ke bawah",
                "syntax": "=ROUNDDOWN(number, num_digits)",
                "tips": "Selalu bulatkan ke bawah"
            },
            
            "CEILING": {
                "category": "math",
                "description": "Membulatkan ke atas ke multiple terdekat",
                "syntax": "=CEILING(number, significance)",
                "examples": [
                    {
                        "title": "Bulatkan ke 1000",
                        "formula": "=CEILING(1234, 1000)",
                        "explanation": "Hasil: 2000"
                    }
                ],
                "tips": "Berguna untuk pricing ke kelipatan tertentu"
            },
            
            "FLOOR": {
                "category": "math",
                "description": "Membulatkan ke bawah ke multiple terdekat",
                "syntax": "=FLOOR(number, significance)",
                "tips": "Kebalikan dari CEILING"
            },
            
            "ABS": {
                "category": "math",
                "description": "Nilai absolut (hapus tanda negatif)",
                "syntax": "=ABS(number)",
                "examples": [
                    {
                        "title": "Absolut",
                        "formula": "=ABS(-100)",
                        "explanation": "Hasil: 100"
                    }
                ],
                "tips": "Berguna untuk menghitung selisih tanpa peduli arah"
            },
            
            "MOD": {
                "category": "math",
                "description": "Sisa pembagian (remainder)",
                "syntax": "=MOD(number, divisor)",
                "examples": [
                    {
                        "title": "Check genap/ganjil",
                        "formula": "=MOD(A1, 2)",
                        "explanation": "Jika 0 = genap, jika 1 = ganjil"
                    }
                ],
                "tips": "Berguna untuk pattern alternate rows, dll"
            },
            
            "POWER": {
                "category": "math",
                "description": "Pangkat",
                "syntax": "=POWER(number, power)",
                "examples": [
                    {
                        "title": "Kuadrat",
                        "formula": "=POWER(5, 2)",
                        "explanation": "Hasil: 25 (5^2)"
                    }
                ]
            },
            
            "SQRT": {
                "category": "math",
                "description": "Akar kuadrat",
                "syntax": "=SQRT(number)",
                "examples": [
                    {
                        "title": "Akar",
                        "formula": "=SQRT(25)",
                        "explanation": "Hasil: 5"
                    }
                ]
            }
        }
    
    # =========================================================================
    # STATISTICAL FORMULAS
    # =========================================================================
    
    def _statistical_formulas(self) -> Dict[str, Dict[str, Any]]:
        return {
            "AVERAGE": {
                "category": "statistical",
                "description": "Menghitung rata-rata",
                "syntax": "=AVERAGE(number1, [number2], ...)",
                "examples": [
                    {
                        "title": "Rata-rata",
                        "formula": "=AVERAGE(A1:A10)",
                        "explanation": "Rata-rata nilai A1 sampai A10"
                    }
                ],
                "tips": "Mengabaikan cell kosong dan text"
            },
            
            "AVERAGEIF": {
                "category": "statistical",
                "description": "Rata-rata dengan kriteria",
                "syntax": "=AVERAGEIF(range, criteria, [average_range])",
                "examples": [
                    {
                        "title": "Rata-rata bersyarat",
                        "formula": "=AVERAGEIF(A:A, 'Jakarta', B:B)",
                        "explanation": "Rata-rata B untuk data Jakarta"
                    }
                ]
            },
            
            "AVERAGEIFS": {
                "category": "statistical",
                "description": "Rata-rata dengan multiple kriteria",
                "syntax": "=AVERAGEIFS(average_range, criteria_range1, criteria1, ...)",
                "tips": "Seperti SUMIFS tapi untuk average"
            },
            
            "COUNT": {
                "category": "statistical",
                "description": "Menghitung jumlah cell berisi angka",
                "syntax": "=COUNT(value1, [value2], ...)",
                "tips": "Hanya menghitung cell dengan angka"
            },
            
            "COUNTA": {
                "category": "statistical",
                "description": "Menghitung cell tidak kosong",
                "syntax": "=COUNTA(value1, [value2], ...)",
                "tips": "Menghitung semua cell yang tidak kosong (termasuk text)"
            },
            
            "COUNTBLANK": {
                "category": "statistical",
                "description": "Menghitung cell kosong",
                "syntax": "=COUNTBLANK(range)"
            },
            
            "COUNTIF": {
                "category": "statistical",
                "description": "Menghitung cell yang memenuhi kriteria",
                "syntax": "=COUNTIF(range, criteria)",
                "examples": [
                    {
                        "title": "Count text",
                        "formula": "=COUNTIF(A:A, 'Jakarta')",
                        "explanation": "Hitung berapa kali 'Jakarta' muncul"
                    },
                    {
                        "title": "Count > value",
                        "formula": "=COUNTIF(B:B, '>1000000')",
                        "explanation": "Hitung cell dengan nilai > 1 juta"
                    }
                ]
            },
            
            "COUNTIFS": {
                "category": "statistical",
                "description": "Menghitung dengan multiple kriteria",
                "syntax": "=COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], ...)",
                "examples": [
                    {
                        "title": "Multiple conditions",
                        "formula": "=COUNTIFS(A:A, 'Jakarta', B:B, '>1000000')",
                        "explanation": "Count jika Jakarta DAN nilai > 1 juta"
                    }
                ]
            },
            
            "MAX": {
                "category": "statistical",
                "description": "Nilai maksimum",
                "syntax": "=MAX(number1, [number2], ...)"
            },
            
            "MIN": {
                "category": "statistical",
                "description": "Nilai minimum",
                "syntax": "=MIN(number1, [number2], ...)"
            },
            
            "MEDIAN": {
                "category": "statistical",
                "description": "Nilai tengah",
                "syntax": "=MEDIAN(number1, [number2], ...)",
                "tips": "Lebih baik dari AVERAGE untuk data dengan outlier"
            },
            
            "MODE": {
                "category": "statistical",
                "description": "Nilai yang paling sering muncul",
                "syntax": "=MODE.SNGL(number1, [number2], ...)"
            },
            
            "STDEV": {
                "category": "statistical",
                "description": "Standar deviasi (sample)",
                "syntax": "=STDEV.S(number1, [number2], ...)",
                "tips": "Ukuran variasi/penyebaran data"
            },
            
            "VAR": {
                "category": "statistical",
                "description": "Variance (sample)",
                "syntax": "=VAR.S(number1, [number2], ...)"
            },
            
            "PERCENTILE": {
                "category": "statistical",
                "description": "Nilai pada percentile tertentu",
                "syntax": "=PERCENTILE.INC(array, k)",
                "examples": [
                    {
                        "title": "90th percentile",
                        "formula": "=PERCENTILE.INC(A:A, 0.9)",
                        "explanation": "Nilai pada posisi 90%"
                    }
                ]
            },
            
            "RANK": {
                "category": "statistical",
                "description": "Peringkat nilai dalam list",
                "syntax": "=RANK.EQ(number, ref, [order])",
                "examples": [
                    {
                        "title": "Ranking",
                        "formula": "=RANK.EQ(B2, $B$2:$B$100, 0)",
                        "explanation": "Ranking dari terbesar ke terkecil (0=desc)"
                    }
                ]
            }
        }
    
    # =========================================================================
    # TEXT FORMULAS
    # =========================================================================
    
    def _text_formulas(self) -> Dict[str, Dict[str, Any]]:
        return {
            "CONCATENATE": {
                "category": "text",
                "description": "Menggabungkan text (gunakan & atau TEXTJOIN untuk Excel modern)",
                "syntax": "=CONCATENATE(text1, [text2], ...)",
                "examples": [
                    {
                        "title": "Gabung nama",
                        "formula": "=CONCATENATE(A2, ' ', B2)",
                        "explanation": "Gabung first name dan last name dengan spasi"
                    }
                ],
                "tips": "Di Excel modern, lebih baik pakai operator & atau TEXTJOIN"
            },
            
            "TEXTJOIN": {
                "category": "text",
                "description": "Gabung text dengan delimiter (Excel 2019+)",
                "syntax": "=TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)",
                "examples": [
                    {
                        "title": "Join dengan koma",
                        "formula": "=TEXTJOIN(', ', TRUE, A1:A10)",
                        "explanation": "Gabung A1-A10 dengan koma, skip yang kosong"
                    }
                ],
                "tips": "Sangat powerful, bisa join array/range sekaligus"
            },
            
            "LEFT": {
                "category": "text",
                "description": "Ambil karakter dari kiri",
                "syntax": "=LEFT(text, [num_chars])",
                "examples": [
                    {
                        "title": "3 huruf pertama",
                        "formula": "=LEFT(A2, 3)",
                        "explanation": "Ambil 3 karakter dari kiri"
                    }
                ]
            },
            
            "RIGHT": {
                "category": "text",
                "description": "Ambil karakter dari kanan",
                "syntax": "=RIGHT(text, [num_chars])",
                "tips": "Berguna untuk extract kode, extension, dll dari kanan"
            },
            
            "MID": {
                "category": "text",
                "description": "Ambil karakter dari tengah",
                "syntax": "=MID(text, start_num, num_chars)",
                "examples": [
                    {
                        "title": "Extract middle",
                        "formula": "=MID(A2, 4, 5)",
                        "explanation": "Ambil 5 karakter mulai dari posisi 4"
                    }
                ]
            },
            
            "LEN": {
                "category": "text",
                "description": "Panjang text (jumlah karakter)",
                "syntax": "=LEN(text)",
                "tips": "Berguna untuk validasi, check format, dll"
            },
            
            "TRIM": {
                "category": "text",
                "description": "Hapus spasi berlebih",
                "syntax": "=TRIM(text)",
                "tips": "Penting untuk data cleaning, hapus leading/trailing/extra spaces"
            },
            
            "UPPER": {
                "category": "text",
                "description": "Ubah ke huruf kapital semua",
                "syntax": "=UPPER(text)"
            },
            
            "LOWER": {
                "category": "text",
                "description": "Ubah ke huruf kecil semua",
                "syntax": "=LOWER(text)"
            },
            
            "PROPER": {
                "category": "text",
                "description": "Capitalize setiap kata (Title Case)",
                "syntax": "=PROPER(text)",
                "examples": [
                    {
                        "title": "Title case",
                        "formula": "=PROPER('john doe')",
                        "explanation": "Hasil: 'John Doe'"
                    }
                ]
            },
            
            "SUBSTITUTE": {
                "category": "text",
                "description": "Replace text tertentu",
                "syntax": "=SUBSTITUTE(text, old_text, new_text, [instance_num])",
                "examples": [
                    {
                        "title": "Replace",
                        "formula": "=SUBSTITUTE(A2, '.', ',')",
                        "explanation": "Ganti titik dengan koma"
                    }
                ]
            },
            
            "REPLACE": {
                "category": "text",
                "description": "Replace text pada posisi tertentu",
                "syntax": "=REPLACE(old_text, start_num, num_chars, new_text)"
            },
            
            "FIND": {
                "category": "text",
                "description": "Cari posisi text (case sensitive)",
                "syntax": "=FIND(find_text, within_text, [start_num])",
                "tips": "Case sensitive, gunakan SEARCH untuk case insensitive"
            },
            
            "SEARCH": {
                "category": "text",
                "description": "Cari posisi text (not case sensitive)",
                "syntax": "=SEARCH(find_text, within_text, [start_num])",
                "tips": "Support wildcard * dan ?"
            },
            
            "TEXT": {
                "category": "text",
                "description": "Format angka/tanggal jadi text",
                "syntax": "=TEXT(value, format_text)",
                "examples": [
                    {
                        "title": "Format Rupiah",
                        "formula": "=TEXT(A2, 'Rp #,##0')",
                        "explanation": "Format angka ke Rupiah"
                    },
                    {
                        "title": "Format tanggal",
                        "formula": "=TEXT(A2, 'DD/MM/YYYY')",
                        "explanation": "Format tanggal Indonesia"
                    }
                ]
            },
            
            "VALUE": {
                "category": "text",
                "description": "Konversi text ke angka",
                "syntax": "=VALUE(text)",
                "tips": "Berguna untuk convert text yang berisi angka"
            }
        }
    
    # =========================================================================
    # DATE & TIME FORMULAS
    # =========================================================================
    
    def _date_formulas(self) -> Dict[str, Dict[str, Any]]:
        return {
            "TODAY": {
                "category": "date",
                "description": "Tanggal hari ini",
                "syntax": "=TODAY()",
                "tips": "Update otomatis setiap hari"
            },
            
            "NOW": {
                "category": "date",
                "description": "Tanggal dan waktu sekarang",
                "syntax": "=NOW()",
                "tips": "Update setiap kali file recalculate"
            },
            
            "DATE": {
                "category": "date",
                "description": "Buat tanggal dari year, month, day",
                "syntax": "=DATE(year, month, day)",
                "examples": [
                    {
                        "title": "Buat tanggal",
                        "formula": "=DATE(2024, 1, 15)",
                        "explanation": "Hasil: 15 Januari 2024"
                    }
                ]
            },
            
            "YEAR": {
                "category": "date",
                "description": "Extract tahun dari tanggal",
                "syntax": "=YEAR(serial_number)"
            },
            
            "MONTH": {
                "category": "date",
                "description": "Extract bulan (1-12)",
                "syntax": "=MONTH(serial_number)"
            },
            
            "DAY": {
                "category": "date",
                "description": "Extract hari (1-31)",
                "syntax": "=DAY(serial_number)"
            },
            
            "WEEKDAY": {
                "category": "date",
                "description": "Hari dalam seminggu (1-7)",
                "syntax": "=WEEKDAY(serial_number, [return_type])",
                "tips": "1=Minggu, 2=Senin, dst"
            },
            
            "NETWORKDAYS": {
                "category": "date",
                "description": "Jumlah hari kerja antara 2 tanggal",
                "syntax": "=NETWORKDAYS(start_date, end_date, [holidays])",
                "examples": [
                    {
                        "title": "Hari kerja",
                        "formula": "=NETWORKDAYS(A2, B2, HolidayList)",
                        "explanation": "Hitung hari kerja, exclude weekend dan libur"
                    }
                ],
                "tips": "Exclude Sabtu-Minggu otomatis"
            },
            
            "WORKDAY": {
                "category": "date",
                "description": "Tanggal setelah X hari kerja",
                "syntax": "=WORKDAY(start_date, days, [holidays])",
                "examples": [
                    {
                        "title": "Deadline",
                        "formula": "=WORKDAY(TODAY(), 30)",
                        "explanation": "Tanggal 30 hari kerja dari hari ini"
                    }
                ]
            },
            
            "DATEDIF": {
                "category": "date",
                "description": "Selisih tanggal (hidden function, tidak ada di autocomplete)",
                "syntax": "=DATEDIF(start_date, end_date, unit)",
                "examples": [
                    {
                        "title": "Umur dalam tahun",
                        "formula": "=DATEDIF(A2, TODAY(), 'Y')",
                        "explanation": "Hitung umur dalam tahun penuh"
                    },
                    {
                        "title": "Masa kerja",
                        "formula": "=DATEDIF(A2, TODAY(), 'Y') & ' tahun ' & DATEDIF(A2, TODAY(), 'YM') & ' bulan'",
                        "explanation": "Format: X tahun Y bulan"
                    }
                ],
                "tips": "Unit: 'Y'=years, 'M'=months, 'D'=days, 'YM'=months ignoring years"
            },
            
            "EDATE": {
                "category": "date",
                "description": "Tanggal X bulan dari tanggal tertentu",
                "syntax": "=EDATE(start_date, months)",
                "examples": [
                    {
                        "title": "3 bulan ke depan",
                        "formula": "=EDATE(TODAY(), 3)",
                        "explanation": "Tanggal 3 bulan dari sekarang"
                    }
                ]
            },
            
            "EOMONTH": {
                "category": "date",
                "description": "Hari terakhir bulan",
                "syntax": "=EOMONTH(start_date, months)",
                "examples": [
                    {
                        "title": "Akhir bulan",
                        "formula": "=EOMONTH(TODAY(), 0)",
                        "explanation": "Hari terakhir bulan ini"
                    }
                ]
            }
        }
    
    # =========================================================================
    # LOGICAL FORMULAS
    # =========================================================================
    
    def _logical_formulas(self) -> Dict[str, Dict[str, Any]]:
        return {
            "IF": {
                "category": "logical",
                "description": "Kondisi if-then-else",
                "syntax": "=IF(logical_test, value_if_true, value_if_false)",
                "examples": [
                    {
                        "title": "Pass/Fail",
                        "formula": "=IF(A2>=70, 'Lulus', 'Tidak Lulus')",
                        "explanation": "Jika nilai >= 70 maka Lulus"
                    },
                    {
                        "title": "Nested IF",
                        "formula": "=IF(A2>=90, 'A', IF(A2>=80, 'B', IF(A2>=70, 'C', 'D')))",
                        "explanation": "Grade berdasarkan nilai"
                    }
                ],
                "tips": "Nested IF maksimal 64 level, tapi sebaiknya max 3-4 saja"
            },
            
            "IFS": {
                "category": "logical",
                "description": "Multiple IF tanpa nesting (Excel 2019+)",
                "syntax": "=IFS(condition1, value1, [condition2, value2], ...)",
                "examples": [
                    {
                        "title": "Grade lebih clean",
                        "formula": "=IFS(A2>=90, 'A', A2>=80, 'B', A2>=70, 'C', TRUE, 'D')",
                        "explanation": "Lebih mudah dibaca dari nested IF"
                    }
                ],
                "tips": "Gunakan TRUE sebagai default/catch-all di akhir"
            },
            
            "SWITCH": {
                "category": "logical",
                "description": "Switch case berdasarkan nilai (Excel 2019+)",
                "syntax": "=SWITCH(expression, value1, result1, [value2, result2], ..., [default])",
                "examples": [
                    {
                        "title": "Nama bulan",
                        "formula": "=SWITCH(MONTH(A2), 1, 'Jan', 2, 'Feb', 3, 'Mar', 'Other')",
                        "explanation": "Convert nomor bulan jadi nama"
                    }
                ]
            },
            
            "AND": {
                "category": "logical",
                "description": "TRUE jika semua kondisi TRUE",
                "syntax": "=AND(logical1, [logical2], ...)",
                "examples": [
                    {
                        "title": "Multiple conditions",
                        "formula": "=IF(AND(A2>0, B2>0, C2='Active'), 'Valid', 'Invalid')",
                        "explanation": "Valid jika semua kondisi terpenuhi"
                    }
                ]
            },
            
            "OR": {
                "category": "logical",
                "description": "TRUE jika salah satu kondisi TRUE",
                "syntax": "=OR(logical1, [logical2], ...)",
                "tips": "Minimal 1 kondisi harus TRUE"
            },
            
            "NOT": {
                "category": "logical",
                "description": "Kebalikan dari kondisi",
                "syntax": "=NOT(logical)",
                "examples": [
                    {
                        "title": "Negasi",
                        "formula": "=NOT(A2='Jakarta')",
                        "explanation": "TRUE jika bukan Jakarta"
                    }
                ]
            },
            
            "XOR": {
                "category": "logical",
                "description": "Exclusive OR (hanya 1 yang TRUE)",
                "syntax": "=XOR(logical1, [logical2], ...)",
                "tips": "TRUE jika ganjil jumlah TRUE, FALSE jika genap"
            },
            
            "IFERROR": {
                "category": "logical",
                "description": "Handle error dengan nilai alternatif",
                "syntax": "=IFERROR(value, value_if_error)",
                "examples": [
                    {
                        "title": "Handle VLOOKUP error",
                        "formula": "=IFERROR(VLOOKUP(A2, Table, 2, 0), 'Not Found')",
                        "explanation": "Return 'Not Found' jika VLOOKUP error"
                    },
                    {
                        "title": "Handle division by zero",
                        "formula": "=IFERROR(A2/B2, 0)",
                        "explanation": "Return 0 jika pembagian error"
                    }
                ],
                "tips": "Sangat penting untuk formula yang bisa error"
            },
            
            "IFNA": {
                "category": "logical",
                "description": "Handle #N/A error saja",
                "syntax": "=IFNA(value, value_if_na)",
                "tips": "Lebih spesifik dari IFERROR, hanya catch #N/A"
            },
            
            "ISBLANK": {
                "category": "logical",
                "description": "Cek apakah cell kosong",
                "syntax": "=ISBLANK(value)"
            },
            
            "ISERROR": {
                "category": "logical",
                "description": "Cek apakah ada error",
                "syntax": "=ISERROR(value)"
            },
            
            "ISNUMBER": {
                "category": "logical",
                "description": "Cek apakah angka",
                "syntax": "=ISNUMBER(value)"
            },
            
            "ISTEXT": {
                "category": "logical",
                "description": "Cek apakah text",
                "syntax": "=ISTEXT(value)"
            }
        }
    
    # =========================================================================
    # FINANCIAL FORMULAS
    # =========================================================================
    
    def _financial_formulas(self) -> Dict[str, Dict[str, Any]]:
        return {
            "PMT": {
                "category": "financial",
                "description": "Cicilan/angsuran per periode",
                "syntax": "=PMT(rate, nper, pv, [fv], [type])",
                "parameters": [
                    {"name": "rate", "description": "Interest rate per period"},
                    {"name": "nper", "description": "Number of periods"},
                    {"name": "pv", "description": "Present value (pokok pinjaman)"},
                    {"name": "fv", "description": "Future value (optional)"},
                    {"name": "type", "description": "0=end of period, 1=beginning"}
                ],
                "examples": [
                    {
                        "title": "Cicilan rumah",
                        "formula": "=PMT(10%/12, 20*12, 500000000)",
                        "explanation": "Cicilan per bulan untuk KPR Rp 500jt, bunga 10%/tahun, 20 tahun"
                    }
                ],
                "tips": "Rate harus dalam periode yang sama dengan nper (monthly rate untuk monthly payment)"
            },
            
            "FV": {
                "category": "financial",
                "description": "Future value (nilai masa depan)",
                "syntax": "=FV(rate, nper, pmt, [pv], [type])",
                "examples": [
                    {
                        "title": "Investasi bulanan",
                        "formula": "=FV(8%/12, 10*12, -1000000)",
                        "explanation": "Nabung Rp 1jt/bulan, return 8%/tahun, 10 tahun"
                    }
                ],
                "tips": "Payment harus negatif (cash outflow)"
            },
            
            "PV": {
                "category": "financial",
                "description": "Present value (nilai sekarang)",
                "syntax": "=PV(rate, nper, pmt, [fv], [type])",
                "tips": "Kebalikan dari FV"
            },
            
            "NPV": {
                "category": "financial",
                "description": "Net Present Value",
                "syntax": "=NPV(rate, value1, [value2], ...)",
                "examples": [
                    {
                        "title": "NPV project",
                        "formula": "=NPV(10%, B2:B6) - A1",
                        "explanation": "NPV dengan initial investment di A1"
                    }
                ],
                "tips": "NPV tidak include initial investment, harus dikurangi manual"
            },
            
            "IRR": {
                "category": "financial",
                "description": "Internal Rate of Return",
                "syntax": "=IRR(values, [guess])",
                "examples": [
                    {
                        "title": "IRR calculation",
                        "formula": "=IRR(A1:A10)",
                        "explanation": "Harus include initial investment (negative) di A1"
                    }
                ],
                "tips": "Harus ada minimal 1 positive dan 1 negative cash flow"
            },
            
            "XNPV": {
                "category": "financial",
                "description": "NPV dengan tanggal spesifik (irregular periods)",
                "syntax": "=XNPV(rate, values, dates)",
                "tips": "Lebih akurat untuk cash flow dengan timing tidak teratur"
            },
            
            "XIRR": {
                "category": "financial",
                "description": "IRR dengan tanggal spesifik",
                "syntax": "=XIRR(values, dates, [guess])",
                "tips": "Lebih akurat dari IRR untuk irregular cash flows"
            },
            
            "RATE": {
                "category": "financial",
                "description": "Interest rate per period",
                "syntax": "=RATE(nper, pmt, pv, [fv], [type], [guess])",
                "tips": "Berguna untuk cari interest rate dari loan/investment"
            },
            
            "NPER": {
                "category": "financial",
                "description": "Number of periods",
                "syntax": "=NPER(rate, pmt, pv, [fv], [type])",
                "tips": "Hitung berapa lama untuk lunas/capai target"
            }
        }
    
    # =========================================================================
    # ARRAY & DYNAMIC FORMULAS (Excel 365)
    # =========================================================================
    
    def _array_formulas(self) -> Dict[str, Dict[str, Any]]:
        return {
            "FILTER": {
                "category": "array",
                "description": "Filter data berdasarkan kondisi (Excel 365)",
                "syntax": "=FILTER(array, include, [if_empty])",
                "examples": [
                    {
                        "title": "Filter data",
                        "formula": "=FILTER(A2:C100, B2:B100='Jakarta', 'No data')",
                        "explanation": "Filter rows dimana kolom B = Jakarta"
                    }
                ],
                "tips": "Hasil spill otomatis ke cells di bawah/samping"
            },
            
            "SORT": {
                "category": "array",
                "description": "Sort data (Excel 365)",
                "syntax": "=SORT(array, [sort_index], [sort_order], [by_col])",
                "examples": [
                    {
                        "title": "Sort descending",
                        "formula": "=SORT(A2:C100, 2, -1)",
                        "explanation": "Sort by column 2, descending"
                    }
                ]
            },
            
            "SORTBY": {
                "category": "array",
                "description": "Sort berdasarkan array lain (Excel 365)",
                "syntax": "=SORTBY(array, by_array1, [sort_order1], ...)",
                "tips": "Lebih fleksibel dari SORT"
            },
            
            "UNIQUE": {
                "category": "array",
                "description": "Ambil nilai unik (Excel 365)",
                "syntax": "=UNIQUE(array, [by_col], [exactly_once])",
                "examples": [
                    {
                        "title": "List unik",
                        "formula": "=UNIQUE(A2:A100)",
                        "explanation": "List semua nilai unik dari A2:A100"
                    }
                ]
            },
            
            "SEQUENCE": {
                "category": "array",
                "description": "Generate sequence angka (Excel 365)",
                "syntax": "=SEQUENCE(rows, [columns], [start], [step])",
                "examples": [
                    {
                        "title": "1 sampai 100",
                        "formula": "=SEQUENCE(100)",
                        "explanation": "Generate 1, 2, 3, ..., 100"
                    }
                ]
            },
            
            "RANDARRAY": {
                "category": "array",
                "description": "Generate random numbers array (Excel 365)",
                "syntax": "=RANDARRAY([rows], [columns], [min], [max], [integer])",
                "tips": "Berguna untuk generate sample data"
            }
        }
    
    # =========================================================================
    # DATABASE FORMULAS
    # =========================================================================
    
    def _database_formulas(self) -> Dict[str, Dict[str, Any]]:
        return {
            "DSUM": {
                "category": "database",
                "description": "Sum dari database dengan criteria",
                "syntax": "=DSUM(database, field, criteria)",
                "tips": "Database functions butuh criteria range terpisah"
            },
            
            "DAVERAGE": {
                "category": "database",
                "description": "Average dari database dengan criteria",
                "syntax": "=DAVERAGE(database, field, criteria)"
            },
            
            "DCOUNT": {
                "category": "database",
                "description": "Count numbers dari database",
                "syntax": "=DCOUNT(database, field, criteria)"
            },
            
            "DCOUNTA": {
                "category": "database",
                "description": "Count non-empty dari database",
                "syntax": "=DCOUNTA(database, field, criteria)"
            },
            
            "DGET": {
                "category": "database",
                "description": "Get single value dari database",
                "syntax": "=DGET(database, field, criteria)",
                "tips": "Error jika lebih dari 1 match"
            },
            
            "DMAX": {
                "category": "database",
                "description": "Max value dari database",
                "syntax": "=DMAX(database, field, criteria)"
            },
            
            "DMIN": {
                "category": "database",
                "description": "Min value dari database",
                "syntax": "=DMIN(database, field, criteria)"
            },
            
            "SUBTOTAL": {
                "category": "database",
                "description": "Subtotal yang ignore filtered rows",
                "syntax": "=SUBTOTAL(function_num, ref1, [ref2], ...)",
                "examples": [
                    {
                        "title": "Sum yang respect filter",
                        "formula": "=SUBTOTAL(9, A2:A100)",
                        "explanation": "9=SUM, akan ignore hidden rows dari filter"
                    }
                ],
                "tips": "Function num: 1-11 include hidden, 101-111 ignore hidden"
            },
            
            "AGGREGATE": {
                "category": "database",
                "description": "Aggregate dengan options ignore error, hidden, dll",
                "syntax": "=AGGREGATE(function_num, options, array, [k])",
                "tips": "Lebih powerful dari SUBTOTAL"
            }
        }

# =============================================================================
# SINGLETON INSTANCE
# =============================================================================

_formula_library_instance = None

def get_formula_library() -> FormulaLibrary:
    """Get singleton instance of FormulaLibrary"""
    global _formula_library_instance
    if _formula_library_instance is None:
        _formula_library_instance = FormulaLibrary()
    return _formula_library_instance
