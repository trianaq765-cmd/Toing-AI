"""
Letter Templates - Template surat resmi Indonesia
"""

from typing import Dict
from datetime import datetime

class LetterTemplates:
    """
    Surat resmi templates untuk berbagai keperluan
    """
    
    @staticmethod
    def surat_keterangan_kerja(
        employee_name: str = "[Nama Karyawan]",
        employee_nik: str = "[NIK]",
        position: str = "[Jabatan]",
        department: str = "[Departemen]",
        start_date: str = "[Tanggal Mulai Kerja]",
        company_name: str = "[Nama Perusahaan]",
        company_address: str = "[Alamat Perusahaan]",
        signer_name: str = "[Nama Penandatangan]",
        signer_title: str = "[Jabatan Penandatangan]"
    ) -> Dict[str, str]:
        """Surat Keterangan Kerja template"""
        
        today = datetime.now().strftime("%d %B %Y")
        letter_number = f"SKK/{datetime.now().strftime('%Y%m')}/001"
        
        return {
            "title": "SURAT KETERANGAN KERJA",
            "body": f"""{company_name}
{company_address}

SURAT KETERANGAN KERJA
Nomor: {letter_number}

Yang bertanda tangan di bawah ini:
Nama    : {signer_name}
Jabatan : {signer_title}

Dengan ini menerangkan bahwa:
Nama        : {employee_name}
NIK         : {employee_nik}
Jabatan     : {position}
Departemen  : {department}

Adalah benar karyawan {company_name} yang telah bekerja sejak {start_date} sampai dengan saat ini.

Surat keterangan ini dibuat untuk dipergunakan sebagaimana mestinya.

{company_address.split(',')[0] if ',' in company_address else '[Kota]'}, {today}
{company_name}


({signer_name})
{signer_title}
"""
        }
    
    @staticmethod
    def surat_permohonan_kerjasama(
        recipient_name: str = "[Nama Penerima]",
        recipient_title: str = "[Jabatan]",
        recipient_company: str = "[Nama Perusahaan]",
        recipient_address: str = "[Alamat]",
        cooperation_type: str = "[Jenis Kerjasama]",
        sender_company: str = "[Nama Perusahaan Pengirim]",
        sender_name: str = "[Nama Pengirim]",
        sender_title: str = "[Jabatan]"
    ) -> Dict[str, str]:
        """Surat Permohonan Kerjasama template"""
        
        today = datetime.now().strftime("%d %B %Y")
        letter_number = f"PKS/{datetime.now().strftime('%Y%m')}/001"
        
        return {
            "title": "SURAT PERMOHONAN KERJASAMA",
            "body": f"""Nomor    : {letter_number}
Lampiran : -
Perihal  : Permohonan Kerjasama {cooperation_type}

Kepada Yth.
{recipient_name}
{recipient_title}
{recipient_company}
{recipient_address}

Dengan hormat,

Kami dari {sender_company} bermaksud mengajukan permohonan kerjasama dengan {recipient_company} dalam bidang {cooperation_type}.

{sender_company} adalah perusahaan yang bergerak di bidang [bidang usaha]. Sejak berdiri pada [tahun berdiri], kami telah [pengalaman/track record].

Adapun bentuk kerjasama yang kami usulkan adalah:
1. [Bentuk kerjasama 1]
2. [Bentuk kerjasama 2]
3. [Bentuk kerjasama 3]

Kami yakin bahwa kerjasama ini akan memberikan manfaat yang saling menguntungkan bagi kedua belah pihak.

Sebagai bahan pertimbangan, kami lampirkan:
1. Company Profile
2. [Dokumen pendukung lainnya]

Kami sangat mengharapkan kesempatan untuk dapat mempresentasikan proposal kerjasama ini lebih lanjut.

Atas perhatian dan kerjasamanya, kami ucapkan terima kasih.

Hormat kami,
{sender_company}


({sender_name})
{sender_title}
"""
        }
    
    @staticmethod
    def surat_pengunduran_diri(
        employee_name: str = "[Nama Karyawan]",
        employee_nik: str = "[NIK]",
        position: str = "[Jabatan]",
        department: str = "[Departemen]",
        last_day: str = "[Tanggal Terakhir Bekerja]",
        company_name: str = "[Nama Perusahaan]",
        hr_name: str = "[Nama HRD]"
    ) -> Dict[str, str]:
        """Surat Pengunduran Diri template"""
        
        today = datetime.now().strftime("%d %B %Y")
        
        return {
            "title": "SURAT PENGUNDURAN DIRI",
            "body": f"""[Kota], {today}

Kepada Yth.
HRD Manager / {hr_name}
{company_name}
di Tempat

Dengan hormat,

Yang bertanda tangan di bawah ini:
Nama        : {employee_name}
NIK         : {employee_nik}
Jabatan     : {position}
Departemen  : {department}

Dengan ini mengajukan pengunduran diri dari {company_name} terhitung mulai tanggal {last_day}.

Keputusan ini saya ambil setelah melalui pertimbangan yang matang demi kepentingan karir dan pengembangan diri saya.

Saya mengucapkan terima kasih sebesar-besarnya atas kesempatan yang telah diberikan selama bekerja di {company_name}. Pengalaman dan ilmu yang saya dapatkan sangat berharga bagi perkembangan karir saya.

Saya berkomitmen untuk menyelesaikan seluruh tanggung jawab dan melakukan proses handover dengan baik sampai dengan tanggal terakhir saya bekerja.

Demikian surat pengunduran diri ini saya sampaikan. Atas perhatian dan pengertiannya, saya ucapkan terima kasih.

Hormat saya,


({employee_name})
"""
        }
    
    @staticmethod
    def memo_internal(
        from_name: str = "[Nama Pengirim]",
        from_dept: str = "[Departemen]",
        to_name: str = "[Nama Penerima]",
        to_dept: str = "[Departemen]",
        subject: str = "[Perihal]",
        content: str = "[Isi Memo]"
    ) -> Dict[str, str]:
        """Memo Internal template"""
        
        today = datetime.now().strftime("%d %B %Y")
        memo_number = f"MEMO/{datetime.now().strftime('%Y%m%d')}/001"
        
        return {
            "title": "MEMO INTERNAL",
            "body": f"""MEMO INTERNAL
Nomor: {memo_number}

Kepada  : {to_name} - {to_dept}
Dari    : {from_name} - {from_dept}
Tanggal : {today}
Perihal : {subject}

═══════════════════════════════════════════════════════════════

{content}

═══════════════════════════════════════════════════════════════

Demikian memo ini disampaikan untuk dapat ditindaklanjuti.


{from_name}
{from_dept}
"""
        }
    
    @staticmethod
    def get_all_templates() -> Dict[str, callable]:
        """Get all letter template methods"""
        return {
            "surat_keterangan_kerja": LetterTemplates.surat_keterangan_kerja,
            "surat_permohonan_kerjasama": LetterTemplates.surat_permohonan_kerjasama,
            "surat_pengunduran_diri": LetterTemplates.surat_pengunduran_diri,
            "memo_internal": LetterTemplates.memo_internal,
        }
