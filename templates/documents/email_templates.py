"""
Email Templates - Template email bisnis Indonesia
"""

from typing import Dict
from datetime import datetime

class EmailTemplates:
    """
    Email templates untuk berbagai keperluan bisnis
    """
    
    @staticmethod
    def meeting_request(
        recipient: str = "[Nama Penerima]",
        subject: str = "[Topik Meeting]",
        proposed_date: str = "[Tanggal]",
        proposed_time: str = "[Waktu]",
        location: str = "[Lokasi/Link]",
        sender_name: str = "[Nama Anda]",
        sender_title: str = "[Jabatan]"
    ) -> Dict[str, str]:
        """Meeting request email template"""
        
        return {
            "subject": f"Undangan Meeting: {subject}",
            "body": f"""Yth. Bapak/Ibu {recipient},

Dengan hormat,

Saya bermaksud mengundang Bapak/Ibu untuk menghadiri meeting terkait {subject}.

Detail Meeting:
- Hari/Tanggal : {proposed_date}
- Waktu        : {proposed_time}
- Tempat       : {location}

Agenda:
1. [Agenda 1]
2. [Agenda 2]
3. [Agenda 3]

Mohon konfirmasi kehadiran Bapak/Ibu paling lambat [tanggal konfirmasi].

Atas perhatian dan kehadirannya, kami ucapkan terima kasih.

Hormat saya,

{sender_name}
{sender_title}
"""
        }
    
    @staticmethod
    def follow_up(
        recipient: str = "[Nama Penerima]",
        reference: str = "[Referensi Meeting/Email Sebelumnya]",
        sender_name: str = "[Nama Anda]",
        sender_title: str = "[Jabatan]"
    ) -> Dict[str, str]:
        """Follow up email template"""
        
        return {
            "subject": f"Follow Up: {reference}",
            "body": f"""Yth. Bapak/Ibu {recipient},

Dengan hormat,

Menindaklanjuti {reference} pada [tanggal], saya ingin menanyakan perkembangan terkait [topik yang dibahas].

Berikut poin-poin yang perlu ditindaklanjuti:
1. [Poin 1]
2. [Poin 2]
3. [Poin 3]

Mohon informasi dan arahannya untuk langkah selanjutnya.

Terima kasih atas perhatian dan kerjasamanya.

Hormat saya,

{sender_name}
{sender_title}
"""
        }
    
    @staticmethod
    def quotation_submission(
        recipient: str = "[Nama Penerima]",
        company: str = "[Nama Perusahaan Client]",
        project: str = "[Nama Project/Produk]",
        sender_name: str = "[Nama Anda]",
        sender_title: str = "[Jabatan]"
    ) -> Dict[str, str]:
        """Quotation submission email template"""
        
        return {
            "subject": f"Penawaran Harga: {project}",
            "body": f"""Yth. Bapak/Ibu {recipient}
{company}

Dengan hormat,

Berdasarkan permintaan penawaran yang kami terima, bersama ini kami sampaikan Quotation untuk {project}.

Terlampir dokumen penawaran harga yang mencakup:
1. Rincian produk/jasa
2. Harga dan terms of payment
3. Waktu pengerjaan/pengiriman
4. Syarat dan ketentuan

Penawaran ini berlaku selama 30 (tiga puluh) hari kalender sejak tanggal surat ini.

Apabila ada pertanyaan atau memerlukan penjelasan lebih lanjut, silakan menghubungi kami.

Besar harapan kami untuk dapat bekerjasama dengan {company}.

Terima kasih atas kesempatan yang diberikan.

Hormat kami,

{sender_name}
{sender_title}
[Nama Perusahaan]
[Nomor Telepon]
[Email]
"""
        }
    
    @staticmethod
    def thank_you(
        recipient: str = "[Nama Penerima]",
        occasion: str = "[Kesempatan]",
        sender_name: str = "[Nama Anda]"
    ) -> Dict[str, str]:
        """Thank you email template"""
        
        return {
            "subject": f"Terima Kasih atas {occasion}",
            "body": f"""Yth. Bapak/Ibu {recipient},

Dengan hormat,

Kami ingin menyampaikan terima kasih yang sebesar-besarnya atas {occasion}.

[Penjelasan spesifik tentang hal yang diapresiasi]

Kami berharap dapat terus menjalin kerjasama yang baik di masa mendatang.

Sekali lagi, terima kasih atas kepercayaan yang diberikan.

Salam hangat,

{sender_name}
"""
        }
    
    @staticmethod
    def apology(
        recipient: str = "[Nama Penerima]",
        issue: str = "[Masalah]",
        solution: str = "[Solusi]",
        sender_name: str = "[Nama Anda]",
        sender_title: str = "[Jabatan]"
    ) -> Dict[str, str]:
        """Apology email template"""
        
        return {
            "subject": f"Permohonan Maaf: {issue}",
            "body": f"""Yth. Bapak/Ibu {recipient},

Dengan hormat,

Kami ingin menyampaikan permohonan maaf yang sebesar-besarnya atas {issue} yang terjadi.

Kami memahami bahwa hal ini telah menyebabkan ketidaknyamanan, dan kami sangat menyesalinya.

Sebagai bentuk tanggung jawab, kami telah mengambil langkah-langkah berikut:
{solution}

Kami berkomitmen untuk memastikan hal serupa tidak terulang kembali.

Besar harapan kami agar Bapak/Ibu dapat memaklumi kondisi ini dan tetap mempercayakan kerjasama kepada kami.

Terima kasih atas pengertian dan kesabarannya.

Hormat kami,

{sender_name}
{sender_title}
"""
        }
    
    @staticmethod
    def get_all_templates() -> Dict[str, callable]:
        """Get all email template methods"""
        return {
            "meeting_request": EmailTemplates.meeting_request,
            "follow_up": EmailTemplates.follow_up,
            "quotation_submission": EmailTemplates.quotation_submission,
            "thank_you": EmailTemplates.thank_you,
            "apology": EmailTemplates.apology,
        }
