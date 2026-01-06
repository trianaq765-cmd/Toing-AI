"""
Document Templates - Email, Letter, Memo templates
"""

from .email_templates import EmailTemplates
from .letter_templates import LetterTemplates

class DocumentTemplates:
    """Collection of document templates"""
    
    email = EmailTemplates
    letter = LetterTemplates

__all__ = [
    "DocumentTemplates",
    "EmailTemplates",
    "LetterTemplates",
]
