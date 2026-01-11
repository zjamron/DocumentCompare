"""
DocumentCompare - Document comparison (redlining) library.

Compare Word (.docx) and PDF documents, generating redlined output showing:
- Deletions in red strikethrough
- Insertions in blue bold
- Moved text in green
- Cell-level table comparison
"""

from .document_compare import (
    compare_documents,
    ComparisonResult,
)

from .compare_preserve_formatting import (
    compare_with_full_formatting,
)

from .pdf_support import (
    PdfParser,
    PdfGenerator,
    extract_paragraphs_from_pdf,
)

__version__ = "1.0.0"
__all__ = [
    "compare_documents",
    "compare_with_full_formatting",
    "ComparisonResult",
    "PdfParser",
    "PdfGenerator",
    "extract_paragraphs_from_pdf",
]
