"""
PDF Support Module for Document Comparison

Provides:
1. PDF text extraction with structure detection
2. PDF redline generation
3. Integration with Word comparison
"""

import fitz  # PyMuPDF
import os
import re
from dataclasses import dataclass
from typing import List, Optional, Tuple
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.colors import red, blue, black, green
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from datetime import datetime


@dataclass
class ExtractedParagraph:
    """Represents a paragraph extracted from a PDF."""
    text: str
    page_num: int
    bbox: Optional[Tuple[float, float, float, float]] = None  # x0, y0, x1, y1
    font_size: Optional[float] = None
    is_bold: bool = False
    is_heading: bool = False


class PdfParser:
    """Extracts text and structure from PDF documents."""

    def __init__(self, filepath: str):
        self.filepath = filepath
        self.doc = fitz.open(filepath)

    def close(self):
        self.doc.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def extract_paragraphs(self) -> List[ExtractedParagraph]:
        """Extract paragraphs from the PDF with basic structure detection."""
        paragraphs = []

        for page_num in range(len(self.doc)):
            page = self.doc[page_num]
            blocks = page.get_text("dict")["blocks"]

            for block in blocks:
                if block["type"] == 0:  # Text block
                    block_text = []
                    max_font_size = 0
                    is_bold = False

                    for line in block.get("lines", []):
                        line_text = []
                        for span in line.get("spans", []):
                            text = span.get("text", "")
                            if text.strip():
                                line_text.append(text)
                                font_size = span.get("size", 12)
                                max_font_size = max(max_font_size, font_size)
                                # Check for bold font
                                font_name = span.get("font", "").lower()
                                if "bold" in font_name:
                                    is_bold = True

                        if line_text:
                            block_text.append(" ".join(line_text))

                    if block_text:
                        full_text = " ".join(block_text)
                        # Detect headings based on font size
                        is_heading = max_font_size > 14

                        paragraphs.append(ExtractedParagraph(
                            text=full_text,
                            page_num=page_num + 1,
                            bbox=block.get("bbox"),
                            font_size=max_font_size,
                            is_bold=is_bold,
                            is_heading=is_heading
                        ))

        return paragraphs

    def extract_text(self) -> str:
        """Extract plain text from the entire PDF."""
        text_parts = []
        for page in self.doc:
            text_parts.append(page.get_text())
        return "\n".join(text_parts)

    def get_page_count(self) -> int:
        return len(self.doc)


class PdfGenerator:
    """Generates PDF documents with redline formatting."""

    def __init__(self, output_path: str):
        self.output_path = output_path
        self.styles = getSampleStyleSheet()
        self._setup_styles()

    def _setup_styles(self):
        """Set up custom styles for redline formatting."""
        # Normal text
        self.styles.add(ParagraphStyle(
            name='Normal_Custom',
            parent=self.styles['Normal'],
            fontSize=11,
            leading=14,
            spaceAfter=6
        ))

        # Deleted text (red strikethrough)
        self.styles.add(ParagraphStyle(
            name='Deleted',
            parent=self.styles['Normal'],
            fontSize=11,
            leading=14,
            textColor=red,
            spaceAfter=0
        ))

        # Inserted text (blue bold)
        self.styles.add(ParagraphStyle(
            name='Inserted',
            parent=self.styles['Normal'],
            fontSize=11,
            leading=14,
            textColor=blue,
            spaceAfter=0
        ))

        # Heading
        self.styles.add(ParagraphStyle(
            name='Heading_Custom',
            parent=self.styles['Heading1'],
            fontSize=14,
            leading=18,
            spaceAfter=12,
            spaceBefore=12
        ))

    def generate_redline(self, diff_paragraphs: List[dict], stats: dict = None,
                         original_name: str = "", modified_name: str = ""):
        """
        Generate a redlined PDF from diff results with Litera-style formatting.

        diff_paragraphs: List of dicts with 'segments' key containing
                        list of (text, type) tuples where type is 'equal', 'delete', or 'insert'
        stats: Statistics dictionary for summary page
        original_name: Original filename for summary
        modified_name: Modified filename for summary
        """
        doc = SimpleDocTemplate(
            self.output_path,
            pagesize=letter,
            rightMargin=72,
            leftMargin=72,
            topMargin=72,
            bottomMargin=72
        )

        story = []

        for para_info in diff_paragraphs:
            segments = para_info.get('segments', [])
            is_heading = para_info.get('is_heading', False)

            if not segments:
                continue

            # Build formatted paragraph text
            formatted_parts = []
            for text, seg_type in segments:
                if not text:
                    continue

                # Escape special XML characters
                escaped_text = self._escape_xml(text)

                if seg_type == 'delete':
                    # Red strikethrough (unchanged)
                    formatted_parts.append(
                        f'<font color="red"><strike>{escaped_text}</strike></font>'
                    )
                elif seg_type == 'insert':
                    # Blue bold for insertions
                    formatted_parts.append(
                        f'<font color="blue"><b>{escaped_text}</b></font>'
                    )
                elif seg_type == 'move_source':
                    # Green strikethrough (moved from here) - unchanged
                    formatted_parts.append(
                        f'<font color="green"><strike>{escaped_text}</strike></font>'
                    )
                elif seg_type == 'move_dest':
                    # Green with underline (moved to here)
                    formatted_parts.append(
                        f'<font color="green"><u>{escaped_text}</u></font>'
                    )
                else:
                    # Normal
                    formatted_parts.append(escaped_text)

            if formatted_parts:
                full_text = "".join(formatted_parts)
                style = self.styles['Heading_Custom'] if is_heading else self.styles['Normal_Custom']
                try:
                    para = Paragraph(full_text, style)
                    story.append(para)
                except Exception as e:
                    # If paragraph fails, add as plain text
                    story.append(Paragraph(self._escape_xml(" ".join(
                        text for text, _ in segments if text
                    )), self.styles['Normal_Custom']))

        # Add summary page if stats provided
        if stats:
            story.append(PageBreak())
            story.extend(self._generate_summary_page(stats, original_name, modified_name))

        if story:
            doc.build(story)
        else:
            # Create empty document with message
            doc.build([Paragraph("No differences found.", self.styles['Normal_Custom'])])

    def _generate_summary_page(self, stats: dict, original_name: str, modified_name: str) -> List:
        """Generate Litera-style summary report elements for PDF."""
        elements = []

        # Build table data
        data = [
            ["Summary report:", ""],
            ["Document Compare for Word Document comparison done on", ""],
            [datetime.now().strftime("%m/%d/%Y %I:%M:%S %p"), ""],
            ["Style name:", "CSM"],
            ["Intelligent Table Comparison:", "Active"],
            ["Original filename:", original_name],
            ["Modified filename:", modified_name],
            ["Changes:", ""],
            ["Add", str(stats.get('insertions', 0))],
            ["Delete", str(stats.get('deletions', 0))],
            ["Move From", str(stats.get('move_from', 0))],
            ["Move To", str(stats.get('move_to', 0))],
            ["Table Insert", str(stats.get('table_insertions', 0))],
            ["Table Delete", str(stats.get('table_deletions', 0))],
            ["Table moves to", str(stats.get('table_moves_to', 0))],
            ["Table moves from", str(stats.get('table_moves_from', 0))],
            ["Embedded Graphics (Visio, ChemDraw, Images etc.)", str(stats.get('embedded_graphics', 0))],
            ["Embedded Excel", str(stats.get('embedded_excel', 0))],
            ["Format changes", str(stats.get('format_changes', 0))],
        ]

        # Calculate total
        total = (stats.get('insertions', 0) + stats.get('deletions', 0) +
                 stats.get('move_from', 0) + stats.get('move_to', 0) +
                 stats.get('table_insertions', 0) + stats.get('table_deletions', 0) +
                 stats.get('table_moves_to', 0) + stats.get('table_moves_from', 0))
        data.append(["Total Changes:", str(total)])

        # Create table
        table = Table(data, colWidths=[350, 80])

        # Style the table
        style = TableStyle([
            ('BOX', (0, 0), (-1, -1), 1, colors.black),
            ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('FONTNAME', (0, 0), (0, 0), 'Helvetica-Bold'),  # Summary report bold
            ('SPAN', (0, 0), (1, 0)),  # Merge header cells
            ('SPAN', (0, 1), (1, 1)),  # Merge second row
            ('SPAN', (0, 2), (1, 2)),  # Merge third row
            ('SPAN', (0, 7), (1, 7)),  # Merge Changes row
            ('FONTNAME', (0, 7), (0, 7), 'Helvetica-Bold'),  # Changes bold
            ('FONTNAME', (0, -1), (0, -1), 'Helvetica-Bold'),  # Total bold
            # Color formatting for change rows
            ('TEXTCOLOR', (0, 8), (0, 8), colors.blue),    # Add
            ('TEXTCOLOR', (0, 9), (0, 9), colors.red),     # Delete
            ('TEXTCOLOR', (0, 10), (0, 10), colors.green), # Move From
            ('TEXTCOLOR', (0, 11), (0, 11), colors.green), # Move To
            ('TEXTCOLOR', (0, 12), (0, 12), colors.green), # Table Insert
            ('TEXTCOLOR', (0, 13), (0, 13), colors.red),   # Table Delete
            ('TEXTCOLOR', (0, 14), (0, 14), colors.green), # Table moves to
            ('TEXTCOLOR', (0, 15), (0, 15), colors.green), # Table moves from
        ])
        table.setStyle(style)

        elements.append(table)
        return elements

    def _escape_xml(self, text: str) -> str:
        """Escape special XML characters for ReportLab."""
        text = text.replace("&", "&amp;")
        text = text.replace("<", "&lt;")
        text = text.replace(">", "&gt;")
        text = text.replace('"', "&quot;")
        text = text.replace("'", "&apos;")
        return text


def extract_paragraphs_from_pdf(pdf_path: str) -> List[ExtractedParagraph]:
    """Helper function to extract paragraphs from a PDF."""
    with PdfParser(pdf_path) as parser:
        return parser.extract_paragraphs()


def convert_pdf_paragraphs_to_text_list(paragraphs: List[ExtractedParagraph]) -> List[str]:
    """Convert extracted paragraphs to simple text list for comparison."""
    return [p.text for p in paragraphs]


# For integration with the main comparison module
class PdfDocumentAdapter:
    """
    Adapter class that makes PDF documents compatible with the comparison engine.
    Mimics the interface used by python-docx Document.
    """

    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        self._paragraphs = None
        self._load()

    def _load(self):
        with PdfParser(self.pdf_path) as parser:
            extracted = parser.extract_paragraphs()
            # Create paragraph-like objects
            self._paragraphs = [PdfParagraphAdapter(p) for p in extracted]

    @property
    def paragraphs(self):
        return self._paragraphs

    @property
    def sections(self):
        # PDFs don't have sections in the same way, return empty list
        return []


class PdfParagraphAdapter:
    """Adapter to make PDF paragraphs look like docx paragraphs."""

    def __init__(self, extracted: ExtractedParagraph):
        self._extracted = extracted
        self._text = extracted.text
        self._runs = [PdfRunAdapter(extracted.text)]

    @property
    def text(self):
        return self._text

    @property
    def runs(self):
        return self._runs

    @property
    def is_heading(self):
        return self._extracted.is_heading


class PdfRunAdapter:
    """Adapter to make PDF text runs look like docx runs."""

    def __init__(self, text: str):
        self._text = text
        self.bold = False
        self.italic = False
        self.font = PdfFontAdapter()

    @property
    def text(self):
        return self._text


class PdfFontAdapter:
    """Adapter for font properties."""

    def __init__(self):
        self.name = None
        self.size = None
        self.strike = False
        self.color = PdfColorAdapter()


class PdfColorAdapter:
    """Adapter for color properties."""

    def __init__(self):
        self.rgb = None


if __name__ == "__main__":
    # Test PDF extraction
    import sys

    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
        print(f"Extracting from: {pdf_path}")

        with PdfParser(pdf_path) as parser:
            paragraphs = parser.extract_paragraphs()
            print(f"\nFound {len(paragraphs)} paragraphs:")
            for i, para in enumerate(paragraphs[:10]):  # Show first 10
                preview = para.text[:80] + "..." if len(para.text) > 80 else para.text
                print(f"  [{i}] Page {para.page_num}: {preview}")
    else:
        print("Usage: python pdf_support.py <pdf_file>")
