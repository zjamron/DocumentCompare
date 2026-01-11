"""
Universal Document Comparison Tool

Supports:
- Word (.docx) to Word comparison
- PDF to PDF comparison
- Word to PDF comparison (and vice versa)
- Output to Word or PDF format
"""

import os
import sys
import difflib
import re
import shutil
from typing import List, Tuple, Optional
from dataclasses import dataclass

# Import Word support
from docx import Document
from docx.shared import RGBColor

# Import PDF support
from pdf_support import (
    PdfParser, PdfGenerator, PdfDocumentAdapter,
    ExtractedParagraph, extract_paragraphs_from_pdf
)


@dataclass
class ComparisonResult:
    """Result of a document comparison."""
    output_path: str
    insertions: int
    deletions: int
    unchanged: int
    success: bool
    error: Optional[str] = None


def get_file_type(filepath: str) -> str:
    """Determine file type from extension."""
    ext = os.path.splitext(filepath)[1].lower()
    if ext == '.docx':
        return 'word'
    elif ext == '.pdf':
        return 'pdf'
    else:
        raise ValueError(f"Unsupported file type: {ext}")


def tokenize(text: str) -> List[str]:
    """Split text into words while preserving whitespace."""
    return re.findall(r'\S+|\s+', text)


def diff_texts(original_text: str, modified_text: str) -> List[Tuple[str, str]]:
    """
    Compute word-level diff between two texts.
    Returns list of (text, type) where type is 'equal', 'delete', or 'insert'
    """
    orig_words = tokenize(original_text)
    mod_words = tokenize(modified_text)

    matcher = difflib.SequenceMatcher(None, orig_words, mod_words, autojunk=False)
    result = []

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            text = ''.join(orig_words[i1:i2])
            if text:
                result.append((text, 'equal'))
        elif tag == 'delete':
            text = ''.join(orig_words[i1:i2])
            if text:
                result.append((text, 'delete'))
        elif tag == 'insert':
            text = ''.join(mod_words[j1:j2])
            if text:
                result.append((text, 'insert'))
        elif tag == 'replace':
            del_text = ''.join(orig_words[i1:i2])
            ins_text = ''.join(mod_words[j1:j2])
            if del_text:
                result.append((del_text, 'delete'))
            if ins_text:
                result.append((ins_text, 'insert'))

    return result


def calculate_similarity(text1: str, text2: str) -> float:
    """Calculate similarity between two texts using multiple methods."""
    if text1 == text2:
        return 1.0
    if not text1 or not text2:
        return 0.0

    text1 = text1.strip()
    text2 = text2.strip()

    if not text1 and not text2:
        return 1.0
    if not text1 or not text2:
        return 0.0

    # Word-based Jaccard similarity
    words1 = set(text1.lower().split())
    words2 = set(text2.lower().split())

    word_sim = 0.0
    if words1 and words2:
        intersection = len(words1 & words2)
        union = len(words1 | words2)
        word_sim = intersection / union if union > 0 else 0.0

    # Character sequence similarity
    seq_sim = difflib.SequenceMatcher(None, text1.lower(), text2.lower()).ratio()

    return max(word_sim, seq_sim)


def align_paragraphs(orig_texts: List[str], mod_texts: List[str]) -> List[Tuple[int, int, str]]:
    """Align paragraphs using LCS algorithm."""
    m, n = len(orig_texts), len(mod_texts)

    # Build LCS table
    lcs = [[0] * (n + 1) for _ in range(m + 1)]

    for i in range(1, m + 1):
        for j in range(1, n + 1):
            if calculate_similarity(orig_texts[i-1], mod_texts[j-1]) >= 0.4:
                lcs[i][j] = lcs[i-1][j-1] + 1
            else:
                lcs[i][j] = max(lcs[i-1][j], lcs[i][j-1])

    # Backtrack
    alignments = []
    i, j = m, n

    while i > 0 or j > 0:
        if i > 0 and j > 0:
            if calculate_similarity(orig_texts[i-1], mod_texts[j-1]) >= 0.4:
                alignments.append((i-1, j-1, 'match'))
                i -= 1
                j -= 1
                continue

        if j > 0 and (i == 0 or lcs[i][j-1] >= lcs[i-1][j]):
            alignments.append((-1, j-1, 'insert'))
            j -= 1
        else:
            alignments.append((i-1, -1, 'delete'))
            i -= 1

    alignments.reverse()
    return alignments


def extract_paragraphs_from_word(doc_path: str) -> List[Tuple[str, dict]]:
    """Extract paragraphs from Word document with metadata."""
    doc = Document(doc_path)
    result = []
    for para in doc.paragraphs:
        metadata = {
            'is_heading': para.style.name.startswith('Heading') if para.style else False
        }
        result.append((para.text, metadata))
    return result


def extract_paragraphs_from_document(doc_path: str) -> List[Tuple[str, dict]]:
    """Extract paragraphs from any supported document type."""
    file_type = get_file_type(doc_path)

    if file_type == 'word':
        return extract_paragraphs_from_word(doc_path)
    elif file_type == 'pdf':
        with PdfParser(doc_path) as parser:
            extracted = parser.extract_paragraphs()
            return [(p.text, {'is_heading': p.is_heading}) for p in extracted]


def compare_documents(
    original_path: str,
    modified_path: str,
    output_path: str,
    output_format: Optional[str] = None
) -> ComparisonResult:
    """
    Compare two documents and generate a redlined output.

    Args:
        original_path: Path to original document (Word or PDF)
        modified_path: Path to modified document (Word or PDF)
        output_path: Path for output redlined document
        output_format: 'word' or 'pdf' (auto-detected from output_path if not specified)

    Returns:
        ComparisonResult with statistics and status
    """
    try:
        # Determine output format
        if output_format is None:
            output_format = get_file_type(output_path)

        print(f"Original: {original_path} ({get_file_type(original_path)})")
        print(f"Modified: {modified_path} ({get_file_type(modified_path)})")
        print(f"Output: {output_path} ({output_format})")
        print()

        # Extract paragraphs from both documents
        print("Extracting paragraphs...")
        orig_paras = extract_paragraphs_from_document(original_path)
        mod_paras = extract_paragraphs_from_document(modified_path)

        print(f"  Original: {len(orig_paras)} paragraphs")
        print(f"  Modified: {len(mod_paras)} paragraphs")

        # Get just the text for alignment
        orig_texts = [text for text, _ in orig_paras]
        mod_texts = [text for text, _ in mod_paras]

        # Align paragraphs
        print("Aligning paragraphs...")
        alignments = align_paragraphs(orig_texts, mod_texts)

        # Compute diffs
        print("Computing differences...")
        diff_results = []
        stats = {'insertions': 0, 'deletions': 0, 'unchanged': 0}

        for orig_idx, mod_idx, align_type in alignments:
            if align_type == 'match':
                orig_text = orig_texts[orig_idx]
                mod_text = mod_texts[mod_idx]
                mod_meta = mod_paras[mod_idx][1]

                if orig_text.strip() != mod_text.strip():
                    # Compute word-level diff
                    segments = diff_texts(orig_text, mod_text)
                else:
                    segments = [(mod_text, 'equal')]

                # Count stats
                for text, seg_type in segments:
                    words = len(text.split())
                    if seg_type == 'insert':
                        stats['insertions'] += words
                    elif seg_type == 'delete':
                        stats['deletions'] += words
                    else:
                        stats['unchanged'] += words

                diff_results.append({
                    'segments': segments,
                    'is_heading': mod_meta.get('is_heading', False)
                })

            elif align_type == 'insert':
                mod_text = mod_texts[mod_idx]
                mod_meta = mod_paras[mod_idx][1]
                if mod_text.strip():
                    diff_results.append({
                        'segments': [(mod_text, 'insert')],
                        'is_heading': mod_meta.get('is_heading', False)
                    })
                    stats['insertions'] += len(mod_text.split())

            elif align_type == 'delete':
                orig_text = orig_texts[orig_idx]
                orig_meta = orig_paras[orig_idx][1]
                if orig_text.strip():
                    diff_results.append({
                        'segments': [(orig_text, 'delete')],
                        'is_heading': orig_meta.get('is_heading', False)
                    })
                    stats['deletions'] += len(orig_text.split())

        # Generate output
        print(f"Generating {output_format} output...")

        if output_format == 'pdf':
            generate_pdf_redline(diff_results, output_path)
        else:  # word
            generate_word_redline(diff_results, output_path, modified_path)

        print(f"\nOutput saved to: {output_path}")

        return ComparisonResult(
            output_path=output_path,
            insertions=stats['insertions'],
            deletions=stats['deletions'],
            unchanged=stats['unchanged'],
            success=True
        )

    except Exception as e:
        import traceback
        traceback.print_exc()
        return ComparisonResult(
            output_path=output_path,
            insertions=0,
            deletions=0,
            unchanged=0,
            success=False,
            error=str(e)
        )


def generate_pdf_redline(diff_results: List[dict], output_path: str):
    """Generate a redlined PDF document."""
    generator = PdfGenerator(output_path)
    generator.generate_redline(diff_results)


def generate_word_redline(diff_results: List[dict], output_path: str, base_doc_path: Optional[str] = None):
    """Generate a redlined Word document."""
    doc = Document()

    for para_info in diff_results:
        segments = para_info.get('segments', [])
        is_heading = para_info.get('is_heading', False)

        if not segments:
            continue

        if is_heading:
            para = doc.add_heading('', level=1)
        else:
            para = doc.add_paragraph()

        for text, seg_type in segments:
            if not text:
                continue

            run = para.add_run(text)

            if seg_type == 'delete':
                run.font.strike = True
                run.font.color.rgb = RGBColor(255, 0, 0)  # Red
            elif seg_type == 'insert':
                run.bold = True
                run.font.color.rgb = RGBColor(0, 0, 255)  # Blue

    doc.save(output_path)


def main():
    """Command-line interface."""
    import argparse

    parser = argparse.ArgumentParser(
        description='Compare two documents and generate a redlined output.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s original.docx modified.docx redline.docx
  %(prog)s original.pdf modified.pdf redline.pdf
  %(prog)s original.docx modified.pdf redline.pdf
        """
    )
    parser.add_argument('original', help='Path to original document (Word or PDF)')
    parser.add_argument('modified', help='Path to modified document (Word or PDF)')
    parser.add_argument('output', help='Path for output redlined document')
    parser.add_argument('--format', choices=['word', 'pdf'],
                       help='Output format (default: auto-detect from output extension)')

    args = parser.parse_args()

    print("=" * 60)
    print("Document Comparison Tool")
    print("=" * 60)
    print()

    result = compare_documents(
        args.original,
        args.modified,
        args.output,
        args.format
    )

    print()
    print("=" * 60)
    if result.success:
        print("COMPARISON COMPLETE")
        print("=" * 60)
        print(f"Output: {result.output_path}")
        print(f"Insertions (blue bold): {result.insertions} words")
        print(f"Deletions (red strikethrough): {result.deletions} words")
        print(f"Unchanged: {result.unchanged} words")

        total = result.insertions + result.deletions + result.unchanged
        if total > 0:
            change_pct = (result.insertions + result.deletions) * 100 / total
            print(f"Change percentage: {change_pct:.1f}%")
    else:
        print("COMPARISON FAILED")
        print("=" * 60)
        print(f"Error: {result.error}")
        sys.exit(1)


if __name__ == '__main__':
    main()
