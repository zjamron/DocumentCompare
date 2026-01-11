"""
Document comparison that PRESERVES original formatting.

Strategy:
1. Copy the MODIFIED document as the base (keeps all formatting intact)
2. For each paragraph, find what was deleted and insert it with strikethrough
3. Mark inserted text with blue bold
"""
import difflib
import re
import copy
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree
import os
import shutil

# Word XML namespaces
WORD_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

def get_paragraph_text(para):
    """Get plain text from a paragraph."""
    return para.text or ''

def tokenize(text):
    """Split text into words while preserving whitespace."""
    words = []
    for match in re.finditer(r'\S+|\s+', text):
        words.append(match.group())
    return words

def diff_texts(original_text, modified_text):
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

def get_first_run_formatting(para):
    """Extract formatting from the first run of a paragraph."""
    formatting = {}
    if para.runs:
        run = para.runs[0]
        formatting['bold'] = run.bold
        formatting['italic'] = run.italic
        formatting['underline'] = run.underline
        formatting['font_name'] = run.font.name
        formatting['font_size'] = run.font.size
    return formatting

def apply_base_formatting(run, formatting):
    """Apply base formatting to a run."""
    if formatting.get('font_name'):
        run.font.name = formatting['font_name']
    if formatting.get('font_size'):
        run.font.size = formatting['font_size']
    # Don't copy bold/italic as we may override them

def clear_paragraph_content(para):
    """Clear all run content from paragraph, preserving paragraph properties."""
    p_element = para._p

    # Find and remove all content elements (runs, hyperlinks, bookmarks, etc.)
    # Keep only pPr (paragraph properties)
    elements_to_remove = []
    for child in p_element:
        tag = child.tag
        # Keep paragraph properties (pPr), remove everything else
        if not tag.endswith('}pPr'):
            elements_to_remove.append(child)

    for elem in elements_to_remove:
        p_element.remove(elem)

def rebuild_paragraph_with_diff(para, diff_segments, base_formatting=None):
    """
    Completely rebuild paragraph content with diff segments.
    """
    # First, completely clear the paragraph of all runs
    clear_paragraph_content(para)

    # Now add new runs for each segment
    for text, seg_type in diff_segments:
        if not text:
            continue

        run = para.add_run(text)

        # Apply base formatting
        if base_formatting:
            apply_base_formatting(run, base_formatting)

        # Apply diff formatting
        if seg_type == 'delete':
            run.font.strike = True
            run.font.color.rgb = RGBColor(255, 0, 0)  # Red
        elif seg_type == 'insert':
            run.bold = True
            run.font.color.rgb = RGBColor(0, 0, 255)  # Blue
        # 'equal' keeps normal formatting

def apply_formatting_to_existing_runs(para, formatting_type):
    """Apply redline formatting to all existing runs in a paragraph."""
    for run in para.runs:
        if formatting_type == 'delete':
            run.font.strike = True
            run.font.color.rgb = RGBColor(255, 0, 0)
        elif formatting_type == 'insert':
            run.bold = True
            run.font.color.rgb = RGBColor(0, 0, 255)

def calculate_similarity(text1, text2):
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

    # Method 1: Word-based Jaccard similarity
    words1 = set(text1.lower().split())
    words2 = set(text2.lower().split())

    word_sim = 0.0
    if words1 and words2:
        intersection = len(words1 & words2)
        union = len(words1 | words2)
        word_sim = intersection / union if union > 0 else 0.0

    # Method 2: Character sequence similarity (better for small edits)
    seq_sim = difflib.SequenceMatcher(None, text1.lower(), text2.lower()).ratio()

    # Use the higher of the two similarities
    return max(word_sim, seq_sim)

def align_paragraphs(orig_paras, mod_paras):
    """Align paragraphs between documents using LCS."""
    m, n = len(orig_paras), len(mod_paras)

    # Build LCS table
    lcs = [[0] * (n + 1) for _ in range(m + 1)]

    for i in range(1, m + 1):
        for j in range(1, n + 1):
            orig_text = get_paragraph_text(orig_paras[i-1])
            mod_text = get_paragraph_text(mod_paras[j-1])

            if calculate_similarity(orig_text, mod_text) >= 0.4:
                lcs[i][j] = lcs[i-1][j-1] + 1
            else:
                lcs[i][j] = max(lcs[i-1][j], lcs[i][j-1])

    # Backtrack
    alignments = []
    i, j = m, n

    while i > 0 or j > 0:
        if i > 0 and j > 0:
            orig_text = get_paragraph_text(orig_paras[i-1])
            mod_text = get_paragraph_text(mod_paras[j-1])
            if calculate_similarity(orig_text, mod_text) >= 0.4:
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

def compare_paragraphs_list(orig_paras, out_paras, stats, context=""):
    """
    Compare two lists of paragraphs and apply diff formatting.
    """
    if not orig_paras and not out_paras:
        return

    # Align paragraphs
    alignments = align_paragraphs(orig_paras, out_paras)

    for orig_idx, mod_idx, align_type in alignments:
        if align_type == 'match' and mod_idx >= 0 and orig_idx >= 0:
            orig_para = orig_paras[orig_idx]
            out_para = out_paras[mod_idx]

            orig_text = get_paragraph_text(orig_para)
            mod_text = get_paragraph_text(out_para)

            # Only process if there's actual text and it differs
            if orig_text.strip() != mod_text.strip() and (orig_text.strip() or mod_text.strip()):
                # Get base formatting from output paragraph
                base_formatting = get_first_run_formatting(out_para)

                # Compute diff
                diff_segments = diff_texts(orig_text, mod_text)

                # Count changes
                for text, seg_type in diff_segments:
                    words = len(text.split())
                    if seg_type == 'insert':
                        stats['insertions'] += words
                    elif seg_type == 'delete':
                        stats['deletions'] += words
                    else:
                        stats['unchanged'] += words

                # Rebuild paragraph with diff formatting
                rebuild_paragraph_with_diff(out_para, diff_segments, base_formatting)
            else:
                stats['unchanged'] += len(mod_text.split())

        elif align_type == 'insert' and mod_idx >= 0:
            out_para = out_paras[mod_idx]
            text = get_paragraph_text(out_para)
            if text.strip():
                apply_formatting_to_existing_runs(out_para, 'insert')
                stats['insertions'] += len(text.split())

        elif align_type == 'delete' and orig_idx >= 0:
            orig_para = orig_paras[orig_idx]
            text = get_paragraph_text(orig_para)
            if text.strip():
                stats['deletions'] += len(text.split())


def compare_with_full_formatting(original_path, modified_path, output_path):
    """
    Compare documents while preserving formatting.
    Uses modified document as base and applies diff formatting.
    Compares body, headers, and footers.
    """
    print("Loading documents...")

    # Start by copying the modified file as our base
    shutil.copy2(modified_path, output_path)

    # Now open both original and output
    original_doc = Document(original_path)
    output_doc = Document(output_path)

    stats = {'insertions': 0, 'deletions': 0, 'unchanged': 0}

    # Compare main body paragraphs
    orig_paras = list(original_doc.paragraphs)
    out_paras = list(output_doc.paragraphs)

    print(f"Original: {len(orig_paras)} body paragraphs")
    print(f"Modified: {len(out_paras)} body paragraphs")

    print("Comparing body paragraphs...")
    compare_paragraphs_list(orig_paras, out_paras, stats, "body")

    # Compare headers and footers for each section
    orig_sections = list(original_doc.sections)
    out_sections = list(output_doc.sections)

    print(f"Sections: {len(orig_sections)} original, {len(out_sections)} modified")

    # Process headers and footers
    for i in range(min(len(orig_sections), len(out_sections))):
        orig_section = orig_sections[i]
        out_section = out_sections[i]

        # Compare headers
        try:
            if orig_section.header and out_section.header:
                orig_header_paras = list(orig_section.header.paragraphs)
                out_header_paras = list(out_section.header.paragraphs)
                if orig_header_paras or out_header_paras:
                    print(f"Comparing header (section {i+1})...")
                    compare_paragraphs_list(orig_header_paras, out_header_paras, stats, f"header_{i}")
        except Exception as e:
            print(f"  Warning: Could not compare header section {i+1}: {e}")

        # Compare first page header if different
        try:
            if orig_section.first_page_header and out_section.first_page_header:
                orig_fp_header_paras = list(orig_section.first_page_header.paragraphs)
                out_fp_header_paras = list(out_section.first_page_header.paragraphs)
                if orig_fp_header_paras or out_fp_header_paras:
                    print(f"Comparing first page header (section {i+1})...")
                    compare_paragraphs_list(orig_fp_header_paras, out_fp_header_paras, stats, f"fp_header_{i}")
        except Exception as e:
            pass  # First page header may not exist

        # Compare footers
        try:
            if orig_section.footer and out_section.footer:
                orig_footer_paras = list(orig_section.footer.paragraphs)
                out_footer_paras = list(out_section.footer.paragraphs)
                if orig_footer_paras or out_footer_paras:
                    print(f"Comparing footer (section {i+1})...")
                    compare_paragraphs_list(orig_footer_paras, out_footer_paras, stats, f"footer_{i}")
        except Exception as e:
            print(f"  Warning: Could not compare footer section {i+1}: {e}")

        # Compare first page footer if different
        try:
            if orig_section.first_page_footer and out_section.first_page_footer:
                orig_fp_footer_paras = list(orig_section.first_page_footer.paragraphs)
                out_fp_footer_paras = list(out_section.first_page_footer.paragraphs)
                if orig_fp_footer_paras or out_fp_footer_paras:
                    print(f"Comparing first page footer (section {i+1})...")
                    compare_paragraphs_list(orig_fp_footer_paras, out_fp_footer_paras, stats, f"fp_footer_{i}")
        except Exception as e:
            pass  # First page footer may not exist

    print(f"Saving to {output_path}...")
    output_doc.save(output_path)

    return stats

if __name__ == '__main__':
    import sys

    if len(sys.argv) >= 4:
        original = sys.argv[1]
        modified = sys.argv[2]
        output = sys.argv[3]
    else:
        # Default test files
        samples_dir = os.path.dirname(os.path.abspath(__file__))
        original = os.path.join(samples_dir, 'contract_v1.docx')
        modified = os.path.join(samples_dir, 'contract_v2.docx')
        output = os.path.join(samples_dir, 'contract_redline_formatted.docx')

    print("="*60)
    print("Document Comparison (Formatting Preserved)")
    print("="*60)
    print(f"Original: {original}")
    print(f"Modified: {modified}")
    print(f"Output:   {output}")
    print()

    stats = compare_with_full_formatting(original, modified, output)

    print()
    print("="*60)
    print("COMPLETE")
    print("="*60)
    print(f"Output: {output}")
    print(f"Insertions (blue bold): {stats['insertions']} words")
    print(f"Deletions (red strikethrough): {stats['deletions']} words")
    print(f"Unchanged: {stats['unchanged']} words")
