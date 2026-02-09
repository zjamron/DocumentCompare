"""
XML-Based Redline Generator for Word Documents

This module provides true document structure preservation by working directly
with the Open XML (.docx) format. Unlike the previous approach that created
a new blank document, this copies the modified document and applies redline
markup at the XML level, preserving all formatting, headers, footers, styles,
numbering, tables, and footnotes.

Key approach:
1. Copy modified document as base (preserves ALL structure)
2. Extract paragraph texts from original document
3. Parse document.xml with lxml
4. Apply word-level diff with redline formatting
5. Insert deleted paragraphs at correct positions
6. Save modified XML back to the docx archive
"""

import os
import re
import shutil
import difflib
from zipfile import ZipFile
from lxml import etree
from typing import List, Tuple, Dict, Optional
from io import BytesIO

# Word Open XML namespaces
WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NSMAP = {'w': WORD_NS}

# Namespace prefix for element creation
W = f'{{{WORD_NS}}}'

# Move detection settings
MOVE_SIMILARITY_THRESHOLD = 0.85
MIN_MOVE_WORDS = 3


class XmlRedlineGenerator:
    """
    Generates redlined Word documents by manipulating Open XML directly.

    This preserves all document structure (headers, footers, styles, tables,
    footnotes) while applying redline markup to show changes.
    """

    def __init__(self, modified_path: str, original_path: str, output_path: str):
        """
        Initialize the generator.

        Args:
            modified_path: Path to the modified (newer) document
            original_path: Path to the original (baseline) document
            output_path: Path for the output redlined document
        """
        self.modified_path = modified_path
        self.original_path = original_path
        self.output_path = output_path

        self.stats = {
            'insertions': 0,
            'deletions': 0,
            'moves': 0,
            'unchanged': 0,
            'move_from': 0,
            'move_to': 0,
        }

    def generate(self) -> Dict:
        """
        Generate the redlined document.

        Returns:
            Statistics dictionary with change counts
        """
        # Step 1: Copy modified document as base (preserves all structure)
        shutil.copy(self.modified_path, self.output_path)

        # Step 2: Extract paragraph texts from original
        orig_paras = self._extract_paragraph_texts(self.original_path)

        # Step 3: Process the output document
        self._process_document(orig_paras)

        return self.stats

    def _extract_paragraph_texts(self, doc_path: str) -> List[str]:
        """Extract paragraph texts from a Word document."""
        texts = []

        with ZipFile(doc_path, 'r') as zf:
            if 'word/document.xml' not in zf.namelist():
                return texts

            doc_xml = zf.read('word/document.xml')
            root = etree.fromstring(doc_xml)

            # Find all paragraphs
            for para in root.findall('.//w:p', NSMAP):
                text = self._get_para_text(para)
                texts.append(text)

        return texts

    def _get_para_text(self, para_elem) -> str:
        """Extract text content from a paragraph element."""
        texts = []
        for t in para_elem.findall('.//w:t', NSMAP):
            if t.text:
                texts.append(t.text)
        return ''.join(texts)

    def _extract_footnotes(self, doc_path: str) -> Dict[str, str]:
        """Extract footnotes from a Word document as id -> text mapping."""
        footnotes = {}

        try:
            with ZipFile(doc_path, 'r') as zf:
                if 'word/footnotes.xml' not in zf.namelist():
                    return footnotes

                footnotes_xml = zf.read('word/footnotes.xml')
                root = etree.fromstring(footnotes_xml)

                # Find all footnotes
                for fn in root.findall('.//w:footnote', NSMAP):
                    fn_id = fn.get(f'{W}id')
                    # Skip separator footnotes (id -1 and 0)
                    if fn_id in ['-1', '0']:
                        continue

                    # Get text from all paragraphs in the footnote
                    text_parts = []
                    for para in fn.findall('.//w:p', NSMAP):
                        para_text = self._get_para_text(para)
                        if para_text.strip():
                            text_parts.append(para_text)

                    if text_parts:
                        footnotes[fn_id] = ' '.join(text_parts)

        except Exception as e:
            print(f"Warning: Could not extract footnotes: {e}")

        return footnotes

    def _process_footnotes(self, footnotes_xml: bytes, orig_footnotes: Dict[str, str]) -> bytes:
        """Process footnotes.xml to apply redline markup."""
        root = etree.fromstring(footnotes_xml)

        # Find all footnotes and compare with original
        for fn in root.findall('.//w:footnote', NSMAP):
            fn_id = fn.get(f'{W}id')
            # Skip separator footnotes
            if fn_id in ['-1', '0']:
                continue

            # Get current text
            mod_text_parts = []
            for para in fn.findall('.//w:p', NSMAP):
                para_text = self._get_para_text(para)
                if para_text.strip():
                    mod_text_parts.append(para_text)
            mod_text = ' '.join(mod_text_parts)

            # Get original text
            orig_text = orig_footnotes.get(fn_id, '')

            if orig_text and mod_text and orig_text.strip() != mod_text.strip():
                # Footnote was modified - apply redline to paragraphs
                for para in fn.findall('.//w:p', NSMAP):
                    para_text = self._get_para_text(para)
                    if para_text.strip():
                        # Find corresponding original portion
                        self._apply_redline_to_para(para, orig_text, mod_text)
                        break  # Only process first paragraph for simplicity

            elif not orig_text and mod_text:
                # New footnote - mark as inserted
                for para in fn.findall('.//w:p', NSMAP):
                    self._mark_para_as_inserted(para)

        return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone='yes')

    def _process_document(self, orig_paras: List[str]):
        """Process the document XML and apply redlines."""
        # Read the document
        with ZipFile(self.output_path, 'r') as zf:
            doc_xml = zf.read('word/document.xml')
            all_files = {name: zf.read(name) for name in zf.namelist()}

        # Also extract original footnotes for comparison
        orig_footnotes = self._extract_footnotes(self.original_path)

        # Parse the document
        root = etree.fromstring(doc_xml)

        # Get body element
        body = root.find('.//w:body', NSMAP)
        if body is None:
            return

        # Get all paragraphs in modified document
        mod_para_elems = body.findall('w:p', NSMAP)
        mod_texts = [self._get_para_text(p) for p in mod_para_elems]

        # Align paragraphs between original and modified
        alignments = self._align_paragraphs(orig_paras, mod_texts)

        # Track which original paragraphs were matched
        matched_orig = set()

        # Apply redlines to matched paragraphs
        for mod_idx, (orig_idx, align_type) in alignments.items():
            if align_type == 'match' and orig_idx is not None:
                matched_orig.add(orig_idx)
                orig_text = orig_paras[orig_idx]
                mod_text = mod_texts[mod_idx]

                if orig_text.strip() != mod_text.strip():
                    self._apply_redline_to_para(mod_para_elems[mod_idx], orig_text, mod_text)
                else:
                    # Unchanged - count words
                    self.stats['unchanged'] += len(mod_text.split())

            elif align_type == 'insert':
                # Mark entire paragraph as inserted
                self._mark_para_as_inserted(mod_para_elems[mod_idx])
                self.stats['insertions'] += len(mod_texts[mod_idx].split())

        # Insert deleted paragraphs
        self._insert_deleted_paragraphs(body, mod_para_elems, orig_paras, alignments, matched_orig)

        # Serialize document
        new_doc_xml = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone='yes')
        all_files['word/document.xml'] = new_doc_xml

        # Process footnotes if present
        if 'word/footnotes.xml' in all_files and orig_footnotes:
            footnotes_xml = all_files['word/footnotes.xml']
            new_footnotes_xml = self._process_footnotes(footnotes_xml, orig_footnotes)
            all_files['word/footnotes.xml'] = new_footnotes_xml

        # Write back to the zip file
        with ZipFile(self.output_path, 'w') as zf:
            for name, content in all_files.items():
                zf.writestr(name, content)

    def _align_paragraphs(self, orig_texts: List[str], mod_texts: List[str]) -> Dict[int, Tuple[Optional[int], str]]:
        """
        Align paragraphs between original and modified documents.

        Returns:
            Dict mapping mod_para_idx -> (orig_para_idx or None, alignment_type)
            alignment_type is 'match' or 'insert'
        """
        m, n = len(orig_texts), len(mod_texts)

        if m == 0:
            # All paragraphs are new
            return {j: (None, 'insert') for j in range(n)}

        if n == 0:
            return {}

        # Build LCS table for alignment
        lcs = [[0] * (n + 1) for _ in range(m + 1)]

        for i in range(1, m + 1):
            for j in range(1, n + 1):
                if self._calculate_similarity(orig_texts[i-1], mod_texts[j-1]) >= 0.4:
                    lcs[i][j] = lcs[i-1][j-1] + 1
                else:
                    lcs[i][j] = max(lcs[i-1][j], lcs[i][j-1])

        # Backtrack to find alignment
        alignments = {}
        i, j = m, n

        while i > 0 or j > 0:
            if i > 0 and j > 0:
                if self._calculate_similarity(orig_texts[i-1], mod_texts[j-1]) >= 0.4:
                    alignments[j-1] = (i-1, 'match')
                    i -= 1
                    j -= 1
                    continue

            if j > 0 and (i == 0 or lcs[i][j-1] >= lcs[i-1][j]):
                alignments[j-1] = (None, 'insert')
                j -= 1
            else:
                i -= 1

        return alignments

    def _calculate_similarity(self, text1: str, text2: str) -> float:
        """Calculate similarity between two texts."""
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

    def _apply_redline_to_para(self, para_elem, orig_text: str, mod_text: str):
        """Apply word-level redline to a paragraph element."""
        # Compute word-level diff
        diff = self._diff_texts(orig_text, mod_text)

        # Detect moves within the diff
        diff = self._detect_moves(diff)

        # Rebuild paragraph with redline runs
        self._rebuild_para_with_redline(para_elem, diff)

    def _diff_texts(self, original: str, modified: str) -> List[Tuple[str, str]]:
        """
        Compute word-level diff between two texts with character-level refinement.

        For words that are similar but not identical (e.g., "word" vs "[word]"),
        this does character-level diff to show only the actual changes.

        Returns:
            List of (text, type) where type is 'equal', 'delete', or 'insert'
        """
        orig_words = self._tokenize(original)
        mod_words = self._tokenize(modified)

        # Normalize words for comparison (lowercase) but keep originals for output
        orig_normalized = [self._normalize_word(w) for w in orig_words]
        mod_normalized = [self._normalize_word(w) for w in mod_words]

        # Use normalized versions for matching
        matcher = difflib.SequenceMatcher(None, orig_normalized, mod_normalized, autojunk=False)
        result = []

        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == 'equal':
                # Words match when normalized - but check if they differ in original form
                for k in range(i2 - i1):
                    orig_word = orig_words[i1 + k]
                    mod_word = mod_words[j1 + k]
                    if orig_word == mod_word:
                        result.append((mod_word, 'equal'))
                    else:
                        # Words matched after normalization but differ in original
                        # This could be case difference or bracket difference
                        # Show as delete/insert since case changes matter
                        result.append((orig_word, 'delete'))
                        result.append((mod_word, 'insert'))
            elif tag == 'delete':
                text = ''.join(orig_words[i1:i2])
                if text:
                    result.append((text, 'delete'))
            elif tag == 'insert':
                text = ''.join(mod_words[j1:j2])
                if text:
                    result.append((text, 'insert'))
            elif tag == 'replace':
                # For replacements, check if we should do character-level diff
                del_text = ''.join(orig_words[i1:i2])
                ins_text = ''.join(mod_words[j1:j2])

                # Check similarity - if texts are similar, do character-level diff
                if del_text and ins_text:
                    char_diff = self._refine_replacement(del_text, ins_text)
                    result.extend(char_diff)
                else:
                    if del_text:
                        result.append((del_text, 'delete'))
                    if ins_text:
                        result.append((ins_text, 'insert'))

        return result

    def _normalize_word(self, word: str) -> str:
        """
        Normalize a word for comparison.

        Only strips brackets and quotes - case is PRESERVED because case changes matter.
        """
        # Strip brackets and quotes for comparison purposes, but keep case
        normalized = re.sub(r'[\[\]\(\)\{\}\"\'""''«»]', '', word)
        return normalized

    def _refine_replacement(self, old_text: str, new_text: str) -> List[Tuple[str, str]]:
        """
        Refine a replacement - just show clean delete/insert for word-level changes.

        Since we now tokenize punctuation separately, this should rarely be called
        for cases like "word" vs "word," - those are handled at the token level.

        Case changes DO matter and should be shown as delete/insert.
        """
        # Just show as clean word-level replacement
        result = []
        if old_text:
            result.append((old_text, 'delete'))
        if new_text:
            result.append((new_text, 'insert'))
        return result

    def _char_level_diff_ignore_case(self, old_text: str, new_text: str) -> List[Tuple[str, str]]:
        """
        Do character-level diff that ignores case changes.

        Only shows actual content changes (punctuation, brackets, etc.), not case changes.
        """
        # Compare lowercased versions to find structural differences
        old_lower = old_text.lower()
        new_lower = new_text.lower()

        matcher = difflib.SequenceMatcher(None, old_lower, new_lower, autojunk=False)
        result = []

        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == 'equal':
                # Use the new text's version (preserves its casing)
                text = new_text[j1:j2]
                if text:
                    result.append((text, 'equal'))
            elif tag == 'delete':
                text = old_text[i1:i2]
                if text:
                    result.append((text, 'delete'))
            elif tag == 'insert':
                text = new_text[j1:j2]
                if text:
                    result.append((text, 'insert'))
            elif tag == 'replace':
                del_text = old_text[i1:i2]
                ins_text = new_text[j1:j2]
                if del_text:
                    result.append((del_text, 'delete'))
                if ins_text:
                    result.append((ins_text, 'insert'))

        return self._merge_adjacent_segments(result)

    def _char_level_diff(self, old_text: str, new_text: str) -> List[Tuple[str, str]]:
        """
        Do character-level diff between two strings.

        Returns list of (text, type) tuples showing only actual character changes.
        """
        matcher = difflib.SequenceMatcher(None, old_text, new_text, autojunk=False)
        result = []

        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == 'equal':
                text = old_text[i1:i2]
                if text:
                    result.append((text, 'equal'))
            elif tag == 'delete':
                text = old_text[i1:i2]
                if text:
                    result.append((text, 'delete'))
            elif tag == 'insert':
                text = new_text[j1:j2]
                if text:
                    result.append((text, 'insert'))
            elif tag == 'replace':
                del_text = old_text[i1:i2]
                ins_text = new_text[j1:j2]
                if del_text:
                    result.append((del_text, 'delete'))
                if ins_text:
                    result.append((ins_text, 'insert'))

        # Merge adjacent segments of the same type for cleaner output
        return self._merge_adjacent_segments(result)

    def _merge_adjacent_segments(self, segments: List[Tuple[str, str]]) -> List[Tuple[str, str]]:
        """Merge adjacent segments of the same type."""
        if not segments:
            return segments

        merged = [segments[0]]
        for text, seg_type in segments[1:]:
            if seg_type == merged[-1][1]:
                # Same type - merge
                merged[-1] = (merged[-1][0] + text, seg_type)
            else:
                merged.append((text, seg_type))

        return merged

    def _tokenize(self, text: str) -> List[str]:
        """
        Split text into words, separating punctuation from words.

        This ensures "word" and "word," are tokenized as ["word"] and ["word", ","]
        so the diff can show just the punctuation change, not the whole word.
        """
        tokens = []
        # Match: leading punctuation, word characters, trailing punctuation, or whitespace
        for match in re.finditer(r'(\s+)|([^\w\s]*)([\w]+)([^\w\s]*)', text):
            whitespace, leading_punct, word, trailing_punct = match.groups()
            if whitespace:
                tokens.append(whitespace)
            else:
                if leading_punct:
                    tokens.append(leading_punct)
                if word:
                    tokens.append(word)
                if trailing_punct:
                    tokens.append(trailing_punct)
        return tokens

    def _detect_moves(self, diff_segments: List[Tuple[str, str]]) -> List[Tuple[str, str]]:
        """Detect moves within diff segments."""
        # Collect deletions and insertions
        deletions = []
        insertions = []

        for i, (text, seg_type) in enumerate(diff_segments):
            words = len(text.split())
            if seg_type == 'delete' and words >= MIN_MOVE_WORDS:
                deletions.append((i, text, self._normalize_for_move(text)))
            elif seg_type == 'insert' and words >= MIN_MOVE_WORDS:
                insertions.append((i, text, self._normalize_for_move(text)))

        if not deletions or not insertions:
            return diff_segments

        # Find matching moves
        moves = {}
        used_insertions = set()

        for del_idx, del_text, del_norm in sorted(deletions, key=lambda x: len(x[1].split()), reverse=True):
            best_match = None
            best_sim = 0

            for ins_idx, ins_text, ins_norm in insertions:
                if ins_idx in used_insertions:
                    continue

                sim = self._calculate_similarity(del_norm, ins_norm)
                if sim >= MOVE_SIMILARITY_THRESHOLD and sim > best_sim:
                    best_sim = sim
                    best_match = ins_idx

            if best_match is not None:
                moves[del_idx] = best_match
                used_insertions.add(best_match)

        if not moves:
            return diff_segments

        # Create new segments with move markers
        result = []
        for i, (text, seg_type) in enumerate(diff_segments):
            if i in moves:
                result.append((text, 'move_source'))
            elif i in used_insertions:
                result.append((text, 'move_dest'))
            else:
                result.append((text, seg_type))

        return result

    def _normalize_for_move(self, text: str) -> str:
        """Normalize text for move detection comparison."""
        text = text.lower().strip()
        text = re.sub(r'\s+', ' ', text)
        return text

    def _rebuild_para_with_redline(self, para_elem, diff_segments: List[Tuple[str, str]]):
        """Rebuild paragraph with redline runs, preserving paragraph properties."""
        # Find and preserve paragraph properties (pPr)
        pPr = para_elem.find('w:pPr', NSMAP)

        # Remove all existing runs and text elements
        for child in list(para_elem):
            if child.tag == f'{W}r' or child.tag == f'{W}bookmarkStart' or child.tag == f'{W}bookmarkEnd':
                para_elem.remove(child)

        # Add new runs for each diff segment
        for text, seg_type in diff_segments:
            if not text:
                continue

            run = self._create_run(text, seg_type)
            para_elem.append(run)

            # Update stats
            word_count = len(text.split())
            if seg_type == 'delete':
                self.stats['deletions'] += word_count
            elif seg_type == 'insert':
                self.stats['insertions'] += word_count
            elif seg_type == 'move_source':
                self.stats['move_from'] += word_count
                self.stats['moves'] += word_count
            elif seg_type == 'move_dest':
                self.stats['move_to'] += word_count
            else:  # equal
                self.stats['unchanged'] += word_count

    def _create_run(self, text: str, seg_type: str):
        """Create a run element with appropriate formatting."""
        run = etree.Element(f'{W}r')

        # Create run properties
        rPr = etree.SubElement(run, f'{W}rPr')

        if seg_type == 'delete':
            # Red strikethrough
            etree.SubElement(rPr, f'{W}strike')
            color = etree.SubElement(rPr, f'{W}color')
            color.set(f'{W}val', 'FF0000')

        elif seg_type == 'insert':
            # Blue bold
            etree.SubElement(rPr, f'{W}b')
            color = etree.SubElement(rPr, f'{W}color')
            color.set(f'{W}val', '0000FF')

        elif seg_type == 'move_source':
            # Green strikethrough
            etree.SubElement(rPr, f'{W}strike')
            color = etree.SubElement(rPr, f'{W}color')
            color.set(f'{W}val', '008000')

        elif seg_type == 'move_dest':
            # Green underline
            u = etree.SubElement(rPr, f'{W}u')
            u.set(f'{W}val', 'single')
            color = etree.SubElement(rPr, f'{W}color')
            color.set(f'{W}val', '008000')

        # If no special formatting needed, remove empty rPr
        if len(rPr) == 0:
            run.remove(rPr)

        # Create text element
        t = etree.SubElement(run, f'{W}t')
        t.text = text
        # Preserve whitespace
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

        return run

    def _mark_para_as_inserted(self, para_elem):
        """Mark an entire paragraph as inserted (blue bold)."""
        # Apply insertion formatting to all runs
        for run in para_elem.findall('.//w:r', NSMAP):
            rPr = run.find('w:rPr', NSMAP)
            if rPr is None:
                rPr = etree.Element(f'{W}rPr')
                run.insert(0, rPr)

            # Add bold
            if rPr.find('w:b', NSMAP) is None:
                etree.SubElement(rPr, f'{W}b')

            # Add blue color
            color = rPr.find('w:color', NSMAP)
            if color is None:
                color = etree.SubElement(rPr, f'{W}color')
            color.set(f'{W}val', '0000FF')

    def _insert_deleted_paragraphs(self, body, mod_para_elems: List, orig_paras: List[str],
                                   alignments: Dict, matched_orig: set):
        """Insert deleted paragraphs at appropriate positions."""
        # Find paragraphs in original that weren't matched
        deleted_indices = []
        for i in range(len(orig_paras)):
            if i not in matched_orig and orig_paras[i].strip():
                deleted_indices.append(i)

        if not deleted_indices:
            return

        # For each deleted paragraph, find the best insertion position
        # Insert before the first modified paragraph that came after it in original

        # Build mapping: orig_idx -> mod_idx for matched paragraphs
        orig_to_mod = {}
        for mod_idx, (orig_idx, align_type) in alignments.items():
            if align_type == 'match' and orig_idx is not None:
                orig_to_mod[orig_idx] = mod_idx

        # Track insertions to adjust positions
        inserted_count = 0

        for del_orig_idx in deleted_indices:
            del_text = orig_paras[del_orig_idx]

            # Find the next matched paragraph after this deleted one
            insert_before_mod_idx = None
            for orig_idx in range(del_orig_idx + 1, len(orig_paras)):
                if orig_idx in orig_to_mod:
                    insert_before_mod_idx = orig_to_mod[orig_idx]
                    break

            # Create deleted paragraph element
            del_para = self._create_deleted_paragraph(del_text)

            # Insert at the right position
            if insert_before_mod_idx is not None:
                # Adjust for previous insertions
                actual_idx = insert_before_mod_idx + inserted_count
                if actual_idx < len(list(body)):
                    body.insert(list(body).index(mod_para_elems[insert_before_mod_idx]) + inserted_count, del_para)
                else:
                    body.append(del_para)
            else:
                # No following paragraph found, append at end
                # But before sectPr if it exists
                sectPr = body.find('w:sectPr', NSMAP)
                if sectPr is not None:
                    body.insert(list(body).index(sectPr), del_para)
                else:
                    body.append(del_para)

            inserted_count += 1
            self.stats['deletions'] += len(del_text.split())

    def _create_deleted_paragraph(self, text: str):
        """Create a paragraph element with deletion formatting (red strikethrough)."""
        para = etree.Element(f'{W}p')

        # Create run with deletion formatting
        run = self._create_run(text, 'delete')
        para.append(run)

        return para


def generate_xml_redline(modified_path: str, original_path: str, output_path: str) -> Dict:
    """
    Convenience function to generate a redlined document.

    Args:
        modified_path: Path to the modified document
        original_path: Path to the original document
        output_path: Path for the output redlined document

    Returns:
        Statistics dictionary
    """
    generator = XmlRedlineGenerator(modified_path, original_path, output_path)
    return generator.generate()


if __name__ == '__main__':
    import sys

    if len(sys.argv) != 4:
        print("Usage: python xml_redline_generator.py <original.docx> <modified.docx> <output.docx>")
        sys.exit(1)

    original = sys.argv[1]
    modified = sys.argv[2]
    output = sys.argv[3]

    print(f"Generating redline...")
    print(f"  Original: {original}")
    print(f"  Modified: {modified}")
    print(f"  Output: {output}")

    stats = generate_xml_redline(modified, original, output)

    print(f"\nStatistics:")
    print(f"  Insertions: {stats['insertions']} words")
    print(f"  Deletions: {stats['deletions']} words")
    print(f"  Moves: {stats['moves']} words")
    print(f"  Unchanged: {stats['unchanged']} words")
    print(f"\nDone! Output saved to: {output}")
