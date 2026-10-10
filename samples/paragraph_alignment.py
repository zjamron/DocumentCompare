"""
Fast paragraph alignment shared by the comparison and the Word redline generator.

The original approach compared every original paragraph with every modified
paragraph using a character-level similarity (O(m*n) comparisons, each itself
quadratic in paragraph length), which took many minutes on a 50-page agreement.

This version produces the same kind of alignment much faster:
  1. Paragraphs that are identical in both versions are matched first using
     difflib on whole paragraphs (near-linear for typical documents).
  2. The fuzzy, similarity-based LCS only runs inside the gaps between those
     anchors, which are small when most of the document is unchanged.
  3. The similarity test short-circuits: the cheap word-overlap score is tried
     first, and difflib's upper-bound estimates rule out pairs that cannot reach
     the threshold before the expensive ratio is computed. The decision for any
     pair is identical to max(jaccard, SequenceMatcher.ratio()) >= threshold.
"""

import difflib
from typing import List, Tuple

PARAGRAPH_MATCH_THRESHOLD = 0.4


def _norm(text: str) -> str:
    return text.strip()


def is_similar(text1: str, text2: str, threshold: float = PARAGRAPH_MATCH_THRESHOLD) -> bool:
    """Same decision as max(word Jaccard, char SequenceMatcher ratio) >= threshold."""
    if text1 == text2:
        return True
    if not text1 or not text2:
        return 0.0 >= threshold

    text1 = text1.strip()
    text2 = text2.strip()
    if not text1 and not text2:
        return True
    if not text1 or not text2:
        return 0.0 >= threshold

    lower1 = text1.lower()
    lower2 = text2.lower()

    words1 = set(lower1.split())
    words2 = set(lower2.split())
    if words1 and words2:
        union = len(words1 | words2)
        if union and len(words1 & words2) / union >= threshold:
            return True

    matcher = difflib.SequenceMatcher(None, lower1, lower2)
    if matcher.real_quick_ratio() < threshold:
        return False
    if matcher.quick_ratio() < threshold:
        return False
    return matcher.ratio() >= threshold


def _fuzzy_align_block(orig: List[str], mod: List[str], i0: int, j0: int,
                       threshold: float) -> List[Tuple[int, int, str]]:
    """LCS alignment (original algorithm) restricted to one gap between anchors."""
    m, n = len(orig), len(mod)
    if m == 0:
        return [(-1, j0 + j, 'insert') for j in range(n)]
    if n == 0:
        return [(i0 + i, -1, 'delete') for i in range(m)]

    similar = [[False] * n for _ in range(m)]
    for i in range(m):
        row = similar[i]
        a = orig[i]
        for j in range(n):
            row[j] = is_similar(a, mod[j], threshold)

    lcs = [[0] * (n + 1) for _ in range(m + 1)]
    for i in range(1, m + 1):
        prev = lcs[i - 1]
        cur = lcs[i]
        sim_row = similar[i - 1]
        for j in range(1, n + 1):
            if sim_row[j - 1]:
                cur[j] = prev[j - 1] + 1
            else:
                cur[j] = prev[j] if prev[j] >= cur[j - 1] else cur[j - 1]

    out = []
    i, j = m, n
    while i > 0 or j > 0:
        if i > 0 and j > 0 and similar[i - 1][j - 1]:
            out.append((i0 + i - 1, j0 + j - 1, 'match'))
            i -= 1
            j -= 1
            continue
        if j > 0 and (i == 0 or lcs[i][j - 1] >= lcs[i - 1][j]):
            out.append((-1, j0 + j - 1, 'insert'))
            j -= 1
        else:
            out.append((i0 + i - 1, -1, 'delete'))
            i -= 1
    out.reverse()
    return out


def align_paragraphs(orig_texts: List[str], mod_texts: List[str],
                     threshold: float = PARAGRAPH_MATCH_THRESHOLD) -> List[Tuple[int, int, str]]:
    """
    Align paragraphs between two documents.

    Returns an ordered list of (orig_idx, mod_idx, type) where type is 'match',
    'insert' (orig_idx == -1) or 'delete' (mod_idx == -1).
    """
    orig_keys = [_norm(t) for t in orig_texts]
    mod_keys = [_norm(t) for t in mod_texts]

    anchor = difflib.SequenceMatcher(None, orig_keys, mod_keys, autojunk=False)
    result: List[Tuple[int, int, str]] = []
    for tag, i1, i2, j1, j2 in anchor.get_opcodes():
        if tag == 'equal':
            for k in range(i2 - i1):
                result.append((i1 + k, j1 + k, 'match'))
        else:
            result.extend(
                _fuzzy_align_block(orig_texts[i1:i2], mod_texts[j1:j2], i1, j1, threshold)
            )
    return result
