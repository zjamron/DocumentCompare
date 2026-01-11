# DocumentCompare

A document comparison (redlining) library that compares two documents and generates a redlined output showing changes. Similar to Litera Compare / DeltaView.

## Features

- **Word document support** (.docx) - Parse and generate Word documents
- **PDF document support** (.pdf) - Parse PDFs and generate redlined PDF output
- **Cross-format comparison** - Compare Word to PDF or PDF to Word
- **Formatting preservation** - Maintains original document formatting, styles, and section numbering
- **Visual redlining**:
  - Deletions shown in **red with strikethrough**
  - Insertions shown in **blue bold**
  - Moved text shown in **green** (strikethrough at source, plain at destination)
- **Word-level diff** - Precise change detection at the word level
- **Paragraph alignment** - Intelligently matches paragraphs between documents
- **Header/footer comparison** - Compares headers and footers across document sections
- **Move detection** - Detects text that was moved (not just deleted and re-inserted)

## Quick Start (Python)

### Universal Document Comparison (Word & PDF)

```bash
cd samples
pip install python-docx pymupdf reportlab

# Compare any combination of Word and PDF files
python document_compare.py original.docx modified.docx redline.docx
python document_compare.py original.pdf modified.pdf redline.pdf
python document_compare.py original.docx modified.pdf redline.pdf
```

### Word-to-Word with Formatting Preservation

For Word documents where you need to preserve complex formatting:

```bash
python compare_preserve_formatting.py original.docx modified.docx redline.docx
```

### Python API

```python
from document_compare import compare_documents

result = compare_documents(
    original_path="contract_v1.docx",
    modified_path="contract_v2.pdf",
    output_path="contract_redline.pdf"
)

print(f"Insertions: {result.insertions}, Deletions: {result.deletions}, Moves: {result.moves}")
```

## Supported Formats

| Input Format | Output Format | Script |
|--------------|---------------|--------|
| Word (.docx) | Word (.docx) | `compare_preserve_formatting.py` (best for complex docs) |
| Word (.docx) | Word (.docx) | `document_compare.py` |
| Word (.docx) | PDF (.pdf) | `document_compare.py` |
| PDF (.pdf) | PDF (.pdf) | `document_compare.py` |
| PDF (.pdf) | Word (.docx) | `document_compare.py` |
| Word + PDF | Either | `document_compare.py` |

## Project Structure

```
DocumentCompare/
├── src/
│   ├── DocumentCompare.Core/     # Core models and diff engine (C#)
│   ├── DocumentCompare.Word/     # Word document support (OpenXML)
│   └── DocumentCompare.Pdf/      # PDF support (C#)
├── tests/                        # Unit tests
└── samples/                      # Python scripts
    ├── document_compare.py       # Universal comparison (Word + PDF)
    ├── compare_preserve_formatting.py  # Word with formatting preservation
    ├── pdf_support.py            # PDF parsing and generation
    ├── test_comparison.py        # Basic Word comparison
    └── create_test_docs.py       # Generate test documents
```

## Output Example

| Change Type | Formatting |
|-------------|------------|
| Deleted text | Red + Strikethrough |
| Inserted text | Blue + Bold |
| Moved text (source) | Green + Strikethrough |
| Moved text (destination) | Green |

## Roadmap

- [x] **Phase 1**: Word-to-Word comparison with formatting preservation
- [x] **Phase 2**: PDF support (input and output)
- [x] **Phase 3**: Move detection (green highlighting)
- [ ] **Phase 4**: Table cell-level comparison

## Dependencies

**Python** (for quick usage):
- python-docx
- pymupdf (PyMuPDF)
- reportlab

**C#/.NET** (for library):
- DocumentFormat.OpenXml 3.0.0
- DiffPlex 1.7.1
- PdfPig 0.1.8
- QuestPDF 2024.3.0

## Installation

```bash
git clone https://github.com/zjamron/DocumentCompare.git
cd DocumentCompare/samples
pip install python-docx pymupdf reportlab
```

## License

MIT
