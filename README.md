# DocumentCompare

A document comparison (redlining) library that compares two documents and generates a redlined output showing changes. Similar to Litera Compare / DeltaView.

## Features

- **Word document support** (.docx) - Parse and generate Word documents
- **Formatting preservation** - Maintains original document formatting, styles, and section numbering
- **Visual redlining**:
  - Deletions shown in **red with strikethrough**
  - Insertions shown in **blue bold**
- **Word-level diff** - Precise change detection at the word level
- **Paragraph alignment** - Intelligently matches paragraphs between documents

## Quick Start (Python)

The fastest way to use this tool - no build required:

```python
import sys
sys.path.insert(0, r"C:\Users\zachary\DocumentCompare\samples")
from test_comparison import create_redlined_document

stats = create_redlined_document(
    original_path="contract_v1.docx",
    modified_path="contract_v2.docx",
    output_path="contract_redline.docx"
)

print(f"Insertions: {stats['insertions']}, Deletions: {stats['deletions']}")
```

Or run from command line:

```bash
cd samples
python create_test_docs.py   # Create sample documents
python test_comparison.py    # Run comparison
```

## C# Library Usage

Requires [.NET 8 SDK](https://dotnet.microsoft.com/download).

```bash
dotnet build
dotnet test
```

```csharp
using DocumentCompare.Word;
using DocumentCompare.Core.Interfaces;

var comparer = WordDocumentComparer.Create();

var result = comparer.Compare(new CompareRequest
{
    OriginalDocumentPath = "original.docx",
    ModifiedDocumentPath = "modified.docx",
    OutputPath = "redline.docx",
    OutputFormat = OutputFormat.Word
});

if (result.Success)
{
    Console.WriteLine($"Created: {result.OutputPath}");
    Console.WriteLine($"Insertions: {result.Statistics.Insertions}");
    Console.WriteLine($"Deletions: {result.Statistics.Deletions}");
}
```

## Project Structure

```
DocumentCompare/
├── src/
│   ├── DocumentCompare.Core/     # Core models and diff engine
│   ├── DocumentCompare.Word/     # Word document support (OpenXML)
│   └── DocumentCompare.Pdf/      # PDF support (planned)
├── tests/                        # Unit tests
└── samples/                      # Sample documents and Python scripts
```

## Output Example

| Change Type | Formatting |
|-------------|------------|
| Deleted text | Red + Strikethrough |
| Inserted text | Blue + Bold |
| Moved text | Green (Phase 3) |

## Roadmap

- [x] **Phase 1**: Word-to-Word comparison with formatting preservation
- [ ] **Phase 2**: PDF support (input and output)
- [ ] **Phase 3**: Move detection (green highlighting)
- [ ] **Phase 4**: Table cell-level comparison

## Dependencies

**Python** (for quick usage):
- python-docx

**C#/.NET** (for library):
- DocumentFormat.OpenXml 3.0.0
- DiffPlex 1.7.1
- PdfPig 0.1.8 (Phase 2)
- QuestPDF 2024.3.0 (Phase 2)

## License

MIT
