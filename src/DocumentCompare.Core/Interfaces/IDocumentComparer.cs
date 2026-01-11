using DocumentCompare.Core.Models;

namespace DocumentCompare.Core.Interfaces;

/// <summary>
/// Main interface for comparing documents and generating redlined output.
/// </summary>
public interface IDocumentComparer
{
    /// <summary>
    /// Compares two documents and generates a redlined output.
    /// </summary>
    CompareResult Compare(CompareRequest request);

    /// <summary>
    /// Compares two documents and returns a redlined document model.
    /// </summary>
    Document CompareToDocument(Document original, Document modified, CompareOptions? options = null);
}

/// <summary>
/// Request for document comparison.
/// </summary>
public class CompareRequest
{
    /// <summary>
    /// Path to the original (baseline) document.
    /// </summary>
    public string OriginalDocumentPath { get; set; } = string.Empty;

    /// <summary>
    /// Path to the modified (new) document.
    /// </summary>
    public string ModifiedDocumentPath { get; set; } = string.Empty;

    /// <summary>
    /// Optional: Stream for the original document.
    /// </summary>
    public Stream? OriginalStream { get; set; }

    /// <summary>
    /// Optional: Stream for the modified document.
    /// </summary>
    public Stream? ModifiedStream { get; set; }

    /// <summary>
    /// Desired output format.
    /// </summary>
    public OutputFormat OutputFormat { get; set; } = OutputFormat.Word;

    /// <summary>
    /// Output file path (required if OutputFormat is not Stream).
    /// </summary>
    public string? OutputPath { get; set; }

    /// <summary>
    /// Comparison options.
    /// </summary>
    public CompareOptions Options { get; set; } = new();
}

/// <summary>
/// Options for document comparison.
/// </summary>
public class CompareOptions
{
    /// <summary>
    /// Whether to attempt move detection (green highlighting).
    /// </summary>
    public bool DetectMoves { get; set; } = false;

    /// <summary>
    /// Ignore differences in whitespace.
    /// </summary>
    public bool IgnoreWhitespace { get; set; } = true;

    /// <summary>
    /// Ignore case differences.
    /// </summary>
    public bool IgnoreCase { get; set; } = false;

    /// <summary>
    /// Ignore formatting differences (only compare text).
    /// </summary>
    public bool IgnoreFormatting { get; set; } = false;

    /// <summary>
    /// Custom styles for redline markup.
    /// </summary>
    public RedlineStyles Styles { get; set; } = new();

    /// <summary>
    /// Granularity of comparison.
    /// </summary>
    public CompareGranularity Granularity { get; set; } = CompareGranularity.Word;
}

/// <summary>
/// Custom styles for redline formatting.
/// </summary>
public class RedlineStyles
{
    /// <summary>
    /// Color for deleted text (hex without #, e.g., "FF0000").
    /// </summary>
    public string DeletionColor { get; set; } = "FF0000";

    /// <summary>
    /// Color for inserted text.
    /// </summary>
    public string InsertionColor { get; set; } = "0000FF";

    /// <summary>
    /// Color for moved text.
    /// </summary>
    public string MoveColor { get; set; } = "008000";

    /// <summary>
    /// Whether insertions should be bold.
    /// </summary>
    public bool InsertionBold { get; set; } = true;

    /// <summary>
    /// Whether deletions should have strikethrough.
    /// </summary>
    public bool DeletionStrikethrough { get; set; } = true;
}

/// <summary>
/// Granularity level for comparison.
/// </summary>
public enum CompareGranularity
{
    /// <summary>
    /// Compare at character level (most detailed, slower).
    /// </summary>
    Character,

    /// <summary>
    /// Compare at word level (balanced).
    /// </summary>
    Word,

    /// <summary>
    /// Compare at sentence level.
    /// </summary>
    Sentence,

    /// <summary>
    /// Compare at paragraph level (least detailed, fastest).
    /// </summary>
    Paragraph
}

/// <summary>
/// Output format for the comparison result.
/// </summary>
public enum OutputFormat
{
    Word,
    Pdf,
    Html
}

/// <summary>
/// Result of a document comparison.
/// </summary>
public class CompareResult
{
    /// <summary>
    /// Path to the output file (if file output was requested).
    /// </summary>
    public string? OutputPath { get; set; }

    /// <summary>
    /// The redlined document model.
    /// </summary>
    public Document? RedlinedDocument { get; set; }

    /// <summary>
    /// Statistics about the comparison.
    /// </summary>
    public CompareStatistics Statistics { get; set; } = new();

    /// <summary>
    /// Whether the comparison was successful.
    /// </summary>
    public bool Success { get; set; }

    /// <summary>
    /// Error message if comparison failed.
    /// </summary>
    public string? ErrorMessage { get; set; }
}

/// <summary>
/// Statistics about the changes found during comparison.
/// </summary>
public class CompareStatistics
{
    /// <summary>
    /// Number of words/segments inserted.
    /// </summary>
    public int Insertions { get; set; }

    /// <summary>
    /// Number of words/segments deleted.
    /// </summary>
    public int Deletions { get; set; }

    /// <summary>
    /// Number of words/segments moved (if move detection enabled).
    /// </summary>
    public int Moves { get; set; }

    /// <summary>
    /// Number of paragraphs in original document.
    /// </summary>
    public int OriginalParagraphs { get; set; }

    /// <summary>
    /// Number of paragraphs in modified document.
    /// </summary>
    public int ModifiedParagraphs { get; set; }

    /// <summary>
    /// Number of unchanged words/segments.
    /// </summary>
    public int Unchanged { get; set; }

    /// <summary>
    /// Percentage of content that changed.
    /// </summary>
    public double ChangePercentage =>
        (Insertions + Deletions + Moves) * 100.0 / Math.Max(1, Insertions + Deletions + Moves + Unchanged);
}
