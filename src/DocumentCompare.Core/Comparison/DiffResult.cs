namespace DocumentCompare.Core.Comparison;

/// <summary>
/// Represents the type of change in a diff operation.
/// </summary>
public enum DiffType
{
    /// <summary>
    /// Content is unchanged.
    /// </summary>
    Unchanged,

    /// <summary>
    /// Content was inserted (exists in modified but not original).
    /// </summary>
    Inserted,

    /// <summary>
    /// Content was deleted (exists in original but not modified).
    /// </summary>
    Deleted,

    /// <summary>
    /// Content was moved from another location.
    /// </summary>
    MovedFrom,

    /// <summary>
    /// Content was moved to another location.
    /// </summary>
    MovedTo
}

/// <summary>
/// Represents a segment of text with its diff status.
/// </summary>
public class DiffSegment
{
    /// <summary>
    /// The text content.
    /// </summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>
    /// The type of change.
    /// </summary>
    public DiffType Type { get; set; }

    /// <summary>
    /// Index in the original document (for deleted/unchanged segments).
    /// </summary>
    public int? OriginalIndex { get; set; }

    /// <summary>
    /// Index in the modified document (for inserted/unchanged segments).
    /// </summary>
    public int? ModifiedIndex { get; set; }

    /// <summary>
    /// For moved segments, the ID linking source and destination.
    /// </summary>
    public string? MoveId { get; set; }
}

/// <summary>
/// Result of comparing two paragraphs.
/// </summary>
public class ParagraphDiffResult
{
    /// <summary>
    /// The diff segments making up the merged content.
    /// </summary>
    public List<DiffSegment> Segments { get; set; } = new();

    /// <summary>
    /// Whether the paragraph was entirely deleted.
    /// </summary>
    public bool IsEntirelyDeleted { get; set; }

    /// <summary>
    /// Whether the paragraph was entirely inserted.
    /// </summary>
    public bool IsEntirelyInserted { get; set; }

    /// <summary>
    /// Count of inserted words.
    /// </summary>
    public int InsertionCount => Segments.Count(s => s.Type == DiffType.Inserted);

    /// <summary>
    /// Count of deleted words.
    /// </summary>
    public int DeletionCount => Segments.Count(s => s.Type == DiffType.Deleted);

    /// <summary>
    /// Count of unchanged words.
    /// </summary>
    public int UnchangedCount => Segments.Count(s => s.Type == DiffType.Unchanged);
}

/// <summary>
/// Result of aligning paragraphs between two documents.
/// </summary>
public class ParagraphAlignment
{
    /// <summary>
    /// Index in the original document (-1 if inserted).
    /// </summary>
    public int OriginalIndex { get; set; } = -1;

    /// <summary>
    /// Index in the modified document (-1 if deleted).
    /// </summary>
    public int ModifiedIndex { get; set; } = -1;

    /// <summary>
    /// The type of change for this paragraph.
    /// </summary>
    public DiffType Type { get; set; }

    /// <summary>
    /// Similarity score (0-1) for matched paragraphs.
    /// </summary>
    public double SimilarityScore { get; set; }
}
