namespace DocumentCompare.Core.Models;

/// <summary>
/// Base class for block-level document elements (paragraphs, tables, etc.).
/// </summary>
public abstract class Block
{
    /// <summary>
    /// Unique identifier for tracking this block through comparison.
    /// </summary>
    public string? Id { get; set; }

    /// <summary>
    /// Creates a deep copy of this block.
    /// </summary>
    public abstract Block Clone();

    /// <summary>
    /// Gets the plain text content of this block for comparison purposes.
    /// </summary>
    public abstract string GetPlainText();
}

/// <summary>
/// Represents a paragraph containing formatted text runs.
/// </summary>
public class Paragraph : Block
{
    /// <summary>
    /// The text runs that make up this paragraph.
    /// </summary>
    public List<Run> Runs { get; set; } = new();

    /// <summary>
    /// Paragraph-level formatting.
    /// </summary>
    public ParagraphStyle Style { get; set; } = new();

    /// <summary>
    /// Numbering information if this paragraph is part of a numbered list.
    /// </summary>
    public NumberingInfo? Numbering { get; set; }

    /// <summary>
    /// Bookmark IDs that start at this paragraph.
    /// </summary>
    public List<string> BookmarkStarts { get; set; } = new();

    /// <summary>
    /// Bookmark IDs that end at this paragraph.
    /// </summary>
    public List<string> BookmarkEnds { get; set; } = new();

    public override Block Clone()
    {
        return new Paragraph
        {
            Id = Id,
            Runs = Runs.Select(r => r.Clone()).ToList(),
            Style = Style.Clone(),
            Numbering = Numbering?.Clone(),
            BookmarkStarts = new List<string>(BookmarkStarts),
            BookmarkEnds = new List<string>(BookmarkEnds)
        };
    }

    public override string GetPlainText()
    {
        return string.Concat(Runs.Select(r => r.Text));
    }

    /// <summary>
    /// Gets the text content normalized for comparison (trimmed, normalized whitespace).
    /// </summary>
    public string GetNormalizedText()
    {
        var text = GetPlainText();
        // Normalize whitespace: collapse multiple spaces, trim
        return System.Text.RegularExpressions.Regex.Replace(text.Trim(), @"\s+", " ");
    }
}

/// <summary>
/// Represents a table in the document.
/// </summary>
public class Table : Block
{
    public List<TableRow> Rows { get; set; } = new();

    /// <summary>
    /// Table-wide properties like width, borders, etc.
    /// </summary>
    public TableProperties? Properties { get; set; }

    public override Block Clone()
    {
        return new Table
        {
            Id = Id,
            Rows = Rows.Select(r => r.Clone()).ToList(),
            Properties = Properties?.Clone()
        };
    }

    public override string GetPlainText()
    {
        return string.Join("\n", Rows.Select(r => r.GetPlainText()));
    }
}

public class TableRow
{
    public List<TableCell> Cells { get; set; } = new();
    public TableRowProperties? Properties { get; set; }

    public TableRow Clone()
    {
        return new TableRow
        {
            Cells = Cells.Select(c => c.Clone()).ToList(),
            Properties = Properties?.Clone()
        };
    }

    public string GetPlainText()
    {
        return string.Join("\t", Cells.Select(c => c.GetPlainText()));
    }
}

public class TableCell
{
    public List<Block> Blocks { get; set; } = new();
    public TableCellProperties? Properties { get; set; }

    public TableCell Clone()
    {
        return new TableCell
        {
            Blocks = Blocks.Select(b => b.Clone()).ToList(),
            Properties = Properties?.Clone()
        };
    }

    public string GetPlainText()
    {
        return string.Join("\n", Blocks.Select(b => b.GetPlainText()));
    }
}

public class TableProperties
{
    public int? Width { get; set; }
    public string? WidthType { get; set; } // auto, dxa (twips), pct
    public string? Alignment { get; set; }

    public TableProperties Clone() => new()
    {
        Width = Width,
        WidthType = WidthType,
        Alignment = Alignment
    };
}

public class TableRowProperties
{
    public int? Height { get; set; }
    public bool IsHeader { get; set; }

    public TableRowProperties Clone() => new()
    {
        Height = Height,
        IsHeader = IsHeader
    };
}

public class TableCellProperties
{
    public int? Width { get; set; }
    public int? ColumnSpan { get; set; }
    public int? RowSpan { get; set; }
    public string? VerticalAlignment { get; set; }

    public TableCellProperties Clone() => new()
    {
        Width = Width,
        ColumnSpan = ColumnSpan,
        RowSpan = RowSpan,
        VerticalAlignment = VerticalAlignment
    };
}
