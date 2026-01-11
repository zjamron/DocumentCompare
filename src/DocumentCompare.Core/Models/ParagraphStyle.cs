namespace DocumentCompare.Core.Models;

/// <summary>
/// Represents paragraph-level formatting and style information.
/// </summary>
public class ParagraphStyle
{
    /// <summary>
    /// The style ID this paragraph uses (e.g., "Heading1", "Normal").
    /// </summary>
    public string? StyleId { get; set; }

    /// <summary>
    /// Heading level (1-9) if this is a heading, null otherwise.
    /// </summary>
    public int? HeadingLevel { get; set; }

    /// <summary>
    /// Paragraph alignment: left, center, right, justify
    /// </summary>
    public string Alignment { get; set; } = "left";

    /// <summary>
    /// Left indent in twips.
    /// </summary>
    public int? LeftIndent { get; set; }

    /// <summary>
    /// Right indent in twips.
    /// </summary>
    public int? RightIndent { get; set; }

    /// <summary>
    /// First line indent in twips (positive for indent, negative for hanging).
    /// </summary>
    public int? FirstLineIndent { get; set; }

    /// <summary>
    /// Space before paragraph in twips.
    /// </summary>
    public int? SpaceBefore { get; set; }

    /// <summary>
    /// Space after paragraph in twips.
    /// </summary>
    public int? SpaceAfter { get; set; }

    /// <summary>
    /// Line spacing value.
    /// </summary>
    public double? LineSpacing { get; set; }

    /// <summary>
    /// Line spacing rule: auto, exact, atLeast
    /// </summary>
    public string? LineSpacingRule { get; set; }

    /// <summary>
    /// Keep with next paragraph (don't page break between).
    /// </summary>
    public bool KeepWithNext { get; set; }

    /// <summary>
    /// Keep lines together (don't page break within).
    /// </summary>
    public bool KeepLinesTogether { get; set; }

    /// <summary>
    /// Page break before this paragraph.
    /// </summary>
    public bool PageBreakBefore { get; set; }

    /// <summary>
    /// Outline level for TOC generation (0-8).
    /// </summary>
    public int? OutlineLevel { get; set; }

    public ParagraphStyle Clone()
    {
        return new ParagraphStyle
        {
            StyleId = StyleId,
            HeadingLevel = HeadingLevel,
            Alignment = Alignment,
            LeftIndent = LeftIndent,
            RightIndent = RightIndent,
            FirstLineIndent = FirstLineIndent,
            SpaceBefore = SpaceBefore,
            SpaceAfter = SpaceAfter,
            LineSpacing = LineSpacing,
            LineSpacingRule = LineSpacingRule,
            KeepWithNext = KeepWithNext,
            KeepLinesTogether = KeepLinesTogether,
            PageBreakBefore = PageBreakBefore,
            OutlineLevel = OutlineLevel
        };
    }
}
