namespace DocumentCompare.Core.Models;

/// <summary>
/// Represents character-level formatting for a run of text.
/// </summary>
public class RunFormatting
{
    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public bool Underline { get; set; }
    public bool Strikethrough { get; set; }
    public string? FontFamily { get; set; }
    public double? FontSize { get; set; }
    public string? Color { get; set; }
    public string? HighlightColor { get; set; }
    public bool Superscript { get; set; }
    public bool Subscript { get; set; }

    /// <summary>
    /// Style name this run inherits from (e.g., "Heading 1 Char")
    /// </summary>
    public string? StyleId { get; set; }

    /// <summary>
    /// Creates a deep copy of this formatting.
    /// </summary>
    public RunFormatting Clone()
    {
        return new RunFormatting
        {
            Bold = Bold,
            Italic = Italic,
            Underline = Underline,
            Strikethrough = Strikethrough,
            FontFamily = FontFamily,
            FontSize = FontSize,
            Color = Color,
            HighlightColor = HighlightColor,
            Superscript = Superscript,
            Subscript = Subscript,
            StyleId = StyleId
        };
    }

    /// <summary>
    /// Creates formatting for deleted text (red strikethrough).
    /// </summary>
    public static RunFormatting ForDeletion(RunFormatting? original = null)
    {
        var formatting = original?.Clone() ?? new RunFormatting();
        formatting.Strikethrough = true;
        formatting.Color = "FF0000"; // Red
        return formatting;
    }

    /// <summary>
    /// Creates formatting for inserted text (bold blue).
    /// </summary>
    public static RunFormatting ForInsertion(RunFormatting? original = null)
    {
        var formatting = original?.Clone() ?? new RunFormatting();
        formatting.Bold = true;
        formatting.Color = "0000FF"; // Blue
        return formatting;
    }

    /// <summary>
    /// Creates formatting for moved text (green).
    /// </summary>
    public static RunFormatting ForMove(RunFormatting? original = null, bool isSource = false)
    {
        var formatting = original?.Clone() ?? new RunFormatting();
        formatting.Color = "008000"; // Green
        if (isSource)
        {
            formatting.Strikethrough = true;
        }
        return formatting;
    }
}
