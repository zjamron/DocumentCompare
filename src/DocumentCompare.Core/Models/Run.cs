namespace DocumentCompare.Core.Models;

/// <summary>
/// Represents a contiguous run of text with consistent formatting.
/// </summary>
public class Run
{
    public string Text { get; set; } = string.Empty;
    public RunFormatting Formatting { get; set; } = new();

    public Run() { }

    public Run(string text, RunFormatting? formatting = null)
    {
        Text = text;
        Formatting = formatting ?? new RunFormatting();
    }

    /// <summary>
    /// Creates a deep copy of this run.
    /// </summary>
    public Run Clone()
    {
        return new Run
        {
            Text = Text,
            Formatting = Formatting.Clone()
        };
    }
}
