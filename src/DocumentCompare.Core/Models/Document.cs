namespace DocumentCompare.Core.Models;

/// <summary>
/// Represents a complete document with all its content, formatting, and metadata.
/// </summary>
public class Document
{
    /// <summary>
    /// The sections that make up this document.
    /// </summary>
    public List<Section> Sections { get; set; } = new();

    /// <summary>
    /// Document-level properties and metadata.
    /// </summary>
    public DocumentProperties Properties { get; set; } = new();

    /// <summary>
    /// All numbering definitions (abstract numberings) used in the document.
    /// </summary>
    public List<NumberingDefinition> NumberingDefinitions { get; set; } = new();

    /// <summary>
    /// Numbering instances that reference the definitions.
    /// </summary>
    public List<NumberingInstance> NumberingInstances { get; set; } = new();

    /// <summary>
    /// Style definitions used in the document.
    /// </summary>
    public List<StyleDefinition> Styles { get; set; } = new();

    /// <summary>
    /// Gets all paragraphs in the document in order.
    /// </summary>
    public IEnumerable<Paragraph> GetAllParagraphs()
    {
        foreach (var section in Sections)
        {
            foreach (var block in section.Blocks)
            {
                if (block is Paragraph paragraph)
                {
                    yield return paragraph;
                }
                else if (block is Table table)
                {
                    foreach (var row in table.Rows)
                    {
                        foreach (var cell in row.Cells)
                        {
                            foreach (var cellBlock in cell.Blocks)
                            {
                                if (cellBlock is Paragraph cellParagraph)
                                {
                                    yield return cellParagraph;
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    /// <summary>
    /// Gets the plain text content of the entire document.
    /// </summary>
    public string GetPlainText()
    {
        var paragraphs = GetAllParagraphs().Select(p => p.GetPlainText());
        return string.Join("\n", paragraphs);
    }

    /// <summary>
    /// Creates a deep copy of this document.
    /// </summary>
    public Document Clone()
    {
        return new Document
        {
            Sections = Sections.Select(s => s.Clone()).ToList(),
            Properties = Properties.Clone(),
            NumberingDefinitions = NumberingDefinitions.Select(n => n.Clone()).ToList(),
            NumberingInstances = NumberingInstances.Select(n => new NumberingInstance
            {
                Id = n.Id,
                DefinitionId = n.DefinitionId,
                LevelOverrides = new Dictionary<int, NumberingLevelOverride>(n.LevelOverrides)
            }).ToList(),
            Styles = Styles.Select(s => s.Clone()).ToList()
        };
    }
}

/// <summary>
/// Document-level properties and metadata.
/// </summary>
public class DocumentProperties
{
    public string? Title { get; set; }
    public string? Author { get; set; }
    public string? Subject { get; set; }
    public string? Description { get; set; }
    public string? Keywords { get; set; }
    public DateTime? Created { get; set; }
    public DateTime? Modified { get; set; }
    public string? Creator { get; set; }
    public string? LastModifiedBy { get; set; }

    /// <summary>
    /// Default font for the document.
    /// </summary>
    public string? DefaultFont { get; set; }

    /// <summary>
    /// Default font size in points.
    /// </summary>
    public double? DefaultFontSize { get; set; }

    public DocumentProperties Clone()
    {
        return new DocumentProperties
        {
            Title = Title,
            Author = Author,
            Subject = Subject,
            Description = Description,
            Keywords = Keywords,
            Created = Created,
            Modified = Modified,
            Creator = Creator,
            LastModifiedBy = LastModifiedBy,
            DefaultFont = DefaultFont,
            DefaultFontSize = DefaultFontSize
        };
    }
}

/// <summary>
/// Represents a style definition (paragraph or character style).
/// </summary>
public class StyleDefinition
{
    public string Id { get; set; } = string.Empty;
    public string? Name { get; set; }
    public StyleType Type { get; set; }
    public string? BasedOn { get; set; }
    public string? NextStyle { get; set; }
    public ParagraphStyle? ParagraphProperties { get; set; }
    public RunFormatting? RunProperties { get; set; }

    public StyleDefinition Clone()
    {
        return new StyleDefinition
        {
            Id = Id,
            Name = Name,
            Type = Type,
            BasedOn = BasedOn,
            NextStyle = NextStyle,
            ParagraphProperties = ParagraphProperties?.Clone(),
            RunProperties = RunProperties?.Clone()
        };
    }
}

public enum StyleType
{
    Paragraph,
    Character,
    Table,
    Numbering
}
