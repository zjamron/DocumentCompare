using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentCompare.Core.Interfaces;
using DocumentCompare.Core.Models;
using Document = DocumentCompare.Core.Models.Document;
using Paragraph = DocumentCompare.Core.Models.Paragraph;
using Run = DocumentCompare.Core.Models.Run;
using Table = DocumentCompare.Core.Models.Table;
using TableRow = DocumentCompare.Core.Models.TableRow;
using TableCell = DocumentCompare.Core.Models.TableCell;
using Section = DocumentCompare.Core.Models.Section;

namespace DocumentCompare.Word;

/// <summary>
/// Parses Word documents (.docx) into the internal document model.
/// </summary>
public class WordParser : IDocumentParser
{
    public IEnumerable<string> SupportedExtensions => new[] { ".docx" };

    public bool CanParse(string filePath)
    {
        return Path.GetExtension(filePath).Equals(".docx", StringComparison.OrdinalIgnoreCase);
    }

    public Document Parse(string filePath)
    {
        using var stream = File.OpenRead(filePath);
        return Parse(stream, filePath);
    }

    public Document Parse(Stream stream, string fileName)
    {
        using var wordDoc = WordprocessingDocument.Open(stream, false);
        return ParseDocument(wordDoc);
    }

    private Document ParseDocument(WordprocessingDocument wordDoc)
    {
        var document = new Document();
        var mainPart = wordDoc.MainDocumentPart;

        if (mainPart?.Document?.Body == null)
        {
            return document;
        }

        // Parse document properties
        document.Properties = ParseDocumentProperties(wordDoc);

        // Parse numbering definitions (CRITICAL for section numbering)
        ParseNumbering(mainPart, document);

        // Parse styles
        ParseStyles(mainPart, document);

        // Parse document body into sections
        var sections = ParseBody(mainPart);
        document.Sections.AddRange(sections);

        return document;
    }

    private DocumentProperties ParseDocumentProperties(WordprocessingDocument wordDoc)
    {
        var props = new DocumentProperties();

        var corePart = wordDoc.CoreFilePropertiesPart;
        if (corePart != null)
        {
            // Core properties are in Dublin Core format
            var coreProps = corePart.GetXDocument();
            var ns = coreProps.Root?.GetDefaultNamespace();

            // Extract basic properties from XML
            props.Title = coreProps.Descendants()
                .FirstOrDefault(e => e.Name.LocalName == "title")?.Value;
            props.Author = coreProps.Descendants()
                .FirstOrDefault(e => e.Name.LocalName == "creator")?.Value;
            props.Subject = coreProps.Descendants()
                .FirstOrDefault(e => e.Name.LocalName == "subject")?.Value;
        }

        return props;
    }

    /// <summary>
    /// Parses numbering definitions - CRITICAL for section numbering preservation.
    /// </summary>
    private void ParseNumbering(MainDocumentPart mainPart, Document document)
    {
        var numberingPart = mainPart.NumberingDefinitionsPart;
        if (numberingPart?.Numbering == null) return;

        var numbering = numberingPart.Numbering;

        // Parse abstract numbering definitions
        foreach (var abstractNum in numbering.Elements<AbstractNum>())
        {
            var definition = new NumberingDefinition
            {
                Id = abstractNum.AbstractNumberId?.Value ?? 0,
                MultiLevel = abstractNum.MultiLevelType?.Val?.Value == MultiLevelValues.HybridMultilevel ||
                            abstractNum.MultiLevelType?.Val?.Value == MultiLevelValues.Multilevel
            };

            foreach (var level in abstractNum.Elements<Level>())
            {
                var numLevel = new NumberingLevel
                {
                    Level = level.LevelIndex?.Value ?? 0,
                    Format = level.NumberingFormat?.Val?.ToString() ?? "decimal",
                    Text = level.LevelText?.Val?.Value ?? "%1.",
                    Start = level.StartNumberingValue?.Val?.Value ?? 1,
                };

                // Parse indentation
                if (level.PreviousParagraphProperties?.Indentation != null)
                {
                    var indent = level.PreviousParagraphProperties.Indentation;
                    if (indent.Left?.Value != null)
                        numLevel.Indent = int.TryParse(indent.Left.Value, out var ind) ? ind : null;
                    if (indent.Hanging?.Value != null)
                        numLevel.HangingIndent = int.TryParse(indent.Hanging.Value, out var hang) ? hang : null;
                }

                // Parse alignment
                if (level.LevelJustification?.Val != null)
                {
                    numLevel.Alignment = level.LevelJustification.Val.ToString()?.ToLower() ?? "left";
                }

                // Parse font
                if (level.NumberingSymbolRunProperties?.RunFonts?.Ascii?.Value != null)
                {
                    numLevel.Font = level.NumberingSymbolRunProperties.RunFonts.Ascii.Value;
                }

                definition.Levels.Add(numLevel);
            }

            document.NumberingDefinitions.Add(definition);
        }

        // Parse numbering instances
        foreach (var numInstance in numbering.Elements<NumberingInstance>())
        {
            var instance = new NumberingInstance
            {
                Id = numInstance.NumberID?.Value ?? 0,
                DefinitionId = numInstance.AbstractNumId?.Val?.Value ?? 0
            };

            // Parse level overrides
            foreach (var lvlOverride in numInstance.Elements<LevelOverride>())
            {
                var over = new NumberingLevelOverride
                {
                    Level = lvlOverride.LevelIndex?.Value ?? 0,
                    StartOverride = lvlOverride.StartOverrideNumberingValue?.Val?.Value
                };

                if (lvlOverride.Level != null)
                {
                    over.LevelDefinition = new NumberingLevel
                    {
                        Level = lvlOverride.Level.LevelIndex?.Value ?? 0,
                        Format = lvlOverride.Level.NumberingFormat?.Val?.ToString() ?? "decimal",
                        Text = lvlOverride.Level.LevelText?.Val?.Value ?? "%1.",
                        Start = lvlOverride.Level.StartNumberingValue?.Val?.Value ?? 1
                    };
                }

                instance.LevelOverrides[over.Level] = over;
            }

            document.NumberingInstances.Add(instance);
        }
    }

    /// <summary>
    /// Parses style definitions.
    /// </summary>
    private void ParseStyles(MainDocumentPart mainPart, Document document)
    {
        var stylesPart = mainPart.StyleDefinitionsPart;
        if (stylesPart?.Styles == null) return;

        foreach (var style in stylesPart.Styles.Elements<Style>())
        {
            var styleDef = new StyleDefinition
            {
                Id = style.StyleId?.Value ?? string.Empty,
                Name = style.StyleName?.Val?.Value,
                BasedOn = style.BasedOn?.Val?.Value,
                NextStyle = style.NextParagraphStyle?.Val?.Value,
                Type = style.Type?.Value switch
                {
                    StyleValues.Paragraph => StyleType.Paragraph,
                    StyleValues.Character => StyleType.Character,
                    StyleValues.Table => StyleType.Table,
                    StyleValues.Numbering => StyleType.Numbering,
                    _ => StyleType.Paragraph
                }
            };

            // Parse paragraph properties
            if (style.StyleParagraphProperties != null)
            {
                styleDef.ParagraphProperties = ParseParagraphStyle(style.StyleParagraphProperties);
            }

            // Parse run properties
            if (style.StyleRunProperties != null)
            {
                styleDef.RunProperties = ParseRunFormatting(style.StyleRunProperties);
            }

            document.Styles.Add(styleDef);
        }
    }

    /// <summary>
    /// Parses the document body into sections.
    /// </summary>
    private List<Section> ParseBody(MainDocumentPart mainPart)
    {
        var sections = new List<Section>();
        var body = mainPart.Document.Body;
        if (body == null) return sections;

        var currentSection = new Section();
        var blocks = new List<Core.Models.Block>();

        foreach (var element in body.Elements())
        {
            if (element is DocumentFormat.OpenXml.Wordprocessing.Paragraph para)
            {
                var paragraph = ParseParagraph(para);
                blocks.Add(paragraph);

                // Check for section break in paragraph properties
                var sectPr = para.ParagraphProperties?.SectionProperties;
                if (sectPr != null)
                {
                    currentSection.Blocks = blocks;
                    currentSection.Properties = ParseSectionProperties(sectPr);
                    sections.Add(currentSection);
                    currentSection = new Section();
                    blocks = new List<Core.Models.Block>();
                }
            }
            else if (element is DocumentFormat.OpenXml.Wordprocessing.Table table)
            {
                var tableBlock = ParseTable(table);
                blocks.Add(tableBlock);
            }
            else if (element is SectionProperties finalSectPr)
            {
                // Final section properties at the end of body
                currentSection.Properties = ParseSectionProperties(finalSectPr);
            }
        }

        // Add the final section
        if (blocks.Count > 0 || sections.Count == 0)
        {
            currentSection.Blocks = blocks;
            sections.Add(currentSection);
        }

        // Parse headers and footers
        ParseHeadersFooters(mainPart, sections);

        return sections;
    }

    /// <summary>
    /// Parses a Word paragraph into the internal model.
    /// </summary>
    private Paragraph ParseParagraph(DocumentFormat.OpenXml.Wordprocessing.Paragraph para)
    {
        var paragraph = new Paragraph();

        // Parse paragraph properties
        if (para.ParagraphProperties != null)
        {
            paragraph.Style = ParseParagraphStyle(para.ParagraphProperties);

            // Parse numbering - CRITICAL
            var numPr = para.ParagraphProperties.NumberingProperties;
            if (numPr != null)
            {
                paragraph.Numbering = new NumberingInfo
                {
                    NumberingId = numPr.NumberingId?.Val?.Value ?? 0,
                    Level = numPr.NumberingLevelReference?.Val?.Value ?? 0
                };
            }
        }

        // Parse runs
        foreach (var run in para.Elements<DocumentFormat.OpenXml.Wordprocessing.Run>())
        {
            var parsedRun = ParseRun(run);
            if (parsedRun != null)
            {
                paragraph.Runs.Add(parsedRun);
            }
        }

        // Parse bookmarks
        foreach (var bookmarkStart in para.Elements<BookmarkStart>())
        {
            if (!string.IsNullOrEmpty(bookmarkStart.Name?.Value))
            {
                paragraph.BookmarkStarts.Add(bookmarkStart.Name.Value);
            }
        }

        foreach (var bookmarkEnd in para.Elements<BookmarkEnd>())
        {
            if (!string.IsNullOrEmpty(bookmarkEnd.Id?.Value))
            {
                paragraph.BookmarkEnds.Add(bookmarkEnd.Id.Value);
            }
        }

        return paragraph;
    }

    /// <summary>
    /// Parses paragraph properties into ParagraphStyle.
    /// </summary>
    private ParagraphStyle ParseParagraphStyle(ParagraphProperties props)
    {
        var style = new ParagraphStyle
        {
            StyleId = props.ParagraphStyleId?.Val?.Value
        };

        // Parse alignment
        if (props.Justification?.Val != null)
        {
            style.Alignment = props.Justification.Val.Value switch
            {
                JustificationValues.Center => "center",
                JustificationValues.Right => "right",
                JustificationValues.Both => "justify",
                _ => "left"
            };
        }

        // Parse indentation
        var indent = props.Indentation;
        if (indent != null)
        {
            if (indent.Left?.Value != null)
                style.LeftIndent = int.TryParse(indent.Left.Value, out var left) ? left : null;
            if (indent.Right?.Value != null)
                style.RightIndent = int.TryParse(indent.Right.Value, out var right) ? right : null;
            if (indent.FirstLine?.Value != null)
                style.FirstLineIndent = int.TryParse(indent.FirstLine.Value, out var first) ? first : null;
            if (indent.Hanging?.Value != null)
                style.FirstLineIndent = -(int.TryParse(indent.Hanging.Value, out var hang) ? hang : 0);
        }

        // Parse spacing
        var spacing = props.SpacingBetweenLines;
        if (spacing != null)
        {
            if (spacing.Before?.Value != null)
                style.SpaceBefore = int.TryParse(spacing.Before.Value, out var before) ? before : null;
            if (spacing.After?.Value != null)
                style.SpaceAfter = int.TryParse(spacing.After.Value, out var after) ? after : null;
            if (spacing.Line?.Value != null)
                style.LineSpacing = int.TryParse(spacing.Line.Value, out var line) ? line : null;
            style.LineSpacingRule = spacing.LineRule?.Value?.ToString()?.ToLower();
        }

        // Parse keep properties
        style.KeepWithNext = props.KeepNext != null;
        style.KeepLinesTogether = props.KeepLines != null;
        style.PageBreakBefore = props.PageBreakBefore != null;

        // Parse outline level
        if (props.OutlineLevel?.Val != null)
        {
            style.OutlineLevel = props.OutlineLevel.Val.Value;
        }

        return style;
    }

    /// <summary>
    /// Parses paragraph properties for styles (different element type).
    /// </summary>
    private ParagraphStyle ParseParagraphStyle(StyleParagraphProperties props)
    {
        var style = new ParagraphStyle();

        if (props.Justification?.Val != null)
        {
            style.Alignment = props.Justification.Val.Value switch
            {
                JustificationValues.Center => "center",
                JustificationValues.Right => "right",
                JustificationValues.Both => "justify",
                _ => "left"
            };
        }

        var indent = props.Indentation;
        if (indent != null)
        {
            if (indent.Left?.Value != null)
                style.LeftIndent = int.TryParse(indent.Left.Value, out var left) ? left : null;
            if (indent.Right?.Value != null)
                style.RightIndent = int.TryParse(indent.Right.Value, out var right) ? right : null;
        }

        var spacing = props.SpacingBetweenLines;
        if (spacing != null)
        {
            if (spacing.Before?.Value != null)
                style.SpaceBefore = int.TryParse(spacing.Before.Value, out var before) ? before : null;
            if (spacing.After?.Value != null)
                style.SpaceAfter = int.TryParse(spacing.After.Value, out var after) ? after : null;
        }

        style.KeepWithNext = props.KeepNext != null;
        style.KeepLinesTogether = props.KeepLines != null;

        return style;
    }

    /// <summary>
    /// Parses a run into the internal model.
    /// </summary>
    private Run? ParseRun(DocumentFormat.OpenXml.Wordprocessing.Run run)
    {
        var text = string.Concat(run.Elements<Text>().Select(t => t.Text));

        // Include tabs and breaks as spaces
        text += string.Concat(Enumerable.Repeat(" ", run.Elements<TabChar>().Count()));
        text += string.Concat(Enumerable.Repeat("\n", run.Elements<Break>().Count()));

        if (string.IsNullOrEmpty(text)) return null;

        var formatting = ParseRunFormatting(run.RunProperties);

        return new Run(text, formatting);
    }

    /// <summary>
    /// Parses run properties into RunFormatting.
    /// </summary>
    private RunFormatting ParseRunFormatting(RunProperties? props)
    {
        var formatting = new RunFormatting();

        if (props == null) return formatting;

        formatting.Bold = props.Bold != null && (props.Bold.Val == null || props.Bold.Val.Value);
        formatting.Italic = props.Italic != null && (props.Italic.Val == null || props.Italic.Val.Value);
        formatting.Underline = props.Underline?.Val != null && props.Underline.Val != UnderlineValues.None;
        formatting.Strikethrough = props.Strike != null && (props.Strike.Val == null || props.Strike.Val.Value);

        if (props.RunFonts?.Ascii?.Value != null)
            formatting.FontFamily = props.RunFonts.Ascii.Value;

        if (props.FontSize?.Val?.Value != null)
            formatting.FontSize = double.Parse(props.FontSize.Val.Value) / 2; // Half-points to points

        if (props.Color?.Val?.Value != null)
            formatting.Color = props.Color.Val.Value;

        if (props.Highlight?.Val != null)
            formatting.HighlightColor = props.Highlight.Val.ToString();

        formatting.Superscript = props.VerticalTextAlignment?.Val == VerticalPositionValues.Superscript;
        formatting.Subscript = props.VerticalTextAlignment?.Val == VerticalPositionValues.Subscript;

        formatting.StyleId = props.RunStyle?.Val?.Value;

        return formatting;
    }

    /// <summary>
    /// Parses run properties for styles.
    /// </summary>
    private RunFormatting ParseRunFormatting(StyleRunProperties? props)
    {
        var formatting = new RunFormatting();

        if (props == null) return formatting;

        formatting.Bold = props.Bold != null && (props.Bold.Val == null || props.Bold.Val.Value);
        formatting.Italic = props.Italic != null && (props.Italic.Val == null || props.Italic.Val.Value);
        formatting.Underline = props.Underline?.Val != null && props.Underline.Val != UnderlineValues.None;
        formatting.Strikethrough = props.Strike != null && (props.Strike.Val == null || props.Strike.Val.Value);

        if (props.RunFonts?.Ascii?.Value != null)
            formatting.FontFamily = props.RunFonts.Ascii.Value;

        if (props.FontSize?.Val?.Value != null)
            formatting.FontSize = double.Parse(props.FontSize.Val.Value) / 2;

        if (props.Color?.Val?.Value != null)
            formatting.Color = props.Color.Val.Value;

        return formatting;
    }

    /// <summary>
    /// Parses a table into the internal model.
    /// </summary>
    private Table ParseTable(DocumentFormat.OpenXml.Wordprocessing.Table table)
    {
        var tableBlock = new Table();

        foreach (var row in table.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>())
        {
            var tableRow = new TableRow();

            foreach (var cell in row.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCell>())
            {
                var tableCell = new TableCell();

                foreach (var para in cell.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
                {
                    tableCell.Blocks.Add(ParseParagraph(para));
                }

                tableRow.Cells.Add(tableCell);
            }

            tableBlock.Rows.Add(tableRow);
        }

        return tableBlock;
    }

    /// <summary>
    /// Parses section properties.
    /// </summary>
    private SectionProperties ParseSectionProperties(DocumentFormat.OpenXml.Wordprocessing.SectionProperties sectPr)
    {
        var props = new SectionProperties();

        var pageSize = sectPr.GetFirstChild<PageSize>();
        if (pageSize != null)
        {
            if (pageSize.Width?.Value != null)
                props.PageWidth = (int)pageSize.Width.Value;
            if (pageSize.Height?.Value != null)
                props.PageHeight = (int)pageSize.Height.Value;
            props.Orientation = pageSize.Orient?.Value == PageOrientationValues.Landscape ? "landscape" : "portrait";
        }

        var margins = sectPr.GetFirstChild<PageMargin>();
        if (margins != null)
        {
            if (margins.Left?.Value != null)
                props.MarginLeft = (int)margins.Left.Value;
            if (margins.Right?.Value != null)
                props.MarginRight = (int)margins.Right.Value;
            props.MarginTop = margins.Top?.Value ?? 1440;
            props.MarginBottom = margins.Bottom?.Value ?? 1440;
            if (margins.Header?.Value != null)
                props.HeaderDistance = (int)margins.Header.Value;
            if (margins.Footer?.Value != null)
                props.FooterDistance = (int)margins.Footer.Value;
        }

        var sectType = sectPr.GetFirstChild<SectionType>();
        if (sectType?.Val != null)
        {
            props.SectionBreakType = sectType.Val.Value switch
            {
                SectionMarkValues.Continuous => "continuous",
                SectionMarkValues.EvenPage => "evenPage",
                SectionMarkValues.OddPage => "oddPage",
                _ => "nextPage"
            };
        }

        props.DifferentFirstPage = sectPr.GetFirstChild<TitlePage>() != null;

        return props;
    }

    /// <summary>
    /// Parses headers and footers.
    /// </summary>
    private void ParseHeadersFooters(MainDocumentPart mainPart, List<Section> sections)
    {
        // Get all header and footer parts
        var headerParts = mainPart.HeaderParts.ToList();
        var footerParts = mainPart.FooterParts.ToList();

        foreach (var section in sections)
        {
            // Parse headers
            foreach (var headerPart in headerParts)
            {
                var header = ParseHeaderFooterContent(headerPart.Header);
                section.Headers.Default ??= header;
            }

            // Parse footers
            foreach (var footerPart in footerParts)
            {
                var footer = ParseHeaderFooterContent(footerPart.Footer);
                section.Footers.Default ??= footer;
            }
        }
    }

    /// <summary>
    /// Parses header/footer content.
    /// </summary>
    private HeaderFooter? ParseHeaderFooterContent(OpenXmlCompositeElement? element)
    {
        if (element == null) return null;

        var headerFooter = new HeaderFooter();

        foreach (var para in element.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
        {
            headerFooter.Blocks.Add(ParseParagraph(para));
        }

        return headerFooter.Blocks.Count > 0 ? headerFooter : null;
    }
}
