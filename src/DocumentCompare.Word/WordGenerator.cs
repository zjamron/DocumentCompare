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
/// Generates Word documents (.docx) from the internal document model.
/// </summary>
public class WordGenerator : IDocumentGenerator
{
    public string OutputFormat => "docx";

    public void Generate(Document document, string outputPath)
    {
        using var stream = File.Create(outputPath);
        Generate(document, stream);
    }

    public void Generate(Document document, Stream outputStream)
    {
        using var wordDoc = WordprocessingDocument.Create(outputStream, WordprocessingDocumentType.Document);

        // Create main document part
        var mainPart = wordDoc.AddMainDocumentPart();
        mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();

        // Generate numbering (CRITICAL for section numbers)
        GenerateNumbering(mainPart, document);

        // Generate styles
        GenerateStyles(mainPart, document);

        // Generate body
        var body = new Body();
        GenerateBody(body, document, mainPart);
        mainPart.Document.Body = body;

        // Save document
        mainPart.Document.Save();
    }

    /// <summary>
    /// Generates numbering definitions and instances - CRITICAL for preserving section numbering.
    /// </summary>
    private void GenerateNumbering(MainDocumentPart mainPart, Document document)
    {
        if (document.NumberingDefinitions.Count == 0 && document.NumberingInstances.Count == 0)
            return;

        var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
        var numbering = new Numbering();

        // Generate abstract numbering definitions
        foreach (var def in document.NumberingDefinitions)
        {
            var abstractNum = new AbstractNum { AbstractNumberId = def.Id };

            // Set multi-level type
            abstractNum.MultiLevelType = new MultiLevelType
            {
                Val = def.MultiLevel ? MultiLevelValues.Multilevel : MultiLevelValues.SingleLevel
            };

            // Generate levels
            foreach (var levelDef in def.Levels.OrderBy(l => l.Level))
            {
                var level = new Level { LevelIndex = levelDef.Level };

                // Number format
                level.NumberingFormat = new NumberingFormat
                {
                    Val = ParseNumberFormat(levelDef.Format)
                };

                // Level text (e.g., "%1.", "%1.%2")
                level.LevelText = new LevelText { Val = levelDef.Text };

                // Start value
                level.StartNumberingValue = new StartNumberingValue { Val = levelDef.Start };

                // Alignment
                level.LevelJustification = new LevelJustification
                {
                    Val = levelDef.Alignment?.ToLower() switch
                    {
                        "center" => LevelJustificationValues.Center,
                        "right" => LevelJustificationValues.Right,
                        _ => LevelJustificationValues.Left
                    }
                };

                // Indentation
                if (levelDef.Indent.HasValue || levelDef.HangingIndent.HasValue)
                {
                    level.PreviousParagraphProperties = new PreviousParagraphProperties
                    {
                        Indentation = new Indentation
                        {
                            Left = levelDef.Indent?.ToString(),
                            Hanging = levelDef.HangingIndent?.ToString()
                        }
                    };
                }

                // Font for numbering symbol
                if (!string.IsNullOrEmpty(levelDef.Font))
                {
                    level.NumberingSymbolRunProperties = new NumberingSymbolRunProperties
                    {
                        RunFonts = new RunFonts { Ascii = levelDef.Font }
                    };
                }

                abstractNum.AppendChild(level);
            }

            numbering.AppendChild(abstractNum);
        }

        // Generate numbering instances
        foreach (var instance in document.NumberingInstances)
        {
            var numInstance = new NumberingInstance { NumberID = instance.Id };
            numInstance.AbstractNumId = new AbstractNumId { Val = instance.DefinitionId };

            // Level overrides
            foreach (var over in instance.LevelOverrides.Values)
            {
                var lvlOverride = new LevelOverride { LevelIndex = over.Level };

                if (over.StartOverride.HasValue)
                {
                    lvlOverride.StartOverrideNumberingValue = new StartOverrideNumberingValue
                    {
                        Val = over.StartOverride.Value
                    };
                }

                numInstance.AppendChild(lvlOverride);
            }

            numbering.AppendChild(numInstance);
        }

        numberingPart.Numbering = numbering;
        numberingPart.Numbering.Save();
    }

    /// <summary>
    /// Parses number format string to OpenXML enum.
    /// </summary>
    private NumberFormatValues ParseNumberFormat(string format)
    {
        return format.ToLower() switch
        {
            "decimal" => NumberFormatValues.Decimal,
            "lowerletter" => NumberFormatValues.LowerLetter,
            "upperletter" => NumberFormatValues.UpperLetter,
            "lowerroman" => NumberFormatValues.LowerRoman,
            "upperroman" => NumberFormatValues.UpperRoman,
            "bullet" => NumberFormatValues.Bullet,
            "none" => NumberFormatValues.None,
            "ordinal" => NumberFormatValues.Ordinal,
            "cardinaltext" => NumberFormatValues.CardinalText,
            "ordinaltext" => NumberFormatValues.OrdinalText,
            _ => NumberFormatValues.Decimal
        };
    }

    /// <summary>
    /// Generates style definitions.
    /// </summary>
    private void GenerateStyles(MainDocumentPart mainPart, Document document)
    {
        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        var styles = new Styles();

        // Add default styles
        styles.AppendChild(CreateDefaultStyle());

        // Add document styles
        foreach (var styleDef in document.Styles)
        {
            var style = new Style
            {
                StyleId = styleDef.Id,
                Type = styleDef.Type switch
                {
                    StyleType.Character => StyleValues.Character,
                    StyleType.Table => StyleValues.Table,
                    StyleType.Numbering => StyleValues.Numbering,
                    _ => StyleValues.Paragraph
                }
            };

            if (!string.IsNullOrEmpty(styleDef.Name))
                style.StyleName = new StyleName { Val = styleDef.Name };

            if (!string.IsNullOrEmpty(styleDef.BasedOn))
                style.BasedOn = new BasedOn { Val = styleDef.BasedOn };

            if (!string.IsNullOrEmpty(styleDef.NextStyle))
                style.NextParagraphStyle = new NextParagraphStyle { Val = styleDef.NextStyle };

            // Paragraph properties
            if (styleDef.ParagraphProperties != null)
            {
                style.StyleParagraphProperties = CreateStyleParagraphProperties(styleDef.ParagraphProperties);
            }

            // Run properties
            if (styleDef.RunProperties != null)
            {
                style.StyleRunProperties = CreateStyleRunProperties(styleDef.RunProperties);
            }

            styles.AppendChild(style);
        }

        stylesPart.Styles = styles;
        stylesPart.Styles.Save();
    }

    /// <summary>
    /// Creates a default Normal style.
    /// </summary>
    private Style CreateDefaultStyle()
    {
        return new Style
        {
            StyleId = "Normal",
            Type = StyleValues.Paragraph,
            Default = true,
            StyleName = new StyleName { Val = "Normal" },
            PrimaryStyle = new PrimaryStyle()
        };
    }

    /// <summary>
    /// Creates style paragraph properties.
    /// </summary>
    private StyleParagraphProperties CreateStyleParagraphProperties(ParagraphStyle paraStyle)
    {
        var props = new StyleParagraphProperties();

        if (paraStyle.Alignment != "left")
        {
            props.Justification = new Justification
            {
                Val = paraStyle.Alignment switch
                {
                    "center" => JustificationValues.Center,
                    "right" => JustificationValues.Right,
                    "justify" => JustificationValues.Both,
                    _ => JustificationValues.Left
                }
            };
        }

        if (paraStyle.SpaceBefore.HasValue || paraStyle.SpaceAfter.HasValue)
        {
            props.SpacingBetweenLines = new SpacingBetweenLines
            {
                Before = paraStyle.SpaceBefore?.ToString(),
                After = paraStyle.SpaceAfter?.ToString()
            };
        }

        return props;
    }

    /// <summary>
    /// Creates style run properties.
    /// </summary>
    private StyleRunProperties CreateStyleRunProperties(RunFormatting formatting)
    {
        var props = new StyleRunProperties();

        if (formatting.Bold)
            props.Bold = new Bold();

        if (formatting.Italic)
            props.Italic = new Italic();

        if (formatting.Underline)
            props.Underline = new Underline { Val = UnderlineValues.Single };

        if (formatting.Strikethrough)
            props.Strike = new Strike();

        if (!string.IsNullOrEmpty(formatting.FontFamily))
            props.RunFonts = new RunFonts { Ascii = formatting.FontFamily };

        if (formatting.FontSize.HasValue)
            props.FontSize = new FontSize { Val = ((int)(formatting.FontSize.Value * 2)).ToString() };

        if (!string.IsNullOrEmpty(formatting.Color))
            props.Color = new Color { Val = formatting.Color };

        return props;
    }

    /// <summary>
    /// Generates the document body.
    /// </summary>
    private void GenerateBody(Body body, Document document, MainDocumentPart mainPart)
    {
        for (int i = 0; i < document.Sections.Count; i++)
        {
            var section = document.Sections[i];
            var isLastSection = i == document.Sections.Count - 1;

            // Generate blocks
            foreach (var block in section.Blocks)
            {
                if (block is Paragraph para)
                {
                    var wordPara = GenerateParagraph(para);

                    // Add section properties to last paragraph of non-last sections
                    if (!isLastSection && block == section.Blocks.Last())
                    {
                        wordPara.ParagraphProperties ??= new ParagraphProperties();
                        wordPara.ParagraphProperties.SectionProperties = GenerateSectionProperties(section.Properties);
                    }

                    body.AppendChild(wordPara);
                }
                else if (block is Table table)
                {
                    body.AppendChild(GenerateTable(table));
                }
            }

            // Add final section properties
            if (isLastSection)
            {
                body.AppendChild(GenerateSectionProperties(section.Properties));
            }
        }
    }

    /// <summary>
    /// Generates a Word paragraph from the internal model.
    /// </summary>
    private DocumentFormat.OpenXml.Wordprocessing.Paragraph GenerateParagraph(Paragraph para)
    {
        var wordPara = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();

        // Generate paragraph properties
        var props = new ParagraphProperties();
        var hasProps = false;

        // Style
        if (!string.IsNullOrEmpty(para.Style.StyleId))
        {
            props.ParagraphStyleId = new ParagraphStyleId { Val = para.Style.StyleId };
            hasProps = true;
        }

        // Numbering - CRITICAL
        if (para.Numbering != null)
        {
            props.NumberingProperties = new NumberingProperties
            {
                NumberingId = new NumberingId { Val = para.Numbering.NumberingId },
                NumberingLevelReference = new NumberingLevelReference { Val = para.Numbering.Level }
            };
            hasProps = true;
        }

        // Alignment
        if (para.Style.Alignment != "left")
        {
            props.Justification = new Justification
            {
                Val = para.Style.Alignment switch
                {
                    "center" => JustificationValues.Center,
                    "right" => JustificationValues.Right,
                    "justify" => JustificationValues.Both,
                    _ => JustificationValues.Left
                }
            };
            hasProps = true;
        }

        // Indentation
        if (para.Style.LeftIndent.HasValue || para.Style.RightIndent.HasValue || para.Style.FirstLineIndent.HasValue)
        {
            var indent = new Indentation();
            if (para.Style.LeftIndent.HasValue)
                indent.Left = para.Style.LeftIndent.Value.ToString();
            if (para.Style.RightIndent.HasValue)
                indent.Right = para.Style.RightIndent.Value.ToString();
            if (para.Style.FirstLineIndent.HasValue)
            {
                if (para.Style.FirstLineIndent.Value >= 0)
                    indent.FirstLine = para.Style.FirstLineIndent.Value.ToString();
                else
                    indent.Hanging = (-para.Style.FirstLineIndent.Value).ToString();
            }
            props.Indentation = indent;
            hasProps = true;
        }

        // Spacing
        if (para.Style.SpaceBefore.HasValue || para.Style.SpaceAfter.HasValue || para.Style.LineSpacing.HasValue)
        {
            var spacing = new SpacingBetweenLines();
            if (para.Style.SpaceBefore.HasValue)
                spacing.Before = para.Style.SpaceBefore.Value.ToString();
            if (para.Style.SpaceAfter.HasValue)
                spacing.After = para.Style.SpaceAfter.Value.ToString();
            if (para.Style.LineSpacing.HasValue)
                spacing.Line = para.Style.LineSpacing.Value.ToString();
            props.SpacingBetweenLines = spacing;
            hasProps = true;
        }

        // Keep properties
        if (para.Style.KeepWithNext)
        {
            props.KeepNext = new KeepNext();
            hasProps = true;
        }
        if (para.Style.KeepLinesTogether)
        {
            props.KeepLines = new KeepLines();
            hasProps = true;
        }
        if (para.Style.PageBreakBefore)
        {
            props.PageBreakBefore = new PageBreakBefore();
            hasProps = true;
        }

        if (hasProps)
        {
            wordPara.ParagraphProperties = props;
        }

        // Generate runs
        foreach (var run in para.Runs)
        {
            wordPara.AppendChild(GenerateRun(run));
        }

        return wordPara;
    }

    /// <summary>
    /// Generates a Word run from the internal model.
    /// </summary>
    private DocumentFormat.OpenXml.Wordprocessing.Run GenerateRun(Run run)
    {
        var wordRun = new DocumentFormat.OpenXml.Wordprocessing.Run();

        // Generate run properties
        var props = GenerateRunProperties(run.Formatting);
        if (props.HasChildren)
        {
            wordRun.RunProperties = props;
        }

        // Add text
        var text = new Text(run.Text);
        if (run.Text.StartsWith(" ") || run.Text.EndsWith(" "))
        {
            text.Space = SpaceProcessingModeValues.Preserve;
        }
        wordRun.AppendChild(text);

        return wordRun;
    }

    /// <summary>
    /// Generates run properties from formatting.
    /// </summary>
    private RunProperties GenerateRunProperties(RunFormatting formatting)
    {
        var props = new RunProperties();

        if (formatting.Bold)
            props.Bold = new Bold();

        if (formatting.Italic)
            props.Italic = new Italic();

        if (formatting.Underline)
            props.Underline = new Underline { Val = UnderlineValues.Single };

        if (formatting.Strikethrough)
            props.Strike = new Strike();

        if (!string.IsNullOrEmpty(formatting.FontFamily))
            props.RunFonts = new RunFonts { Ascii = formatting.FontFamily, HighAnsi = formatting.FontFamily };

        if (formatting.FontSize.HasValue)
            props.FontSize = new FontSize { Val = ((int)(formatting.FontSize.Value * 2)).ToString() };

        if (!string.IsNullOrEmpty(formatting.Color))
            props.Color = new Color { Val = formatting.Color };

        if (!string.IsNullOrEmpty(formatting.HighlightColor))
        {
            if (Enum.TryParse<HighlightColorValues>(formatting.HighlightColor, true, out var highlight))
            {
                props.Highlight = new Highlight { Val = highlight };
            }
        }

        if (formatting.Superscript)
            props.VerticalTextAlignment = new VerticalTextAlignment { Val = VerticalPositionValues.Superscript };
        else if (formatting.Subscript)
            props.VerticalTextAlignment = new VerticalTextAlignment { Val = VerticalPositionValues.Subscript };

        if (!string.IsNullOrEmpty(formatting.StyleId))
            props.RunStyle = new RunStyle { Val = formatting.StyleId };

        return props;
    }

    /// <summary>
    /// Generates a Word table from the internal model.
    /// </summary>
    private DocumentFormat.OpenXml.Wordprocessing.Table GenerateTable(Table table)
    {
        var wordTable = new DocumentFormat.OpenXml.Wordprocessing.Table();

        // Table properties
        var tableProps = new TableProperties(
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4 },
                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                new RightBorder { Val = BorderValues.Single, Size = 4 },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
            )
        );
        wordTable.AppendChild(tableProps);

        // Generate rows
        foreach (var row in table.Rows)
        {
            var wordRow = new DocumentFormat.OpenXml.Wordprocessing.TableRow();

            foreach (var cell in row.Cells)
            {
                var wordCell = new DocumentFormat.OpenXml.Wordprocessing.TableCell();

                // Cell properties
                var cellProps = new TableCellProperties();
                if (cell.Properties?.Width.HasValue == true)
                {
                    cellProps.TableCellWidth = new TableCellWidth
                    {
                        Width = cell.Properties.Width.Value.ToString(),
                        Type = TableWidthUnitValues.Dxa
                    };
                }
                wordCell.AppendChild(cellProps);

                // Cell content
                foreach (var block in cell.Blocks)
                {
                    if (block is Paragraph para)
                    {
                        wordCell.AppendChild(GenerateParagraph(para));
                    }
                }

                // Ensure at least one paragraph
                if (!wordCell.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Any())
                {
                    wordCell.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());
                }

                wordRow.AppendChild(wordCell);
            }

            wordTable.AppendChild(wordRow);
        }

        return wordTable;
    }

    /// <summary>
    /// Generates section properties.
    /// </summary>
    private DocumentFormat.OpenXml.Wordprocessing.SectionProperties GenerateSectionProperties(SectionProperties props)
    {
        var sectPr = new DocumentFormat.OpenXml.Wordprocessing.SectionProperties();

        // Page size
        sectPr.AppendChild(new PageSize
        {
            Width = (uint)props.PageWidth,
            Height = (uint)props.PageHeight,
            Orient = props.Orientation == "landscape" ? PageOrientationValues.Landscape : PageOrientationValues.Portrait
        });

        // Page margins
        sectPr.AppendChild(new PageMargin
        {
            Left = (uint)props.MarginLeft,
            Right = (uint)props.MarginRight,
            Top = props.MarginTop,
            Bottom = props.MarginBottom,
            Header = (uint)props.HeaderDistance,
            Footer = (uint)props.FooterDistance
        });

        // Section type
        if (props.SectionBreakType != "nextPage")
        {
            sectPr.AppendChild(new SectionType
            {
                Val = props.SectionBreakType switch
                {
                    "continuous" => SectionMarkValues.Continuous,
                    "evenPage" => SectionMarkValues.EvenPage,
                    "oddPage" => SectionMarkValues.OddPage,
                    _ => SectionMarkValues.NextPage
                }
            });
        }

        // Different first page
        if (props.DifferentFirstPage)
        {
            sectPr.AppendChild(new TitlePage());
        }

        return sectPr;
    }
}
