using DocumentCompare.Core.Models;
using FluentAssertions;
using Xunit;

namespace DocumentCompare.Core.Tests;

public class DocumentModelTests
{
    [Fact]
    public void Paragraph_GetPlainText_ConcatenatesRuns()
    {
        // Arrange
        var paragraph = new Paragraph
        {
            Runs = new List<Run>
            {
                new Run("Hello "),
                new Run("world"),
                new Run("!")
            }
        };

        // Act
        var text = paragraph.GetPlainText();

        // Assert
        text.Should().Be("Hello world!");
    }

    [Fact]
    public void Paragraph_GetNormalizedText_TrimsAndNormalizesWhitespace()
    {
        // Arrange
        var paragraph = new Paragraph
        {
            Runs = new List<Run>
            {
                new Run("  Hello   "),
                new Run("  world  ")
            }
        };

        // Act
        var text = paragraph.GetNormalizedText();

        // Assert
        text.Should().Be("Hello world");
    }

    [Fact]
    public void Paragraph_Clone_CreatesDeepCopy()
    {
        // Arrange
        var original = new Paragraph
        {
            Id = "test-id",
            Runs = new List<Run>
            {
                new Run("Hello", new RunFormatting { Bold = true })
            },
            Style = new ParagraphStyle { Alignment = "center" },
            Numbering = new NumberingInfo { NumberingId = 1, Level = 0 }
        };

        // Act
        var clone = (Paragraph)original.Clone();

        // Assert
        clone.Should().NotBeSameAs(original);
        clone.Id.Should().Be(original.Id);
        clone.Runs.Should().NotBeSameAs(original.Runs);
        clone.Runs[0].Should().NotBeSameAs(original.Runs[0]);
        clone.Runs[0].Text.Should().Be(original.Runs[0].Text);
        clone.Runs[0].Formatting.Bold.Should().Be(original.Runs[0].Formatting.Bold);
        clone.Style.Should().NotBeSameAs(original.Style);
        clone.Numbering.Should().NotBeSameAs(original.Numbering);
    }

    [Fact]
    public void RunFormatting_ForDeletion_SetsRedStrikethrough()
    {
        // Act
        var formatting = RunFormatting.ForDeletion();

        // Assert
        formatting.Strikethrough.Should().BeTrue();
        formatting.Color.Should().Be("FF0000");
    }

    [Fact]
    public void RunFormatting_ForInsertion_SetsBoldBlue()
    {
        // Act
        var formatting = RunFormatting.ForInsertion();

        // Assert
        formatting.Bold.Should().BeTrue();
        formatting.Color.Should().Be("0000FF");
    }

    [Fact]
    public void RunFormatting_ForMove_SetsGreen()
    {
        // Act
        var formatting = RunFormatting.ForMove();

        // Assert
        formatting.Color.Should().Be("008000");
    }

    [Fact]
    public void RunFormatting_ForMoveSource_SetsGreenStrikethrough()
    {
        // Act
        var formatting = RunFormatting.ForMove(isSource: true);

        // Assert
        formatting.Color.Should().Be("008000");
        formatting.Strikethrough.Should().BeTrue();
    }

    [Fact]
    public void RunFormatting_ForDeletion_PreservesOriginalFormatting()
    {
        // Arrange
        var original = new RunFormatting
        {
            Bold = true,
            FontFamily = "Arial",
            FontSize = 12
        };

        // Act
        var formatting = RunFormatting.ForDeletion(original);

        // Assert
        formatting.Bold.Should().BeTrue();
        formatting.FontFamily.Should().Be("Arial");
        formatting.FontSize.Should().Be(12);
        formatting.Strikethrough.Should().BeTrue();
        formatting.Color.Should().Be("FF0000");
    }

    [Fact]
    public void Document_GetAllParagraphs_ReturnsParagraphsInOrder()
    {
        // Arrange
        var document = new Document
        {
            Sections = new List<Section>
            {
                new Section
                {
                    Blocks = new List<Block>
                    {
                        new Paragraph { Runs = new List<Run> { new Run("First") } },
                        new Paragraph { Runs = new List<Run> { new Run("Second") } }
                    }
                },
                new Section
                {
                    Blocks = new List<Block>
                    {
                        new Paragraph { Runs = new List<Run> { new Run("Third") } }
                    }
                }
            }
        };

        // Act
        var paragraphs = document.GetAllParagraphs().ToList();

        // Assert
        paragraphs.Should().HaveCount(3);
        paragraphs[0].GetPlainText().Should().Be("First");
        paragraphs[1].GetPlainText().Should().Be("Second");
        paragraphs[2].GetPlainText().Should().Be("Third");
    }

    [Fact]
    public void Document_GetPlainText_ConcatenatesAllParagraphs()
    {
        // Arrange
        var document = new Document
        {
            Sections = new List<Section>
            {
                new Section
                {
                    Blocks = new List<Block>
                    {
                        new Paragraph { Runs = new List<Run> { new Run("Hello") } },
                        new Paragraph { Runs = new List<Run> { new Run("World") } }
                    }
                }
            }
        };

        // Act
        var text = document.GetPlainText();

        // Assert
        text.Should().Be("Hello\nWorld");
    }

    [Fact]
    public void NumberingDefinition_Clone_CreatesDeepCopy()
    {
        // Arrange
        var original = new NumberingDefinition
        {
            Id = 1,
            Name = "Legal",
            MultiLevel = true,
            Levels = new List<NumberingLevel>
            {
                new NumberingLevel { Level = 0, Format = "decimal", Text = "%1." },
                new NumberingLevel { Level = 1, Format = "decimal", Text = "%1.%2" }
            }
        };

        // Act
        var clone = original.Clone();

        // Assert
        clone.Should().NotBeSameAs(original);
        clone.Id.Should().Be(original.Id);
        clone.Levels.Should().NotBeSameAs(original.Levels);
        clone.Levels[0].Should().NotBeSameAs(original.Levels[0]);
        clone.Levels[0].Format.Should().Be(original.Levels[0].Format);
    }
}
