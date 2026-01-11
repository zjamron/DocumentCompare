using DocumentCompare.Core.Interfaces;
using DocumentCompare.Core.Models;
using FluentAssertions;
using Xunit;

namespace DocumentCompare.Word.Tests;

public class WordComparisonTests
{
    [Fact]
    public void WordParser_CanParse_ReturnsTrueForDocx()
    {
        // Arrange
        var parser = new WordParser();

        // Act & Assert
        parser.CanParse("document.docx").Should().BeTrue();
        parser.CanParse("DOCUMENT.DOCX").Should().BeTrue();
        parser.CanParse("document.pdf").Should().BeFalse();
    }

    [Fact]
    public void WordGenerator_OutputFormat_ReturnsDocx()
    {
        // Arrange
        var generator = new WordGenerator();

        // Assert
        generator.OutputFormat.Should().Be("docx");
    }

    [Fact]
    public void WordDocumentComparer_Create_ReturnsConfiguredComparer()
    {
        // Act
        var comparer = WordDocumentComparer.Create();

        // Assert
        comparer.Should().NotBeNull();
        comparer.Should().BeAssignableTo<IDocumentComparer>();
    }

    [Fact]
    public void WordGenerator_Generate_CreatesValidDocument()
    {
        // Arrange
        var generator = new WordGenerator();
        var document = CreateSampleDocument();
        var outputPath = Path.Combine(Path.GetTempPath(), $"test_output_{Guid.NewGuid()}.docx");

        try
        {
            // Act
            generator.Generate(document, outputPath);

            // Assert
            File.Exists(outputPath).Should().BeTrue();
            new FileInfo(outputPath).Length.Should().BeGreaterThan(0);
        }
        finally
        {
            // Cleanup
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }

    [Fact]
    public void WordGenerator_Generate_PreservesNumbering()
    {
        // Arrange
        var generator = new WordGenerator();
        var document = CreateDocumentWithNumbering();
        var outputPath = Path.Combine(Path.GetTempPath(), $"test_numbering_{Guid.NewGuid()}.docx");

        try
        {
            // Act
            generator.Generate(document, outputPath);

            // Assert - Parse the output and verify numbering
            var parser = new WordParser();
            var parsedDoc = parser.Parse(outputPath);

            parsedDoc.NumberingDefinitions.Should().NotBeEmpty();
            parsedDoc.NumberingInstances.Should().NotBeEmpty();
        }
        finally
        {
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }

    [Fact]
    public void RoundTrip_GenerateAndParse_PreservesContent()
    {
        // Arrange
        var generator = new WordGenerator();
        var parser = new WordParser();
        var originalText = "Hello world, this is a test document.";
        var document = CreateDocumentWithText(originalText);
        var outputPath = Path.Combine(Path.GetTempPath(), $"test_roundtrip_{Guid.NewGuid()}.docx");

        try
        {
            // Act
            generator.Generate(document, outputPath);
            var parsedDoc = parser.Parse(outputPath);

            // Assert
            parsedDoc.GetPlainText().Trim().Should().Contain("Hello world");
        }
        finally
        {
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }

    [Fact]
    public void RoundTrip_RedlinedDocument_PreservesFormatting()
    {
        // Arrange
        var generator = new WordGenerator();
        var parser = new WordParser();
        var document = CreateRedlinedDocument();
        var outputPath = Path.Combine(Path.GetTempPath(), $"test_redline_{Guid.NewGuid()}.docx");

        try
        {
            // Act
            generator.Generate(document, outputPath);
            var parsedDoc = parser.Parse(outputPath);

            // Assert
            var paragraphs = parsedDoc.GetAllParagraphs().ToList();
            paragraphs.Should().NotBeEmpty();

            // Check that formatting is preserved
            var runs = paragraphs.SelectMany(p => p.Runs).ToList();
            runs.Should().Contain(r => r.Formatting.Strikethrough && r.Formatting.Color == "FF0000"); // Deletion
            runs.Should().Contain(r => r.Formatting.Bold && r.Formatting.Color == "0000FF"); // Insertion
        }
        finally
        {
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }

    private Document CreateSampleDocument()
    {
        return new Document
        {
            Sections = new List<Section>
            {
                new Section
                {
                    Blocks = new List<Block>
                    {
                        new Paragraph
                        {
                            Runs = new List<Run>
                            {
                                new Run("This is a sample document.")
                            }
                        }
                    }
                }
            }
        };
    }

    private Document CreateDocumentWithText(string text)
    {
        return new Document
        {
            Sections = new List<Section>
            {
                new Section
                {
                    Blocks = new List<Block>
                    {
                        new Paragraph
                        {
                            Runs = new List<Run> { new Run(text) }
                        }
                    }
                }
            }
        };
    }

    private Document CreateDocumentWithNumbering()
    {
        var document = new Document
        {
            NumberingDefinitions = new List<NumberingDefinition>
            {
                new NumberingDefinition
                {
                    Id = 0,
                    MultiLevel = true,
                    Levels = new List<NumberingLevel>
                    {
                        new NumberingLevel { Level = 0, Format = "decimal", Text = "%1.", Start = 1 },
                        new NumberingLevel { Level = 1, Format = "decimal", Text = "%1.%2", Start = 1 }
                    }
                }
            },
            NumberingInstances = new List<NumberingInstance>
            {
                new NumberingInstance { Id = 1, DefinitionId = 0 }
            },
            Sections = new List<Section>
            {
                new Section
                {
                    Blocks = new List<Block>
                    {
                        new Paragraph
                        {
                            Runs = new List<Run> { new Run("First item") },
                            Numbering = new NumberingInfo { NumberingId = 1, Level = 0 }
                        },
                        new Paragraph
                        {
                            Runs = new List<Run> { new Run("Sub item") },
                            Numbering = new NumberingInfo { NumberingId = 1, Level = 1 }
                        },
                        new Paragraph
                        {
                            Runs = new List<Run> { new Run("Second item") },
                            Numbering = new NumberingInfo { NumberingId = 1, Level = 0 }
                        }
                    }
                }
            }
        };

        return document;
    }

    private Document CreateRedlinedDocument()
    {
        return new Document
        {
            Sections = new List<Section>
            {
                new Section
                {
                    Blocks = new List<Block>
                    {
                        new Paragraph
                        {
                            Runs = new List<Run>
                            {
                                new Run("This text is unchanged. "),
                                new Run("This was deleted.", RunFormatting.ForDeletion()),
                                new Run(" "),
                                new Run("This was added.", RunFormatting.ForInsertion()),
                                new Run(" More unchanged text.")
                            }
                        }
                    }
                }
            }
        };
    }
}
