using DocumentCompare.Core.Comparison;
using DocumentCompare.Core.Interfaces;
using DocumentCompare.Core.Models;
using FluentAssertions;
using Xunit;

namespace DocumentCompare.Core.Tests;

public class DiffEngineTests
{
    private readonly DiffEngine _diffEngine;

    public DiffEngineTests()
    {
        _diffEngine = new DiffEngine();
    }

    [Fact]
    public void DiffParagraphs_IdenticalText_ReturnsUnchanged()
    {
        // Arrange
        var original = CreateParagraph("Hello world");
        var modified = CreateParagraph("Hello world");

        // Act
        var result = _diffEngine.DiffParagraphs(original, modified);

        // Assert
        result.Segments.Should().ContainSingle();
        result.Segments[0].Type.Should().Be(DiffType.Unchanged);
        result.InsertionCount.Should().Be(0);
        result.DeletionCount.Should().Be(0);
    }

    [Fact]
    public void DiffParagraphs_AddedWord_ReturnsInsertion()
    {
        // Arrange
        var original = CreateParagraph("Hello world");
        var modified = CreateParagraph("Hello beautiful world");

        // Act
        var result = _diffEngine.DiffParagraphs(original, modified);

        // Assert
        result.InsertionCount.Should().BeGreaterThan(0);
        result.Segments.Should().Contain(s => s.Type == DiffType.Inserted && s.Text.Contains("beautiful"));
    }

    [Fact]
    public void DiffParagraphs_RemovedWord_ReturnsDeletion()
    {
        // Arrange
        var original = CreateParagraph("Hello beautiful world");
        var modified = CreateParagraph("Hello world");

        // Act
        var result = _diffEngine.DiffParagraphs(original, modified);

        // Assert
        result.DeletionCount.Should().BeGreaterThan(0);
        result.Segments.Should().Contain(s => s.Type == DiffType.Deleted && s.Text.Contains("beautiful"));
    }

    [Fact]
    public void DiffParagraphs_ChangedWord_ReturnsDeleteAndInsert()
    {
        // Arrange
        var original = CreateParagraph("Hello world");
        var modified = CreateParagraph("Hello universe");

        // Act
        var result = _diffEngine.DiffParagraphs(original, modified);

        // Assert
        result.DeletionCount.Should().BeGreaterThan(0);
        result.InsertionCount.Should().BeGreaterThan(0);
        result.Segments.Should().Contain(s => s.Type == DiffType.Deleted && s.Text.Contains("world"));
        result.Segments.Should().Contain(s => s.Type == DiffType.Inserted && s.Text.Contains("universe"));
    }

    [Fact]
    public void DiffParagraphs_EmptyOriginal_ReturnsAllInserted()
    {
        // Arrange
        var original = CreateParagraph("");
        var modified = CreateParagraph("New content");

        // Act
        var result = _diffEngine.DiffParagraphs(original, modified);

        // Assert
        result.IsEntirelyInserted.Should().BeTrue();
    }

    [Fact]
    public void DiffParagraphs_EmptyModified_ReturnsAllDeleted()
    {
        // Arrange
        var original = CreateParagraph("Old content");
        var modified = CreateParagraph("");

        // Act
        var result = _diffEngine.DiffParagraphs(original, modified);

        // Assert
        result.IsEntirelyDeleted.Should().BeTrue();
    }

    [Fact]
    public void AlignParagraphs_IdenticalParagraphs_AllUnchanged()
    {
        // Arrange
        var original = new List<Paragraph>
        {
            CreateParagraph("First paragraph"),
            CreateParagraph("Second paragraph"),
            CreateParagraph("Third paragraph")
        };
        var modified = new List<Paragraph>
        {
            CreateParagraph("First paragraph"),
            CreateParagraph("Second paragraph"),
            CreateParagraph("Third paragraph")
        };

        // Act
        var alignments = _diffEngine.AlignParagraphs(original, modified);

        // Assert
        alignments.Should().HaveCount(3);
        alignments.Should().OnlyContain(a => a.Type == DiffType.Unchanged);
    }

    [Fact]
    public void AlignParagraphs_InsertedParagraph_DetectsInsertion()
    {
        // Arrange
        var original = new List<Paragraph>
        {
            CreateParagraph("First paragraph"),
            CreateParagraph("Third paragraph")
        };
        var modified = new List<Paragraph>
        {
            CreateParagraph("First paragraph"),
            CreateParagraph("Second paragraph"), // Inserted
            CreateParagraph("Third paragraph")
        };

        // Act
        var alignments = _diffEngine.AlignParagraphs(original, modified);

        // Assert
        alignments.Should().HaveCount(3);
        alignments.Should().Contain(a => a.Type == DiffType.Inserted);
    }

    [Fact]
    public void AlignParagraphs_DeletedParagraph_DetectsDeletion()
    {
        // Arrange
        var original = new List<Paragraph>
        {
            CreateParagraph("First paragraph"),
            CreateParagraph("Second paragraph"), // Will be deleted
            CreateParagraph("Third paragraph")
        };
        var modified = new List<Paragraph>
        {
            CreateParagraph("First paragraph"),
            CreateParagraph("Third paragraph")
        };

        // Act
        var alignments = _diffEngine.AlignParagraphs(original, modified);

        // Assert
        alignments.Should().HaveCount(3);
        alignments.Should().Contain(a => a.Type == DiffType.Deleted);
    }

    [Fact]
    public void Compare_SimpleDocuments_ProducesRedlinedDocument()
    {
        // Arrange
        var original = CreateDocument("Original text here");
        var modified = CreateDocument("Modified text here");

        // Act
        var (redlined, statistics) = _diffEngine.Compare(original, modified);

        // Assert
        redlined.Should().NotBeNull();
        redlined.Sections.Should().NotBeEmpty();
        statistics.Should().NotBeNull();
    }

    [Fact]
    public void Compare_WithChanges_TracksStatistics()
    {
        // Arrange
        var original = CreateDocument("Hello world");
        var modified = CreateDocument("Hello beautiful world");

        // Act
        var (_, statistics) = _diffEngine.Compare(original, modified);

        // Assert
        statistics.Insertions.Should().BeGreaterThan(0);
        statistics.OriginalParagraphs.Should().Be(1);
        statistics.ModifiedParagraphs.Should().Be(1);
    }

    private Paragraph CreateParagraph(string text)
    {
        return new Paragraph
        {
            Runs = string.IsNullOrEmpty(text)
                ? new List<Run>()
                : new List<Run> { new Run(text) }
        };
    }

    private Document CreateDocument(string text)
    {
        return new Document
        {
            Sections = new List<Section>
            {
                new Section
                {
                    Blocks = new List<Block>
                    {
                        CreateParagraph(text)
                    }
                }
            }
        };
    }
}
