// Sample usage of DocumentCompare library
//
// This file demonstrates how to compare two Word documents and generate
// a redlined output showing deletions and insertions.

using DocumentCompare.Core.Interfaces;
using DocumentCompare.Word;

namespace DocumentCompare.Samples;

public class SampleUsage
{
    /// <summary>
    /// Basic usage: Compare two Word documents
    /// </summary>
    public static void CompareDocuments()
    {
        // Create a comparer configured for Word documents
        var comparer = WordDocumentComparer.Create();

        // Set up the comparison request
        var request = new CompareRequest
        {
            OriginalDocumentPath = @"C:\Documents\contract_v1.docx",
            ModifiedDocumentPath = @"C:\Documents\contract_v2.docx",
            OutputFormat = OutputFormat.Word,
            OutputPath = @"C:\Documents\contract_redline.docx",
            Options = new CompareOptions
            {
                IgnoreWhitespace = true,
                IgnoreCase = false,
                Granularity = CompareGranularity.Word
            }
        };

        // Perform the comparison
        var result = comparer.Compare(request);

        if (result.Success)
        {
            Console.WriteLine($"Redline document created: {result.OutputPath}");
            Console.WriteLine($"Statistics:");
            Console.WriteLine($"  Insertions: {result.Statistics.Insertions}");
            Console.WriteLine($"  Deletions: {result.Statistics.Deletions}");
            Console.WriteLine($"  Unchanged: {result.Statistics.Unchanged}");
            Console.WriteLine($"  Change percentage: {result.Statistics.ChangePercentage:F1}%");
        }
        else
        {
            Console.WriteLine($"Comparison failed: {result.ErrorMessage}");
        }
    }

    /// <summary>
    /// Advanced usage: Customize redline styles
    /// </summary>
    public static void CompareWithCustomStyles()
    {
        var comparer = WordDocumentComparer.Create();

        var request = new CompareRequest
        {
            OriginalDocumentPath = @"C:\Documents\original.docx",
            ModifiedDocumentPath = @"C:\Documents\modified.docx",
            OutputFormat = OutputFormat.Word,
            Options = new CompareOptions
            {
                Styles = new RedlineStyles
                {
                    // Custom colors (hex without #)
                    DeletionColor = "C00000",      // Dark red
                    InsertionColor = "0070C0",     // Blue
                    MoveColor = "00B050",          // Green

                    // Custom formatting
                    DeletionStrikethrough = true,
                    InsertionBold = true
                }
            }
        };

        var result = comparer.Compare(request);
        // Handle result...
    }

    /// <summary>
    /// Using streams instead of file paths
    /// </summary>
    public static void CompareFromStreams(Stream originalStream, Stream modifiedStream)
    {
        var comparer = WordDocumentComparer.Create();

        var request = new CompareRequest
        {
            OriginalDocumentPath = "original.docx", // Used for format detection
            ModifiedDocumentPath = "modified.docx",
            OriginalStream = originalStream,
            ModifiedStream = modifiedStream,
            OutputFormat = OutputFormat.Word,
            OutputPath = @"C:\Temp\redline_output.docx"
        };

        var result = comparer.Compare(request);
        // Handle result...
    }

    /// <summary>
    /// Direct document model manipulation
    /// </summary>
    public static void WorkWithDocumentModel()
    {
        var parser = new WordParser();
        var generator = new WordGenerator();

        // Parse documents
        var original = parser.Parse(@"C:\Documents\original.docx");
        var modified = parser.Parse(@"C:\Documents\modified.docx");

        // Access document structure
        Console.WriteLine($"Original has {original.Sections.Count} sections");
        Console.WriteLine($"Original has {original.GetAllParagraphs().Count()} paragraphs");

        // Check numbering definitions
        Console.WriteLine($"Numbering definitions: {original.NumberingDefinitions.Count}");
        foreach (var numDef in original.NumberingDefinitions)
        {
            Console.WriteLine($"  Definition {numDef.Id}: {numDef.Levels.Count} levels");
        }

        // Create comparison using the document models
        var comparer = WordDocumentComparer.Create();
        var redlined = comparer.CompareToDocument(original, modified);

        // Generate output
        generator.Generate(redlined, @"C:\Documents\redline.docx");
    }
}
