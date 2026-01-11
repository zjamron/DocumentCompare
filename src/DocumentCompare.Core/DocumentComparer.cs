using DocumentCompare.Core.Comparison;
using DocumentCompare.Core.Interfaces;
using DocumentCompare.Core.Models;

namespace DocumentCompare.Core;

/// <summary>
/// Main service for comparing documents and generating redlined output.
/// </summary>
public class DocumentComparer : IDocumentComparer
{
    private readonly List<IDocumentParser> _parsers;
    private readonly List<IDocumentGenerator> _generators;

    public DocumentComparer(IEnumerable<IDocumentParser> parsers, IEnumerable<IDocumentGenerator> generators)
    {
        _parsers = parsers.ToList();
        _generators = generators.ToList();
    }

    /// <summary>
    /// Compares two documents and generates a redlined output file.
    /// </summary>
    public CompareResult Compare(CompareRequest request)
    {
        try
        {
            // Parse original document
            Document original;
            if (request.OriginalStream != null)
            {
                var parser = GetParser(request.OriginalDocumentPath);
                original = parser.Parse(request.OriginalStream, request.OriginalDocumentPath);
            }
            else
            {
                var parser = GetParser(request.OriginalDocumentPath);
                original = parser.Parse(request.OriginalDocumentPath);
            }

            // Parse modified document
            Document modified;
            if (request.ModifiedStream != null)
            {
                var parser = GetParser(request.ModifiedDocumentPath);
                modified = parser.Parse(request.ModifiedStream, request.ModifiedDocumentPath);
            }
            else
            {
                var parser = GetParser(request.ModifiedDocumentPath);
                modified = parser.Parse(request.ModifiedDocumentPath);
            }

            // Perform comparison
            var (redlinedDocument, statistics) = CompareDocuments(original, modified, request.Options);

            // Generate output
            var generator = GetGenerator(request.OutputFormat);
            var outputPath = request.OutputPath ?? GenerateOutputPath(request.OriginalDocumentPath, request.OutputFormat);

            generator.Generate(redlinedDocument, outputPath);

            return new CompareResult
            {
                Success = true,
                OutputPath = outputPath,
                RedlinedDocument = redlinedDocument,
                Statistics = statistics
            };
        }
        catch (Exception ex)
        {
            return new CompareResult
            {
                Success = false,
                ErrorMessage = ex.Message
            };
        }
    }

    /// <summary>
    /// Compares two documents and returns the redlined document model.
    /// </summary>
    public Document CompareToDocument(Document original, Document modified, CompareOptions? options = null)
    {
        var (redlinedDocument, _) = CompareDocuments(original, modified, options ?? new CompareOptions());
        return redlinedDocument;
    }

    /// <summary>
    /// Internal method to perform the actual comparison.
    /// </summary>
    private (Document RedlinedDocument, CompareStatistics Statistics) CompareDocuments(
        Document original,
        Document modified,
        CompareOptions options)
    {
        var diffEngine = new DiffEngine(options);
        return diffEngine.Compare(original, modified);
    }

    /// <summary>
    /// Gets a parser that can handle the given file.
    /// </summary>
    private IDocumentParser GetParser(string filePath)
    {
        var parser = _parsers.FirstOrDefault(p => p.CanParse(filePath));
        if (parser == null)
        {
            throw new NotSupportedException($"No parser available for file: {filePath}");
        }
        return parser;
    }

    /// <summary>
    /// Gets a generator for the specified output format.
    /// </summary>
    private IDocumentGenerator GetGenerator(OutputFormat format)
    {
        var formatString = format switch
        {
            OutputFormat.Word => "docx",
            OutputFormat.Pdf => "pdf",
            OutputFormat.Html => "html",
            _ => throw new NotSupportedException($"Unsupported output format: {format}")
        };

        var generator = _generators.FirstOrDefault(g =>
            g.OutputFormat.Equals(formatString, StringComparison.OrdinalIgnoreCase));

        if (generator == null)
        {
            throw new NotSupportedException($"No generator available for format: {format}");
        }

        return generator;
    }

    /// <summary>
    /// Generates a default output path based on input and format.
    /// </summary>
    private string GenerateOutputPath(string inputPath, OutputFormat format)
    {
        var directory = Path.GetDirectoryName(inputPath) ?? Environment.CurrentDirectory;
        var baseName = Path.GetFileNameWithoutExtension(inputPath);
        var extension = format switch
        {
            OutputFormat.Word => ".docx",
            OutputFormat.Pdf => ".pdf",
            OutputFormat.Html => ".html",
            _ => ".docx"
        };

        return Path.Combine(directory, $"{baseName}_redline{extension}");
    }
}

/// <summary>
/// Factory for creating DocumentComparer instances with default configuration.
/// </summary>
public static class DocumentComparerFactory
{
    /// <summary>
    /// Creates a DocumentComparer with Word document support.
    /// </summary>
    public static IDocumentComparer Create()
    {
        // This will be populated with actual parsers/generators when the Word assembly is loaded
        return new DocumentComparer(
            new List<IDocumentParser>(),
            new List<IDocumentGenerator>()
        );
    }

    /// <summary>
    /// Creates a DocumentComparer with the specified parsers and generators.
    /// </summary>
    public static IDocumentComparer Create(
        IEnumerable<IDocumentParser> parsers,
        IEnumerable<IDocumentGenerator> generators)
    {
        return new DocumentComparer(parsers, generators);
    }
}
