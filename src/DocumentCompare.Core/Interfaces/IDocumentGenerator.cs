using DocumentCompare.Core.Models;

namespace DocumentCompare.Core.Interfaces;

/// <summary>
/// Interface for generating documents in various output formats.
/// </summary>
public interface IDocumentGenerator
{
    /// <summary>
    /// Gets the output format this generator produces (e.g., "docx", "pdf", "html").
    /// </summary>
    string OutputFormat { get; }

    /// <summary>
    /// Generates a document to the specified file path.
    /// </summary>
    void Generate(Document document, string outputPath);

    /// <summary>
    /// Generates a document to a stream.
    /// </summary>
    void Generate(Document document, Stream outputStream);
}
