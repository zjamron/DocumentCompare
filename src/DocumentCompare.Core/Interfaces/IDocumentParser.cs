using DocumentCompare.Core.Models;

namespace DocumentCompare.Core.Interfaces;

/// <summary>
/// Interface for parsing documents from various formats into the internal model.
/// </summary>
public interface IDocumentParser
{
    /// <summary>
    /// Gets the file extensions this parser supports (e.g., ".docx", ".pdf").
    /// </summary>
    IEnumerable<string> SupportedExtensions { get; }

    /// <summary>
    /// Determines if this parser can handle the specified file.
    /// </summary>
    bool CanParse(string filePath);

    /// <summary>
    /// Parses a document from the specified file path.
    /// </summary>
    Document Parse(string filePath);

    /// <summary>
    /// Parses a document from a stream.
    /// </summary>
    Document Parse(Stream stream, string fileName);
}
