using DocumentCompare.Core;
using DocumentCompare.Core.Interfaces;

namespace DocumentCompare.Word;

/// <summary>
/// Factory for creating a DocumentComparer with Word document support.
/// </summary>
public static class WordDocumentComparer
{
    /// <summary>
    /// Creates a DocumentComparer configured for Word document comparison.
    /// </summary>
    public static IDocumentComparer Create()
    {
        var parsers = new IDocumentParser[] { new WordParser() };
        var generators = new IDocumentGenerator[] { new WordGenerator() };

        return DocumentComparerFactory.Create(parsers, generators);
    }
}
