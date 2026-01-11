using System.Text.RegularExpressions;
using DiffPlex;
using DiffPlex.DiffBuilder;
using DiffPlex.DiffBuilder.Model;
using DocumentCompare.Core.Interfaces;
using DocumentCompare.Core.Models;

namespace DocumentCompare.Core.Comparison;

/// <summary>
/// Core diff engine that compares documents and generates redlined output.
/// </summary>
public class DiffEngine
{
    private readonly CompareOptions _options;
    private readonly Differ _differ;

    public DiffEngine(CompareOptions? options = null)
    {
        _options = options ?? new CompareOptions();
        _differ = new Differ();
    }

    /// <summary>
    /// Compares two documents and produces a redlined document.
    /// </summary>
    public (Document RedlinedDocument, CompareStatistics Statistics) Compare(Document original, Document modified)
    {
        var statistics = new CompareStatistics
        {
            OriginalParagraphs = original.GetAllParagraphs().Count(),
            ModifiedParagraphs = modified.GetAllParagraphs().Count()
        };

        // Create the redlined document based on modified structure
        var redlined = CreateRedlinedDocument(original, modified, statistics);

        return (redlined, statistics);
    }

    /// <summary>
    /// Creates a redlined document by merging original and modified content.
    /// </summary>
    private Document CreateRedlinedDocument(Document original, Document modified, CompareStatistics statistics)
    {
        var redlined = new Document
        {
            Properties = modified.Properties.Clone(),
            NumberingDefinitions = modified.NumberingDefinitions.Select(n => n.Clone()).ToList(),
            NumberingInstances = modified.NumberingInstances.Select(n => new NumberingInstance
            {
                Id = n.Id,
                DefinitionId = n.DefinitionId,
                LevelOverrides = new Dictionary<int, NumberingLevelOverride>(n.LevelOverrides)
            }).ToList(),
            Styles = modified.Styles.Select(s => s.Clone()).ToList()
        };

        // Get paragraphs from both documents
        var originalParagraphs = original.GetAllParagraphs().ToList();
        var modifiedParagraphs = modified.GetAllParagraphs().ToList();

        // Align paragraphs using LCS
        var alignments = AlignParagraphs(originalParagraphs, modifiedParagraphs);

        // Create redlined content
        var redlinedBlocks = new List<Block>();

        foreach (var alignment in alignments)
        {
            if (alignment.Type == DiffType.Deleted)
            {
                // Paragraph was deleted - show it with deletion formatting
                var para = originalParagraphs[alignment.OriginalIndex];
                var deletedPara = CreateDeletedParagraph(para);
                redlinedBlocks.Add(deletedPara);
                statistics.Deletions += CountWords(para.GetPlainText());
            }
            else if (alignment.Type == DiffType.Inserted)
            {
                // Paragraph was inserted - show it with insertion formatting
                var para = modifiedParagraphs[alignment.ModifiedIndex];
                var insertedPara = CreateInsertedParagraph(para);
                redlinedBlocks.Add(insertedPara);
                statistics.Insertions += CountWords(para.GetPlainText());
            }
            else
            {
                // Paragraph exists in both - do word-level diff
                var originalPara = originalParagraphs[alignment.OriginalIndex];
                var modifiedPara = modifiedParagraphs[alignment.ModifiedIndex];
                var diffResult = DiffParagraphs(originalPara, modifiedPara);
                var redlinedPara = CreateRedlinedParagraph(modifiedPara, diffResult);
                redlinedBlocks.Add(redlinedPara);

                statistics.Insertions += diffResult.InsertionCount;
                statistics.Deletions += diffResult.DeletionCount;
                statistics.Unchanged += diffResult.UnchangedCount;
            }
        }

        // Create a single section with all redlined content
        redlined.Sections.Add(new Section
        {
            Blocks = redlinedBlocks,
            Properties = modified.Sections.FirstOrDefault()?.Properties.Clone() ?? new SectionProperties(),
            Headers = modified.Sections.FirstOrDefault()?.Headers.Clone() ?? new HeaderFooterSet(),
            Footers = modified.Sections.FirstOrDefault()?.Footers.Clone() ?? new HeaderFooterSet()
        });

        return redlined;
    }

    /// <summary>
    /// Aligns paragraphs between original and modified documents using LCS algorithm.
    /// </summary>
    public List<ParagraphAlignment> AlignParagraphs(List<Paragraph> original, List<Paragraph> modified)
    {
        var result = new List<ParagraphAlignment>();

        // Build similarity matrix
        var lcs = ComputeLCS(original, modified);

        // Backtrack to find alignment
        int i = original.Count;
        int j = modified.Count;
        var alignments = new Stack<ParagraphAlignment>();

        while (i > 0 || j > 0)
        {
            if (i > 0 && j > 0 && AreParagraphsSimilar(original[i - 1], modified[j - 1]))
            {
                alignments.Push(new ParagraphAlignment
                {
                    OriginalIndex = i - 1,
                    ModifiedIndex = j - 1,
                    Type = DiffType.Unchanged,
                    SimilarityScore = CalculateSimilarity(original[i - 1].GetNormalizedText(), modified[j - 1].GetNormalizedText())
                });
                i--;
                j--;
            }
            else if (j > 0 && (i == 0 || lcs[i, j - 1] >= lcs[i - 1, j]))
            {
                alignments.Push(new ParagraphAlignment
                {
                    OriginalIndex = -1,
                    ModifiedIndex = j - 1,
                    Type = DiffType.Inserted
                });
                j--;
            }
            else
            {
                alignments.Push(new ParagraphAlignment
                {
                    OriginalIndex = i - 1,
                    ModifiedIndex = -1,
                    Type = DiffType.Deleted
                });
                i--;
            }
        }

        while (alignments.Count > 0)
        {
            result.Add(alignments.Pop());
        }

        return result;
    }

    /// <summary>
    /// Computes the LCS table for paragraph alignment.
    /// </summary>
    private int[,] ComputeLCS(List<Paragraph> original, List<Paragraph> modified)
    {
        int m = original.Count;
        int n = modified.Count;
        var lcs = new int[m + 1, n + 1];

        for (int i = 1; i <= m; i++)
        {
            for (int j = 1; j <= n; j++)
            {
                if (AreParagraphsSimilar(original[i - 1], modified[j - 1]))
                {
                    lcs[i, j] = lcs[i - 1, j - 1] + 1;
                }
                else
                {
                    lcs[i, j] = Math.Max(lcs[i - 1, j], lcs[i, j - 1]);
                }
            }
        }

        return lcs;
    }

    /// <summary>
    /// Determines if two paragraphs are similar enough to be considered a match.
    /// </summary>
    private bool AreParagraphsSimilar(Paragraph a, Paragraph b)
    {
        var textA = a.GetNormalizedText();
        var textB = b.GetNormalizedText();

        // Empty paragraphs match each other
        if (string.IsNullOrWhiteSpace(textA) && string.IsNullOrWhiteSpace(textB))
            return true;

        // If one is empty and other isn't, they don't match
        if (string.IsNullOrWhiteSpace(textA) || string.IsNullOrWhiteSpace(textB))
            return false;

        // Calculate similarity
        var similarity = CalculateSimilarity(textA, textB);
        return similarity >= 0.5; // 50% similarity threshold
    }

    /// <summary>
    /// Calculates similarity between two strings (0-1).
    /// </summary>
    private double CalculateSimilarity(string a, string b)
    {
        if (_options.IgnoreCase)
        {
            a = a.ToLowerInvariant();
            b = b.ToLowerInvariant();
        }

        if (a == b) return 1.0;
        if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b)) return 0.0;

        // Use word-based Jaccard similarity
        var wordsA = TokenizeToWords(a).ToHashSet();
        var wordsB = TokenizeToWords(b).ToHashSet();

        if (wordsA.Count == 0 && wordsB.Count == 0) return 1.0;
        if (wordsA.Count == 0 || wordsB.Count == 0) return 0.0;

        var intersection = wordsA.Intersect(wordsB).Count();
        var union = wordsA.Union(wordsB).Count();

        return (double)intersection / union;
    }

    /// <summary>
    /// Performs word-level diff between two paragraphs.
    /// </summary>
    public ParagraphDiffResult DiffParagraphs(Paragraph original, Paragraph modified)
    {
        var originalText = original.GetPlainText();
        var modifiedText = modified.GetPlainText();

        var result = new ParagraphDiffResult();

        // Handle empty cases
        if (string.IsNullOrEmpty(originalText) && string.IsNullOrEmpty(modifiedText))
        {
            return result;
        }

        if (string.IsNullOrEmpty(originalText))
        {
            result.IsEntirelyInserted = true;
            result.Segments.Add(new DiffSegment { Text = modifiedText, Type = DiffType.Inserted });
            return result;
        }

        if (string.IsNullOrEmpty(modifiedText))
        {
            result.IsEntirelyDeleted = true;
            result.Segments.Add(new DiffSegment { Text = originalText, Type = DiffType.Deleted });
            return result;
        }

        // Tokenize to words
        var originalWords = TokenizeToWords(originalText).ToArray();
        var modifiedWords = TokenizeToWords(modifiedText).ToArray();

        // Use DiffPlex for word-level diff
        var diffBuilder = new InlineDiffBuilder(_differ);
        var diff = diffBuilder.BuildDiffModel(
            string.Join(" ", originalWords),
            string.Join(" ", modifiedWords));

        foreach (var line in diff.Lines)
        {
            if (string.IsNullOrEmpty(line.Text)) continue;

            var segment = new DiffSegment
            {
                Text = line.Text + " ",
                Type = line.Type switch
                {
                    ChangeType.Inserted => DiffType.Inserted,
                    ChangeType.Deleted => DiffType.Deleted,
                    _ => DiffType.Unchanged
                }
            };
            result.Segments.Add(segment);
        }

        // Clean up trailing space on last segment
        if (result.Segments.Count > 0)
        {
            var last = result.Segments[^1];
            last.Text = last.Text.TrimEnd();
        }

        return result;
    }

    /// <summary>
    /// Tokenizes text into words, preserving whitespace for reconstruction.
    /// </summary>
    private IEnumerable<string> TokenizeToWords(string text)
    {
        // Split on whitespace while keeping words
        var matches = Regex.Matches(text, @"\S+");
        foreach (Match match in matches)
        {
            yield return match.Value;
        }
    }

    /// <summary>
    /// Creates a paragraph showing deleted content with deletion formatting.
    /// </summary>
    private Paragraph CreateDeletedParagraph(Paragraph original)
    {
        var deleted = (Paragraph)original.Clone();

        // Apply deletion formatting to all runs
        foreach (var run in deleted.Runs)
        {
            run.Formatting = RunFormatting.ForDeletion(run.Formatting);
        }

        return deleted;
    }

    /// <summary>
    /// Creates a paragraph showing inserted content with insertion formatting.
    /// </summary>
    private Paragraph CreateInsertedParagraph(Paragraph modified)
    {
        var inserted = (Paragraph)modified.Clone();

        // Apply insertion formatting to all runs
        foreach (var run in inserted.Runs)
        {
            run.Formatting = RunFormatting.ForInsertion(run.Formatting);
        }

        return inserted;
    }

    /// <summary>
    /// Creates a redlined paragraph by merging diff segments with appropriate formatting.
    /// </summary>
    private Paragraph CreateRedlinedParagraph(Paragraph modified, ParagraphDiffResult diffResult)
    {
        var redlined = new Paragraph
        {
            Style = modified.Style.Clone(),
            Numbering = modified.Numbering?.Clone(),
            BookmarkStarts = new List<string>(modified.BookmarkStarts),
            BookmarkEnds = new List<string>(modified.BookmarkEnds)
        };

        // Build runs from diff segments
        foreach (var segment in diffResult.Segments)
        {
            if (string.IsNullOrEmpty(segment.Text)) continue;

            var formatting = segment.Type switch
            {
                DiffType.Deleted => RunFormatting.ForDeletion(),
                DiffType.Inserted => RunFormatting.ForInsertion(),
                DiffType.MovedFrom => RunFormatting.ForMove(isSource: true),
                DiffType.MovedTo => RunFormatting.ForMove(isSource: false),
                _ => new RunFormatting()
            };

            redlined.Runs.Add(new Run(segment.Text, formatting));
        }

        return redlined;
    }

    /// <summary>
    /// Counts words in text.
    /// </summary>
    private int CountWords(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return 0;
        return TokenizeToWords(text).Count();
    }
}
