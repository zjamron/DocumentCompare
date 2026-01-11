namespace DocumentCompare.Core.Models;

/// <summary>
/// Represents a section of the document with its own page layout and headers/footers.
/// </summary>
public class Section
{
    /// <summary>
    /// The block-level content of this section.
    /// </summary>
    public List<Block> Blocks { get; set; } = new();

    /// <summary>
    /// Headers for this section (first page, odd pages, even pages).
    /// </summary>
    public HeaderFooterSet Headers { get; set; } = new();

    /// <summary>
    /// Footers for this section.
    /// </summary>
    public HeaderFooterSet Footers { get; set; } = new();

    /// <summary>
    /// Page layout properties for this section.
    /// </summary>
    public SectionProperties Properties { get; set; } = new();

    public Section Clone()
    {
        return new Section
        {
            Blocks = Blocks.Select(b => b.Clone()).ToList(),
            Headers = Headers.Clone(),
            Footers = Footers.Clone(),
            Properties = Properties.Clone()
        };
    }
}

/// <summary>
/// Collection of headers or footers for different page types.
/// </summary>
public class HeaderFooterSet
{
    /// <summary>
    /// Default header/footer (used for all pages unless overridden).
    /// </summary>
    public HeaderFooter? Default { get; set; }

    /// <summary>
    /// First page header/footer (if different from default).
    /// </summary>
    public HeaderFooter? First { get; set; }

    /// <summary>
    /// Even page header/footer (for different odd/even page headers).
    /// </summary>
    public HeaderFooter? Even { get; set; }

    public HeaderFooterSet Clone()
    {
        return new HeaderFooterSet
        {
            Default = Default?.Clone(),
            First = First?.Clone(),
            Even = Even?.Clone()
        };
    }
}

/// <summary>
/// Represents a header or footer containing block-level content.
/// </summary>
public class HeaderFooter
{
    public List<Block> Blocks { get; set; } = new();

    public HeaderFooter Clone()
    {
        return new HeaderFooter
        {
            Blocks = Blocks.Select(b => b.Clone()).ToList()
        };
    }
}

/// <summary>
/// Page layout and section break properties.
/// </summary>
public class SectionProperties
{
    /// <summary>
    /// Page width in twips.
    /// </summary>
    public int PageWidth { get; set; } = 12240; // 8.5 inches

    /// <summary>
    /// Page height in twips.
    /// </summary>
    public int PageHeight { get; set; } = 15840; // 11 inches

    /// <summary>
    /// Page orientation: portrait or landscape.
    /// </summary>
    public string Orientation { get; set; } = "portrait";

    /// <summary>
    /// Left margin in twips.
    /// </summary>
    public int MarginLeft { get; set; } = 1440; // 1 inch

    /// <summary>
    /// Right margin in twips.
    /// </summary>
    public int MarginRight { get; set; } = 1440;

    /// <summary>
    /// Top margin in twips.
    /// </summary>
    public int MarginTop { get; set; } = 1440;

    /// <summary>
    /// Bottom margin in twips.
    /// </summary>
    public int MarginBottom { get; set; } = 1440;

    /// <summary>
    /// Header distance from top of page in twips.
    /// </summary>
    public int HeaderDistance { get; set; } = 720; // 0.5 inch

    /// <summary>
    /// Footer distance from bottom of page in twips.
    /// </summary>
    public int FooterDistance { get; set; } = 720;

    /// <summary>
    /// Section break type: continuous, nextPage, evenPage, oddPage
    /// </summary>
    public string SectionBreakType { get; set; } = "nextPage";

    /// <summary>
    /// Whether to use different first page header/footer.
    /// </summary>
    public bool DifferentFirstPage { get; set; }

    /// <summary>
    /// Whether to use different odd/even page headers/footers.
    /// </summary>
    public bool DifferentOddEven { get; set; }

    public SectionProperties Clone()
    {
        return new SectionProperties
        {
            PageWidth = PageWidth,
            PageHeight = PageHeight,
            Orientation = Orientation,
            MarginLeft = MarginLeft,
            MarginRight = MarginRight,
            MarginTop = MarginTop,
            MarginBottom = MarginBottom,
            HeaderDistance = HeaderDistance,
            FooterDistance = FooterDistance,
            SectionBreakType = SectionBreakType,
            DifferentFirstPage = DifferentFirstPage,
            DifferentOddEven = DifferentOddEven
        };
    }
}
