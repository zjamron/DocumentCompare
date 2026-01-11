namespace DocumentCompare.Core.Models;

/// <summary>
/// Represents the numbering format for a single level in a numbering definition.
/// </summary>
public class NumberingLevel
{
    /// <summary>
    /// The level index (0-8, where 0 is the top level).
    /// </summary>
    public int Level { get; set; }

    /// <summary>
    /// The number format: decimal, lowerLetter, upperLetter, lowerRoman, upperRoman, bullet, none
    /// </summary>
    public string Format { get; set; } = "decimal";

    /// <summary>
    /// The text pattern, e.g., "%1.", "%1.%2", "(%1)", "Section %1.%2"
    /// %1 = level 1 number, %2 = level 2 number, etc.
    /// </summary>
    public string Text { get; set; } = "%1.";

    /// <summary>
    /// Starting number for this level.
    /// </summary>
    public int Start { get; set; } = 1;

    /// <summary>
    /// Indentation in twips (1/1440 of an inch).
    /// </summary>
    public int? Indent { get; set; }

    /// <summary>
    /// Hanging indent in twips.
    /// </summary>
    public int? HangingIndent { get; set; }

    /// <summary>
    /// Text alignment for the number: left, center, right
    /// </summary>
    public string Alignment { get; set; } = "left";

    /// <summary>
    /// Font for the numbering text.
    /// </summary>
    public string? Font { get; set; }

    public NumberingLevel Clone()
    {
        return new NumberingLevel
        {
            Level = Level,
            Format = Format,
            Text = Text,
            Start = Start,
            Indent = Indent,
            HangingIndent = HangingIndent,
            Alignment = Alignment,
            Font = Font
        };
    }
}

/// <summary>
/// Represents a complete numbering definition (abstract numbering in Word terms).
/// </summary>
public class NumberingDefinition
{
    /// <summary>
    /// Unique identifier for this numbering definition.
    /// </summary>
    public int Id { get; set; }

    /// <summary>
    /// Optional name for the numbering definition.
    /// </summary>
    public string? Name { get; set; }

    /// <summary>
    /// The levels (0-8) defining the numbering format at each depth.
    /// </summary>
    public List<NumberingLevel> Levels { get; set; } = new();

    /// <summary>
    /// Whether this is a multi-level list (true) or single-level (false).
    /// </summary>
    public bool MultiLevel { get; set; }

    public NumberingDefinition Clone()
    {
        return new NumberingDefinition
        {
            Id = Id,
            Name = Name,
            Levels = Levels.Select(l => l.Clone()).ToList(),
            MultiLevel = MultiLevel
        };
    }
}

/// <summary>
/// Represents a numbering instance that references a definition.
/// Multiple paragraphs can share the same NumberingInstance to continue numbering.
/// </summary>
public class NumberingInstance
{
    /// <summary>
    /// Unique identifier for this instance.
    /// </summary>
    public int Id { get; set; }

    /// <summary>
    /// The abstract numbering definition this instance uses.
    /// </summary>
    public int DefinitionId { get; set; }

    /// <summary>
    /// Level overrides for specific levels in this instance.
    /// </summary>
    public Dictionary<int, NumberingLevelOverride> LevelOverrides { get; set; } = new();
}

/// <summary>
/// Override for a specific level in a numbering instance.
/// </summary>
public class NumberingLevelOverride
{
    public int Level { get; set; }

    /// <summary>
    /// Override the starting number at this level.
    /// </summary>
    public int? StartOverride { get; set; }

    /// <summary>
    /// Complete replacement for the level definition.
    /// </summary>
    public NumberingLevel? LevelDefinition { get; set; }
}

/// <summary>
/// References numbering information for a paragraph.
/// </summary>
public class NumberingInfo
{
    /// <summary>
    /// The numbering instance ID this paragraph belongs to.
    /// </summary>
    public int NumberingId { get; set; }

    /// <summary>
    /// The level within the numbering (0-8).
    /// </summary>
    public int Level { get; set; }

    public NumberingInfo Clone()
    {
        return new NumberingInfo
        {
            NumberingId = NumberingId,
            Level = Level
        };
    }
}
