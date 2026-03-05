namespace Nedev.DocToDocx.Models;

/// <summary>
/// Table model
/// </summary>
public class TableModel
{
    public int Index { get; set; }
    public int RowCount { get; set; }
    public int ColumnCount { get; set; }
    public List<TableRowModel> Rows { get; set; } = new();
    public TableProperties? Properties { get; set; }
    public int StartParagraphIndex { get; set; }
    public int EndParagraphIndex { get; set; }
    
    /// <summary>
    /// Current implementation models only top-level tables laid out on the main
    /// document text stream. Nested tables (a table entirely inside a cell of
    /// another table) are flattened into multiple top-level TableModel instances
    /// that share paragraph ranges. This keeps the reader/writer pipeline simple
    /// but means extremely complex nested layouts may not be reproduced exactly.
    /// </summary>
}

/// <summary>
/// Table row model
/// </summary>
public class TableRowModel
{
    public int Index { get; set; }
    public List<TableCellModel> Cells { get; set; } = new();
    public TableRowProperties? Properties { get; set; }
}

/// <summary>
/// Table cell model
/// </summary>
public class TableCellModel
{
    public int Index { get; set; }
    public int RowIndex { get; set; }
    public int ColumnIndex { get; set; }
    public int RowSpan { get; set; } = 1;
    public int ColumnSpan { get; set; } = 1;
    public List<ParagraphModel> Paragraphs { get; set; } = new();
    public TableCellProperties? Properties { get; set; }
}

/// <summary>
/// Table properties (TAP)
/// </summary>
public class TableProperties
{
    public int StyleIndex { get; set; } = -1;
    public int CellSpacing { get; set; }
    public int Indent { get; set; }
    public TableAlignment Alignment { get; set; } = TableAlignment.Left;
    
    /// <summary>
    /// Preferred table width in twips when specified by the TAP. A value of 0
    /// means "auto" and lets Word lay out the table based on content.
    /// </summary>
    public int PreferredWidth { get; set; }
    
    public BorderInfo? BorderTop { get; set; }
    public BorderInfo? BorderBottom { get; set; }
    public BorderInfo? BorderLeft { get; set; }
    public BorderInfo? BorderRight { get; set; }
    public BorderInfo? BorderInsideH { get; set; }
    public BorderInfo? BorderInsideV { get; set; }
    public ShadingInfo? Shading { get; set; }
}

public enum TableAlignment
{
    Left,
    Center,
    Right
}

/// <summary>
/// Table row properties
/// </summary>
public class TableRowProperties
{
    public int Height { get; set; }
    public bool HeightIsExact { get; set; }
    public bool IsHeaderRow { get; set; }
    public bool AllowBreakAcrossPages { get; set; } = true;
}

/// <summary>
/// Table cell properties
/// </summary>
public class TableCellProperties
{
    public int Width { get; set; }
    public BorderInfo? BorderTop { get; set; }
    public BorderInfo? BorderBottom { get; set; }
    public BorderInfo? BorderLeft { get; set; }
    public BorderInfo? BorderRight { get; set; }
    public ShadingInfo? Shading { get; set; }
    public VerticalAlignment VerticalAlignment { get; set; } = VerticalAlignment.Top;
    public bool FitText { get; set; }
    public bool NoWrap { get; set; }
}

public enum VerticalAlignment
{
    Top,
    Center,
    Bottom,
    Both
}

/// <summary>
/// Image model
/// </summary>
public class ImageModel
{
    public string Id { get; set; } = string.Empty;
    public string RelationshipId { get; set; } = string.Empty;
    public string FileName { get; set; } = string.Empty;
    public string ContentType { get; set; } = string.Empty;
    public byte[] Data { get; set; } = Array.Empty<byte>();
    public int Width { get; set; }
    public int Height { get; set; }
    public int WidthEMU { get; set; }
    public int HeightEMU { get; set; }
    public ImageType Type { get; set; }
    public int Left { get; set; }
    public int Top { get; set; }
    public int ScaleX { get; set; } = 100000;
    public int ScaleY { get; set; } = 100000;
    public bool IsLinked { get; set; }
    public string? LinkPath { get; set; }
    public int PictureOffset { get; set; }
}

public enum ImageType
{
    Unknown,
    Wmf,
    Emf,
    Png,
    Jpeg,
    Dib,
    Gif,
    Tiff
}

/// <summary>
/// Style sheet model
/// </summary>
public class StyleSheet
{
    public List<StyleDefinition> Styles { get; set; } = new();
    public List<FontDefinition> Fonts { get; set; } = new();
}

/// <summary>
/// Style definition
/// </summary>
public class StyleDefinition
{
    public ushort StyleId { get; set; }
    public string Name { get; set; } = string.Empty;
    public StyleType Type { get; set; }
    public ParagraphProperties? ParagraphProperties { get; set; }
    public RunProperties? RunProperties { get; set; }
    public ushort? BasedOn { get; set; }
    public ushort? NextParagraphStyle { get; set; }
    public bool IsHidden { get; set; }
    public bool IsQuickStyle { get; set; }
    public int Priority { get; set; }
    public bool IsAutoRedefined { get; set; }
    public bool IsLinked { get; set; }
    public bool IsPrimary { get; set; }
}

public enum StyleType
{
    Paragraph,
    Character,
    Table,
    Numbering
}

/// <summary>
/// Font definition
/// </summary>
public class FontDefinition
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
    public int Family { get; set; }
    public int Charset { get; set; }
    public int Pitch { get; set; }
    public int Type { get; set; }
    public string? AltName { get; set; }
}

/// <summary>
/// List format information (from LSTF structure)
/// </summary>
public class ListFormat
{
    public int ListId { get; set; }
    public ListType Type { get; set; }
    public List<ListLevel> Levels { get; set; } = new();
}

public enum ListType
{
    Bullet,
    Numbered,
    Outline,
    Simple
}

/// <summary>
/// List level information (from LVLF structure)
/// </summary>
public class ListLevel
{
    public int Level { get; set; }
    public NumberFormat NumberFormat { get; set; }
    public string? NumberText { get; set; }
    public int StartAt { get; set; } = 1;
    public int Indent { get; set; }
    public int Space { get; set; }
    public int Alignment { get; set; }
    public byte[]? NumberTextBytes { get; set; }
    public ParagraphProperties? ParagraphProperties { get; set; }
    public RunProperties? RunProperties { get; set; }
}

public enum NumberFormat
{
    None = 0,
    Bullet = 1,
    Decimal = 2,
    LowerRoman = 3,
    UpperRoman = 4,
    LowerLetter = 5,
    UpperLetter = 6,
    Ordinal = 7,
    CardinalText = 8,
    OrdinalText = 9,
    Hex = 10,
    Chicago = 11,
    IdeographDigital = 12,
    JapaneseCounting = 13,
    Aiueo = 14,
    Iroha = 15,
    DecimalFullWidth = 16,
    DecimalHalfWidth = 17,
    JapaneseLegal = 18,
    JapaneseDigitalTenThousand = 19,
    DecimalEnclosedCircle = 20,
    DecimalFullWidth2 = 21,
    AiueoFullWidth = 22,
    IrohaFullWidth = 23,
    DecimalZero = 24,
    Bullet2 = 25,
    Ganada = 26,
    Chosung = 27,
    DecimalEnclosedFullstop = 28,
    DecimalEnclosedParen = 29,
    DecimalEnclosedCircleChinese = 30,
    IdeographEnclosedCircle = 31,
    IdeographTraditional = 32,
    IdeographZodiac = 33,
    IdeographZodiacTraditional = 34,
    TaiwaneseCounting = 35,
    IdeographLegalTraditional = 36,
    TaiwaneseCountingThousand = 37,
    TaiwaneseDigital = 38,
    ChineseCounting = 39,
    ChineseLegalSimplified = 40,
    ChineseLegalTraditional = 41,
    JapaneseCounting2 = 42,
    JapaneseDigitalHundredCount = 43,
    JapaneseDigitalThousandCount = 44,
    // Additional formats used by ListReader nfc mapping
    OrdinalNumber = 45,
    ChineseCountingThousand = 46,
    KoreanDigital = 47,
    KoreanCounting = 48,
    Hebrew1 = 49,
    Hebrew2 = 50,
    ArabicAlpha = 51,
    ArabicAbjad = 52,
    HindiVowels = 53
}

/// <summary>
/// List format override (from LFO structure)
/// </summary>
public class ListFormatOverride
{
    public int ListId { get; set; }
    public int Level { get; set; }
    public int StartAt { get; set; }
}

/// <summary>
/// Numbering definition for DOCX numbering.xml
/// </summary>
public class NumberingDefinition
{
    public int Id { get; set; }
    public List<NumberingLevel> Levels { get; set; } = new();
}

/// <summary>
/// Numbering level definition
/// </summary>
public class NumberingLevel
{
    public int Level { get; set; }
    public NumberFormat NumberFormat { get; set; }
    public string? Text { get; set; }
    public int Start { get; set; } = 1;
    public ParagraphProperties? ParagraphProperties { get; set; }
    public RunProperties? RunProperties { get; set; }
}

/// <summary>
/// Header/Footer model
/// </summary>
public class HeaderFooterModel
{
    public HeaderFooterType Type { get; set; }
    public int SectionIndex { get; set; }
    public string Text { get; set; } = string.Empty;
    public int CharacterPosition { get; set; }
    public int CharacterLength { get; set; }
    public string? RelationshipId { get; set; }
}

/// <summary>
/// Header/Footer types
/// </summary>
public enum HeaderFooterType
{
    HeaderFirst,    // First page header
    FooterFirst,    // First page footer
    HeaderOdd,      // Odd page header (default)
    FooterOdd,      // Odd page footer (default)
    HeaderEven,     // Even page header
    FooterEven      // Even page footer
}

/// <summary>
/// Bookmark model
/// </summary>
public class BookmarkModel
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
    public int StartCp { get; set; }
    public int EndCp { get; set; }
    public bool IsHidden { get; set; }
    public string? RelationshipId { get; set; }
}

/// <summary>
/// Hyperlink model
/// </summary>
public class HyperlinkModel
{
    /// <summary>Target URL or file path</summary>
    public string Url { get; set; } = string.Empty;

    /// <summary>Display text (if different from URL)</summary>
    public string? DisplayText { get; set; }

    /// <summary>Bookmark anchor within the target</summary>
    public string? Bookmark { get; set; }

    /// <summary>Screen tip text</summary>
    public string? ScreenTip { get; set; }

    /// <summary>Is this an external link (vs. internal bookmark)</summary>
    public bool IsExternal { get; set; }

    /// <summary>Relationship ID for DOCX</summary>
    public string? RelationshipId { get; set; }

    /// <summary>Target frame for web links</summary>
    public string? TargetFrame { get; set; }
}

/// <summary>
/// Field model
/// </summary>
public class FieldModel
{
    public FieldType Type { get; set; }
    public string RawCode { get; set; } = string.Empty;
    public string FieldName { get; set; } = string.Empty;
    public string Arguments { get; set; } = string.Empty;
    public Dictionary<string, string> Switches { get; set; } = new();
    public string? Result { get; set; }
    public bool IsLocked { get; set; }
    public bool IsDirty { get; set; }
}

/// <summary>
/// Field types
/// </summary>
public enum FieldType
{
    Unknown,
    // Page numbers
    PageNumber,
    NumPages,
    SectionNumber,
    // Dates and times
    Date,
    Time,
    CreateDate,
    SaveDate,
    PrintDate,
    EditTime,
    // Document properties
    Author,
    Title,
    Subject,
    Keywords,
    Comments,
    FileName,
    Template,
    DocProperty,
    // User info
    UserName,
    UserInitials,
    UserAddress,
    // Links and references
    Hyperlink,
    Reference,
    PageReference,
    Bookmark,
    StyleReference,
    // Document structure
    TableOfContents,
    Index,
    IndexEntry,
    TocEntry,
    // Special
    Sequence,
    Ask,
    FillIn,
    MergeField,
    If,
    Compare,
    Formula,
    Quote,
    Symbol,
    Embed,
    Link,
    IncludeText,
    IncludePicture
}
