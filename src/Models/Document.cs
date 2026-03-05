using System.Text;
using Nedev.DocToDocx.Utils;

namespace Nedev.DocToDocx.Models;

/// <summary>
/// Represents the main document structure
/// </summary>
public class DocumentModel
{
    public List<ParagraphModel> Paragraphs { get; set; } = new();
    public List<TableModel> Tables { get; set; } = new();
    public List<ImageModel> Images { get; set; } = new();
    public List<ShapeModel> Shapes { get; set; } = new();
    /// <summary>
    /// High-level chart objects extracted from OLE/OfficeArt structures.
    /// For the initial implementation charts may carry only minimal metadata
    /// and placeholder data series but are emitted as real DOCX charts so
    /// they can be edited inside Word.
    /// </summary>
    public List<ChartModel> Charts { get; set; } = new();
    public List<BookmarkModel> Bookmarks { get; set; } = new();
    public List<HyperlinkModel> Hyperlinks { get; set; } = new();
    public List<FootnoteModel> Footnotes { get; set; } = new();
    public List<EndnoteModel> Endnotes { get; set; } = new();
    public List<AnnotationModel> Annotations { get; set; } = new();
    public List<TextboxModel> Textboxes { get; set; } = new();
    public List<OleObjectModel> OleObjects { get; set; } = new();
    public byte[]? VbaProject { get; set; }
    public StyleSheet Styles { get; set; } = new();
    public DocumentProperties Properties { get; set; } = new();
    public HeaderFooterInfo HeadersFooters { get; set; } = new();
    public List<NumberingDefinition> NumberingDefinitions { get; set; } = new();
    public List<ListFormat> ListFormats { get; set; } = new();
}

/// <summary>
/// Document-level properties
/// </summary>
public class DocumentProperties
{
    public DateTime Created { get; set; }
    public DateTime Modified { get; set; }
    public int PageWidth { get; set; } = 12240; // Default 8.5" in twips
    public int PageHeight { get; set; } = 15840; // Default 11" in twips
    public int MarginTop { get; set; } = 1440; // 1" in twips
    public int MarginBottom { get; set; } = 1440;
    public int MarginLeft { get; set; } = 1440;
    public int MarginRight { get; set; } = 1440;
    public int SectionStartPageNumber { get; set; } = 1;
    public bool IsLandscape { get; set; }
    public List<SectionInfo> Sections { get; set; } = new();

    // Document metadata for fields
    public string? Author { get; set; }
    public string? Title { get; set; }
    public string? Subject { get; set; }
    public string? Keywords { get; set; }
    public string? Comments { get; set; }
    public string? FileName { get; set; }
    public string? Template { get; set; }

    // DOP flags (from MS-DOC §2.7.4 Dop97)
    public bool FWidowControl { get; set; } // fWidowControl
    public bool FPaginated { get; set; } // fPaginated
    public bool FFacingPages { get; set; } // fFacingPages
    public bool FBreaks { get; set; } // fBreaks
    public bool FAutoHyphenate { get; set; } // fAutoHyphenate
    public bool FDoHyphenation { get; set; } // fDoHyphenation
    public bool FFELayout { get; set; } // fFELayout
    public bool FLayoutSameAsWin95 { get; set; } // fLayoutSameAsWin95
    public bool FPrintBodyBeforeHeaders { get; set; } // fPrintBodyBeforeHeaders
    public bool FSuppressBottomSpacing { get; set; } // fSuppressBottomSpacing
    public bool FWrapAuto { get; set; } // fWrapAuto
    public bool FPrintPaperBefore { get; set; } // fPrintPaperBefore
    public bool FSuppressSpacings { get; set; } // fSuppressSpacings
    public bool FMirrorMargins { get; set; } // fMirrorMargins
    public bool FUsePrinterMetrics { get; set; } // fUsePrinterMetrics
    public bool FNoPgp { get; set; } // fNoPgp
    public bool FShrinkToFit { get; set; } // fShrinkToFit
    public bool FPrintFormsData { get; set; } // fPrintFormsData
    public bool FAllowPositionOnOnly { get; set; } // fAllowPositionOnOnly
    public bool FDisplayBackground { get; set; } // fDisplayBackground
    public bool FDisplayLineNumbers { get; set; } // fDisplayLineNumbers
    public bool FPrintMicros { get; set; } // fPrintMicros
    public bool FSaveFormsData { get; set; } // fSaveFormsData
    public bool FDisplayColBreak { get; set; } // fDisplayColBreak
    public bool FDisplayPageEnd { get; set; } // fDisplayPageEnd
    public bool FDisplayUnits { get; set; } // fDisplayUnits
    public bool FProtectForms { get; set; } // fProtectForms
    public bool FProtectSparce { get; set; } // fProtectSparce
    public bool FConsecutiveHyphen { get; set; } // fConsecutiveHyphen
    public bool FLetterFinal { get; set; } // fLetterFinal
    public bool FLetterSparce { get; set; } // fLetterSparce
    public bool FLinePrint { get; set; } // fLinePrint
    public bool FSubFontOnDoc { get; set; } // fSubFontOnDoc
    public bool FNoLeading { get; set; } // fNoLeading
    public bool FMScript { get; set; } // fMScript
    public bool FOutlineMode { get; set; } // fOutlineMode
    public bool FLayoutInCell { get; set; } // fLayoutInCell
    public bool FKeyBoard { get; set; } // fKeyBoard
    public bool FSameFont { get; set; } // fSameFont
    public bool FEmbedTrueTypeFonts { get; set; } // fEmbedTrueTypeFonts
    public bool FSaveRGB { get; set; } // fSaveRGB
    public bool FNoSuperscript { get; set; } // fNoSuperscript
    public int DxaTab { get; set; } = 288; // Default 2" in twips
    public int DxaColumns { get; set; } // dxaColumns
    public int ITxtWrap { get; set; } // iTxtWrap
}

/// <summary>
/// Section information
/// </summary>
public class SectionInfo
{
    public int StartParagraphIndex { get; set; }
    public int PageWidth { get; set; }
    public int PageHeight { get; set; }
    public int MarginTop { get; set; }
    public int MarginBottom { get; set; }
    public int MarginLeft { get; set; }
    public int MarginRight { get; set; }
    public bool IsLandscape { get; set; }
    public HeaderFooterReferenceType HeaderReference { get; set; }
    public HeaderFooterReferenceType FooterReference { get; set; }
}

public enum HeaderFooterReferenceType
{
    None,
    Default,
    First,
    Even
}

/// <summary>
/// Header and footer information
/// </summary>
public class HeaderFooterInfo
{
    public string? DefaultHeader { get; set; }
    public string? DefaultFooter { get; set; }
    public string? FirstPageHeader { get; set; }
    public string? FirstPageFooter { get; set; }
    public string? EvenPageHeader { get; set; }
    public string? EvenPageFooter { get; set; }
    
    /// <summary>Header/footer models for detailed processing</summary>
    public List<HeaderFooterModel> Headers { get; set; } = new();
    public List<HeaderFooterModel> Footers { get; set; } = new();
}

/// <summary>
/// Paragraph model
/// </summary>
public class ParagraphModel
{
    public int Index { get; set; }
    public List<RunModel> Runs { get; set; } = new();
    public ParagraphProperties? Properties { get; set; }
    public ParagraphType Type { get; set; } = ParagraphType.Normal;
    public int TableRowIndex { get; set; } = -1;
    public int TableCellIndex { get; set; } = -1;
    public int NestingLevel { get; set; }
    
    /// <summary>Set when Type == ParagraphType.NestedTable</summary>
    public TableModel? NestedTable { get; set; }
    
    /// <summary>List format ID (ilfo) - 0 if not in a list</summary>
    public int ListFormatId { get; set; }
    
    /// <summary>List level (ilvl) - 0-8 for list levels</summary>
    public int ListLevel { get; set; }
    
    /// <summary>Is this paragraph part of a numbered list</summary>
    public bool IsNumberedList => ListFormatId > 0;
    
    /// <summary>
    /// Gets the text content of this paragraph by combining all runs
    /// </summary>
    public string Text => string.Join("", Runs.Select(r => r.Text));
}

public enum ParagraphType
{
    Normal,
    TableRow,
    TableCell,
    BookmarkStart,
    BookmarkEnd,
    Heading,
    PageBreak,
    SectionBreak,
    NestedTable
}

/// <summary>
/// Run (text segment) model
/// </summary>
public class RunModel
{
    public string Text { get; set; } = string.Empty;
    public RunProperties? Properties { get; set; }
    public bool IsField { get; set; }
    public string? FieldCode { get; set; }
    public int CharacterPosition { get; set; }
    public int CharacterLength { get; set; }

    /// <summary>Is this run a picture/image</summary>
    public bool IsPicture { get; set; }

    /// <summary>Image index in document.Images list</summary>
    public int ImageIndex { get; set; } = -1;

    /// <summary>File character offset in Data stream for picture (from sprmCPicLocation).</summary>
    public uint FcPic { get; set; }

    /// <summary>Image relationship ID for DOCX</summary>
    public string? ImageRelationshipId { get; set; }

    /// <summary>Is this run an OLE object preview</summary>
    public bool IsOle { get; set; }

    /// <summary>OLE Object ID (matches ObjectId in OleObjectModel)</summary>
    public string? OleObjectId { get; set; }

    /// <summary>OLE Program ID (e.g. Excel.Sheet.8)</summary>
    public string? OleProgId { get; set; }

    /// <summary>Is this run a hyperlink</summary>
    public bool IsHyperlink { get; set; }

    /// <summary>Hyperlink URL or bookmark reference</summary>
    public string? HyperlinkUrl { get; set; }

    /// <summary>Hyperlink relationship ID for DOCX</summary>
    public string? HyperlinkRelationshipId { get; set; }

    /// <summary>Is this run part of a bookmark</summary>
    public bool IsBookmark { get; set; }

    /// <summary>Bookmark name if this is a bookmark start/end</summary>
    public string? BookmarkName { get; set; }

    /// <summary>Is this a bookmark start (vs. end)</summary>
    public bool IsBookmarkStart { get; set; }
}

/// <summary>
/// Paragraph properties (PAP)
/// </summary>
public class ParagraphProperties
{
    public int StyleIndex { get; set; } = -1;
    public ParagraphAlignment Alignment { get; set; } = ParagraphAlignment.Left;
    public int IndentLeft { get; set; }
    public int IndentRight { get; set; }
    public int IndentFirstLine { get; set; }
    public int SpaceBefore { get; set; }
    public int SpaceAfter { get; set; }
    public int LineSpacing { get; set; } = 240;
    public int LineSpacingMultiple { get; set; } = 1;
    public bool KeepWithNext { get; set; }
    public bool KeepTogether { get; set; }
    public bool PageBreakBefore { get; set; }
    public BorderInfo? BorderTop { get; set; }
    public BorderInfo? BorderBottom { get; set; }
    public BorderInfo? BorderLeft { get; set; }
    public BorderInfo? BorderRight { get; set; }
    public ShadingInfo? Shading { get; set; }
    
    // List properties
    public int ListFormatId { get; set; }
    public int ListLevel { get; set; }
    public int OutlineLevel { get; set; }
    public NumberFormat? NumberFormat { get; set; }
    public string? NumberText { get; set; }
    
    // Phase 1 Additions (Typography & Layout)
    public bool SnapToGrid { get; set; } = true; // Default true in Word
    public bool AutoSpaceDe { get; set; } = true; // Adjust space between Asian and Latin text
    public bool AutoSpaceDn { get; set; } = true; // Adjust space between Asian text and numbers
    public bool WordWrap { get; set; } = true;
    public bool Kinsoku { get; set; } = true; // Asian typography rules
    public bool OverflowPunct { get; set; } // Allow punctuation to extend past margin
    public bool TopLinePunct { get; set; } // Allow punctuation to start line
}

public enum ParagraphAlignment
{
    Left = 0,
    Center = 1,
    Right = 2,
    Justify = 3,
    Distributed = 4,
    ThaiJustify = 5
}

/// <summary>
/// Lightweight shape model used for OfficeArt/Escher-based drawing objects.
/// </summary>
public class ShapeModel
{
    public int Id { get; set; }
    public ShapeType Type { get; set; } = ShapeType.Unknown;
    public ShapeAnchor? Anchor { get; set; }
    public int? ImageIndex { get; set; }
    public string? Text { get; set; }
    /// <summary>
    /// Hint for where this shape should be emitted relative to paragraphs.
    /// -1 means "no preference / fall back to document-level placement".
    /// </summary>
    public int ParagraphIndexHint { get; set; } = -1;

    // Basic style info for non-picture shapes (OfficeArt vectors).
    public int FillColor { get; set; }   // ICO or COLORREF, 0 = auto/none
    public int LineColor { get; set; }   // ICO or COLORREF, 0 = auto/none
    public int LineWidth { get; set; }   // In twips or similar logical units

    /// <summary>
    /// Vertices for non-rectangular text wrapping (wp:wrapPolygon).
    /// </summary>
    public List<System.Drawing.Point>? WrapPolygonVertices { get; set; }
}

public enum ShapeType
{
    Unknown,
    Picture,
    Rectangle,
    Ellipse,
    Textbox
}

/// <summary>
/// Anchor information for floating/inline shapes, expressed in twips.
/// </summary>
public class ShapeAnchor
{
    public bool IsFloating { get; set; }
    public int PageIndex { get; set; }
    public int ParagraphIndex { get; set; }
    public int X { get; set; }
    public int Y { get; set; }
    public int Width { get; set; }
    public int Height { get; set; }
    public ShapeRelativeTo HorizontalRelativeTo { get; set; } = ShapeRelativeTo.Page;
    public ShapeRelativeTo VerticalRelativeTo { get; set; } = ShapeRelativeTo.Page;
    public ShapeWrapType WrapType { get; set; } = ShapeWrapType.Square;
    public int ZOrder { get; set; }
}

/// <summary>
/// Reference frame used for floating shape positioning.
/// </summary>
public enum ShapeRelativeTo
{
    Page,
    Margin,
    Column,
    Paragraph
}

/// <summary>
/// Text wrapping mode for floating shapes.
/// </summary>
public enum ShapeWrapType
{
    None,
    Square
}

/// <summary>
/// Run properties (CHP)
/// </summary>
public class RunProperties
{
    public int FontIndex { get; set; } = -1;
    public string? FontName { get; set; }
    public int FontSize { get; set; } = 24; // In half-points (24 = 12pt)
    public int FontSizeCs { get; set; } = 24; // Complex script
    public bool IsBold { get; set; }
    public bool IsBoldCs { get; set; }
    public bool IsItalic { get; set; }
    public bool IsItalicCs { get; set; }
    public bool IsUnderline { get; set; }
    public UnderlineType UnderlineType { get; set; } = UnderlineType.None;
    public bool IsStrikeThrough { get; set; }
    public bool IsDoubleStrikeThrough { get; set; }
    public bool IsSmallCaps { get; set; }
    public bool IsAllCaps { get; set; }
    public bool IsHidden { get; set; }
    public bool IsSuperscript { get; set; }
    public bool IsSubscript { get; set; }
    public int Color { get; set; } = 0; // ICO color index or direct RGB
    public int BgColor { get; set; } = -1; // None
    public int CharacterSpacingAdjustment { get; set; }
    public int Language { get; set; }
    public string? LanguageAsia { get; set; }
    public string? LanguageCs { get; set; }
    // Phase 3 additions
    public byte HighlightColor { get; set; }
    public uint RgbColor { get; set; }
    public bool HasRgbColor { get; set; }
    public bool IsOutline { get; set; }
    public bool IsShadow { get; set; }
    public bool IsEmboss { get; set; }
    public bool IsImprint { get; set; }
    public int Kerning { get; set; }
    public int Position { get; set; }
    public int CharacterScale { get; set; } = 100; // Character scaling in % (sprmCHwcr)

    // Phase 1 Additions (Typography)
    public bool SnapToGrid { get; set; } = true; // Character level snap to grid (sprmCFIco/sprmCFUsePgsuSettings)
    // Basic Ruby (Furigana) storage
    public string? RubyText { get; set; } // Phonics text if this run is ruby
    
    // Track Changes
    public bool IsDeleted { get; set; }
    public bool IsInserted { get; set; }
    public ushort AuthorIndexDel { get; set; }
    public ushort AuthorIndexIns { get; set; }
    public uint DateDel { get; set; }
    public uint DateIns { get; set; }
}

public enum UnderlineType
{
    None = 0,
    Single = 1,
    WordsOnly = 2,
    Double = 3,
    Dotted = 4,
    Thick = 5,
    Dash = 6,
    DotDash = 7,
    DotDotDash = 8,
    Wave = 9,
    ThickWave = 10
}

/// <summary>
/// Border information
/// </summary>
public class BorderInfo
{
    public BorderStyle Style { get; set; } = BorderStyle.None;
    public int Width { get; set; }
    public int Color { get; set; }
    public int Space { get; set; }
}

public enum BorderStyle
{
    None,
    Single,
    Thick,
    Double,
    Dotted,
    Dashed,
    DotDash,
    DotDotDash,
    Triple,
    ThinThickSmallGap,
    ThickThinSmallGap,
    ThinThickThinSmallGap,
    ThinThickMediumGap,
    ThickThinMediumGap,
    ThinThickThinMediumGap,
    ThinThickLargeGap,
    ThickThinLargeGap,
    ThinThickThinLargeGap,
    Wave
}

/// <summary>
/// Shading information (paragraph/table/cell background and pattern).
/// </summary>
public class ShadingInfo
{
    public ShadingPattern Pattern { get; set; } = ShadingPattern.Clear;
    /// <summary>OOXML w:shd val (e.g. "pct20", "horzStripe") when parsed from SHD ipat; null to use Pattern.</summary>
    public string? PatternVal { get; set; }
    public int ForegroundColor { get; set; }
    public int BackgroundColor { get; set; }
}

public enum ShadingPattern
{
    Clear,
    Percent5,
    Percent10,
    Percent20,
    Percent25,
    Percent30,
    Percent40,
    Percent50,
    Percent60,
    Percent70,
    Percent75,
    Percent80,
    Percent90,
    LightHorizontal,
    DarkHorizontal,
    LightVertical,
    DarkVertical,
    LightDiagonalDown,
    LightDiagonalUp,
    DarkDiagonalDown,
    DarkDiagonalUp,
    Outlined,
    Solid,
    Check,
    DarkGrid,
    DarkTrellis,
    LightGray,
    MediumGray,
    DarkGray
}

/// <summary>
/// Footnote model
/// </summary>
public abstract class NoteModelBase
{
    public int Index { get; set; }
    public string? ReferenceMark { get; set; }
    public List<RunModel> Runs { get; set; } = new();
    public List<ParagraphModel> Paragraphs { get; set; } = new();
    public int CharacterPosition { get; set; }
    public int CharacterLength { get; set; }
}

/// <summary>
/// Footnote model
/// </summary>
public class FootnoteModel : NoteModelBase
{
}

/// <summary>
/// Endnote model
/// </summary>
public class EndnoteModel : NoteModelBase
{
}

/// <summary>
/// Annotation (comment) model
/// </summary>
public class AnnotationModel
{
    public string? Id { get; set; }
    public string? Author { get; set; }
    public DateTime Date { get; set; }
    public string? Initials { get; set; }
    public List<RunModel> Runs { get; set; } = new();
    public List<ParagraphModel> Paragraphs { get; set; } = new();
    public int StartCharacterPosition { get; set; }
    public int EndCharacterPosition { get; set; }
}

/// <summary>
/// Textbox model
/// </summary>
public class TextboxModel
{
    public int Index { get; set; }
    public string? Name { get; set; }
    public List<ParagraphModel> Paragraphs { get; set; } = new();
    public List<RunModel> Runs { get; set; } = new();
    public int Left { get; set; }
    public int Top { get; set; }
    public int Width { get; set; }
    public int Height { get; set; }
    public TextboxWrapMode WrapMode { get; set; } = TextboxWrapMode.Inline;
    public TextboxVerticalAlignment VerticalAlignment { get; set; } = TextboxVerticalAlignment.Top;
    public TextboxHorizontalAlignment HorizontalAlignment { get; set; } = TextboxHorizontalAlignment.Left;
}

public enum TextboxWrapMode
{
    Inline,
    Square,
    Tight,
    Through,
    TopBottom,
    Behind,
    InFront
}

public enum TextboxVerticalAlignment
{
    Top,
    Center,
    Bottom,
    Inside,
    Outside
}

public enum TextboxHorizontalAlignment
{
    Left,
    Center,
    Right,
    Inside,
    Outside
}
