using System.Text;
using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Utils;

namespace Nedev.FileConverters.DocToDocx.Readers;

internal sealed class TextboxAnchorFieldInfo
{
    public int FieldStartCharacterPosition { get; set; }
    public int? FieldSeparatorCharacterPosition { get; set; }
    public int? FieldEndCharacterPosition { get; set; }
    public int AnchorParagraphIndex { get; set; } = -1;
}

/// <summary>
/// Main document reader that orchestrates all the sub-readers.
/// 
/// Read pipeline:
///   1. CfbReader — extract streams from OLE2 container
///   2. FibReader — parse FIB from WordDocument stream
///   3. StyleReader — parse styles from Table stream
///   4. TextReader — parse text via CLX/Piece Table
///   5. Parse paragraphs and runs using FKP/SPRM data
///   6. TableReader — identify table structures
///   7. ImageReader — extract embedded images
/// </summary>
public class DocReader : IDisposable
{

    // ensure legacy code pages are available on .NET 10+ so we can decode ANSI
    // text in East Asian documents (GBK, Shift-JIS, etc).  doing this once in
    // a static ctor keeps the library self-contained and avoids crashes when
    // callers use Encoding.GetEncoding in the readers.
    static DocReader()
    {
        Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
    }

    private CfbReader? _cfb;
    private BinaryReader? _wordDocReader;
    private BinaryReader? _tableReader;
    private BinaryReader? _dataReader;
    private readonly string? _password;

    private FibReader? _fibReader;
    private TextReader? _textReader;
    private StyleReader? _styleReader;
    private DocumentPropertiesReader? _dopReader;
    private TableReader? _tableParseReader;
    private ImageReader? _imageReader;
    private FkpParser? _fkpParser;
    private FootnoteReader? _footnoteReader;
    private AnnotationReader? _annotationReader;
    private SectionReader? _sectionReader;
    private TextboxReader? _textboxReader;
    private HeaderFooterReader? _headerFooterReader;
    private ListReader? _listReader;
    private BookmarkReader? _bookmarkReader;
    private FieldReader? _fieldReader;
    private HyperlinkReader? _hyperlinkReader;
    private OfficeArtReader? _officeArtReader;
    private List<FspaInfo> _fspaAnchors = new();

    // Keep streams alive for reader lifetime
    private Stream? _wordDocStream;
    private Stream? _tableStream;
    private Stream? _dataStream;
    private Stream? _footnoteStream;
    private Stream? _endnoteStream;
    private Stream? _anotStream;
    private Stream? _txbxStream;

    public DocumentModel Document { get; private set; } = new();
    public bool IsLoaded { get; private set; }

    public DocReader(string filePath, string? password = null)
    {
        _cfb = new CfbReader(filePath);
        _password = password;
        InitializeStreams();
    }

    

    public DocReader(Stream stream, string? password = null)
    {
        _cfb = new CfbReader(stream, leaveOpen: true);
        _password = password;
        InitializeStreams();
    }

    private void InitializeStreams()
    {
        // Extract WordDocument stream (required)
        if (!_cfb!.HasStream("WordDocument"))
            throw new InvalidDataException("Not a valid Word document: 'WordDocument' stream not found.");

        // First, read the raw WordDocument stream to parse the FIB
        _wordDocStream = _cfb.GetStream("WordDocument");
        _wordDocReader = new BinaryReader(_wordDocStream, Encoding.Default, leaveOpen: true);

        _fibReader = new FibReader(_wordDocReader);
        _fibReader.Read();
        long rawFibPrefixLength = _wordDocStream.Position;

        // Extract Table stream (0Table or 1Table)
        var tableName = _fibReader.TableStreamName;
        if (!_cfb.HasStream(tableName))
        {
            // Try the other one
            tableName = tableName == "1Table" ? "0Table" : "1Table";
            if (!_cfb.HasStream(tableName))
                throw new InvalidDataException($"Table stream not found. Tried '{_fibReader.TableStreamName}' and '{tableName}'.");
        }
        _tableStream = _cfb.GetStream(tableName);

        // Check for RC4 encryption
        bool isRc4Encrypted = _fibReader.FEncrypted && !_fibReader.FObfuscated;
        EncryptionHelper.DecryptionContext? rc4Context = null;

        if (isRc4Encrypted)
        {
            rc4Context = EncryptionHelper.GetRc4BaseHash(_tableStream, _fibReader.LKey, _password);
            if (rc4Context == null)
            {
                throw new UnauthorizedAccessException("Document is RC4 encrypted. A valid password is required.");
            }

            long wordDocumentClearPrefixLength = CalculateWordDocumentRc4ClearPrefixLength(rawFibPrefixLength, _wordDocStream.Length);
            Logger.Warning($"RC4 WordDocument decryption is preserving the first {wordDocumentClearPrefixLength} bytes as an unencrypted prefix based on the raw FIB length ({rawFibPrefixLength} bytes).");

            // Wrap WordDocument stream
            _wordDocReader.Dispose();
            _wordDocStream.Dispose();
            
            _wordDocStream = _cfb.GetStream("WordDocument");
        }
        else if (_fibReader.FEncrypted && _fibReader.FObfuscated)
        {
            _cfb.SetEncryptionKey(_fibReader.LKey);

            _wordDocReader.Dispose();
            _wordDocStream.Dispose();

            _wordDocStream = _cfb.GetDecryptedStream("WordDocument");
            _tableStream.Dispose();
            _tableStream = _cfb.GetDecryptedStream(tableName);
        }

        // Initialize Table Reader (now we have raw or RC4-wrapped or XOR-wrapped table stream)
        Stream finalTableStream = _tableStream;
        if (isRc4Encrypted)
        {
            // For Table stream, the first lKey bytes are the EncryptionHeader, which are UNENCRYPTED.
            // But RC4 offset starts at 0. So block encryption is based on absolute offset.
            finalTableStream = new Rc4Stream(_tableStream, rc4Context!.BaseKey, streamStartOffset: 0, useSha1: rc4Context.UseSha1, leaveOpen: true);
        }
        _tableReader = new BinaryReader(finalTableStream, Encoding.Default, leaveOpen: true);

        // Redo reader for WordDocument (if RC4)
        Stream finalWordDocStream = _wordDocStream;
        if (isRc4Encrypted)
        {
            long wordDocumentClearPrefixLength = CalculateWordDocumentRc4ClearPrefixLength(rawFibPrefixLength, _wordDocStream.Length);
            finalWordDocStream = new Rc4Stream(_wordDocStream, rc4Context!.BaseKey, streamStartOffset: 0, useSha1: rc4Context.UseSha1, leaveOpen: true, clearPrefixLength: wordDocumentClearPrefixLength);
        }
        _wordDocReader = new BinaryReader(finalWordDocStream, Encoding.Default, leaveOpen: true);

        // Read floating shape anchors from PlcfSpaMom (best-effort).
        try
        {
            _fspaAnchors = FspaReader.ReadPlcSpaMom(_tableReader, _fibReader);
        }
        catch (Exception ex)
        {
            Logger.Warning("Failed to read floating shape anchors; continuing without anchor metadata.", ex);
            _fspaAnchors = new List<FspaInfo>();
        }

        // Extract Data stream (optional — contains pictures, OLE objects)
        if (_cfb.HasStream("Data"))
        {
            var rawDataStream = _cfb.GetStream("Data");
            Stream finalDataStream = rawDataStream;
            
            if (isRc4Encrypted)
            {
                finalDataStream = new Rc4Stream(rawDataStream, rc4Context!.BaseKey, streamStartOffset: 0, useSha1: rc4Context.UseSha1, leaveOpen: true);
            }
            else if (_fibReader.FEncrypted && _fibReader.FObfuscated)
            {
                rawDataStream.Dispose();
                finalDataStream = _cfb.GetDecryptedStream("Data");
            }

            _dataStream = finalDataStream;
            _dataReader = new BinaryReader(_dataStream, Encoding.Default, leaveOpen: true);

            // Initialize OfficeArt/Escher reader on the Data stream (best-effort).
            try
            {
                _officeArtReader = new OfficeArtReader(_dataStream);
            }
            catch (Exception ex)
            {
                Logger.Warning("Failed to initialize OfficeArt reader; continuing without OfficeArt mapping.", ex);
                _officeArtReader = null;
            }
        }

        // Extract footnote/endnote streams (optional)
        if (_cfb.HasStream("Footnote"))
        {
            _footnoteStream = OpenOptionalStoryStream("Footnote", isRc4Encrypted, rc4Context);
        }

        if (_cfb.HasStream("Endnote"))
        {
            _endnoteStream = OpenOptionalStoryStream("Endnote", isRc4Encrypted, rc4Context);
        }

        // Extract annotation stream (optional)
        if (_cfb.HasStream("Anot"))
        {
            _anotStream = OpenOptionalStoryStream("Anot", isRc4Encrypted, rc4Context);
        }

        // Extract textbox stream (optional)
        if (_cfb.HasStream("Txbx"))
        {
            _txbxStream = OpenOptionalStoryStream("Txbx", isRc4Encrypted, rc4Context);
        }

        // Initialize sub-readers
        _textReader = new TextReader(_wordDocReader!, _tableReader!, _fibReader!);
        _styleReader = new StyleReader(_tableReader!, _fibReader!);
        _dopReader = new DocumentPropertiesReader(_tableReader!, _fibReader!);
        _fkpParser = new FkpParser(_wordDocReader!, _tableReader!, _fibReader!, _textReader!);
        _tableParseReader = new TableReader(_wordDocReader!, _tableReader!, _fibReader!, _fkpParser);
        _imageReader = new ImageReader(_wordDocReader!, _dataReader, _fibReader!, _cfb);
        _footnoteReader = new FootnoteReader(_fibReader!, _textReader!, _fkpParser);
        _annotationReader = new AnnotationReader(_tableReader!, _fibReader!, _textReader!);
        _textboxReader = new TextboxReader(_tableReader!, _fibReader!, _textReader!, _fkpParser, Document.Styles);
        _headerFooterReader = new HeaderFooterReader(_tableReader!, _wordDocReader!, _fibReader!, _textReader!);
        _listReader = new ListReader(_tableReader!, _fibReader!);
        _bookmarkReader = new BookmarkReader(_tableReader!, _wordDocReader!, _fibReader!);
        _fieldReader = new FieldReader();
        _hyperlinkReader = new HyperlinkReader();
        _sectionReader = new SectionReader(_tableReader!, _wordDocReader!, _fibReader!);
    }

    private Stream OpenOptionalStoryStream(string streamName, bool isRc4Encrypted, EncryptionHelper.DecryptionContext? rc4Context)
    {
        if (!_fibReader!.FEncrypted)
            return _cfb!.GetStream(streamName);

        if (isRc4Encrypted)
        {
            var rawStream = _cfb!.GetStream(streamName);
            return new Rc4Stream(rawStream, rc4Context!.BaseKey, streamStartOffset: 0, useSha1: rc4Context.UseSha1, leaveOpen: false);
        }

        return _cfb!.GetDecryptedStream(streamName);
    }

    private static long CalculateWordDocumentRc4ClearPrefixLength(long rawFibPrefixLength, long streamLength)
    {
        if (rawFibPrefixLength <= 0 || streamLength <= 0)
            return 0;

        long alignedPrefixLength = ((rawFibPrefixLength + 511) / 512) * 512;
        return Math.Min(streamLength, Math.Max(rawFibPrefixLength, alignedPrefixLength));
    }

    /// <summary>
    /// Loads and parses the document.
    /// </summary>
    public void Load()
    {
        // Step 1: Read document properties
        Document.Properties = _dopReader!.Read();

        // Step 1.5: Read style sheet and themes
        _styleReader!.Read();
        Document.Styles = _styleReader.Styles;
        ThemeReader.Read(_cfb!, Document);
        
        // Step 1.6: Read revision authors
        Document.RevisionAuthors = SttbfHelper.ReadSttbf(_tableReader!, _fibReader.FcSttbfRgtlv, _fibReader.LcbSttbfRgtlv);

        // Step 2: Read list definitions
        _listReader!.Styles = Document.Styles;
        _listReader!.Read();
        Document.NumberingDefinitions = _listReader.NumberingDefinitions;
        Document.ListFormats = _listReader.ListFormats;
        Document.ListFormatOverrides = _listReader.ListFormatOverrides;

        // Step 3: Read text content via Piece Table (with per-run Lid for encoding)
        int totalCp = _fibReader!.CcpText + _fibReader.CcpFtn + _fibReader.CcpHdd + _fibReader.CcpAtn + _fibReader.CcpEdn + _fibReader.CcpTxbx + _fibReader.CcpHdrTxbx;

        totalCp = RepairFootnoteStoryLength(totalCp);

        _textReader!.ReadClx();
        _globalChpMap = _fkpParser!.ReadChpProperties();
        _globalPapMap = _fkpParser.ReadPapProperties();
        _textReader.SetTextFromPieces(totalCp, _globalChpMap);

        if (_bookmarkReader != null)
        {
            _bookmarkReader.Read();
            Document.Bookmarks = _bookmarkReader.Bookmarks.ToList();
        }

        // Step 4: Parse paragraphs and runs
        int mainDocLength = Math.Min(_textReader.Text.Length, _fibReader.CcpText);
        Document.Paragraphs = ParseParagraphsRange(_textReader.Text, 0, mainDocLength, _globalChpMap, _globalPapMap, ref _globalImageCounter);

        // Step 4.5: Parse sections
        ParseSections();

        // Step 5: Parse tables
        _tableParseReader!.ParseTables(Document);

        // Step 6: Extract images
        _imageReader!.ExtractImages(Document);

        // Step 6.1: Scan for additional OLE objects in ObjectPool
        ScanObjectPool(Document);

        // Step 6.5: Parse OfficeArt/Escher shapes and map basic anchors
        if (_officeArtReader != null)
        {
            OfficeArtMapper.AttachShapes(Document, _officeArtReader, _fspaAnchors);
            ApplyPictureShapeDisplaySizes(Document);
        }

        // Step 6.6: Best-effort chart detection
        IdentifyCharts();

        // Step 7: Read footnotes
        if (_footnoteReader != null)
        {
            Document.Footnotes = _footnoteReader.ReadFootnotesWithOffset();
        }

        // Step 8: Read endnotes
        if (_footnoteReader != null)
        {
            Document.Endnotes = _footnoteReader.ReadEndnotesWithOffset();
        }

        if (_annotationReader != null)
        {
            Document.Annotations = _annotationReader.ReadAnnotations();
        }

        // Step 9: Read textboxes
        if (_textboxReader != null)
        {
            Document.Textboxes = _textboxReader.ReadTextboxes();
            AttachTextboxAnchorHints(Document, ReadTextboxAnchorFields());
            MergeTextboxShapesIntoTextboxes(Document);
        }

        // Step 10: Read headers/footers
        if (_headerFooterReader != null)
        {
            _headerFooterReader.Read(Document);
            
            // Extract paragraphs for each header/footer
            if (_globalChpMap != null && _globalPapMap != null)
            {
                foreach (var hf in _headerFooterReader.Headers.Concat(_headerFooterReader.Footers))
                {
                    int headerStoryStartCp = _fibReader.CcpText + _fibReader.CcpFtn;
                    int headerStoryEndCp = headerStoryStartCp + _fibReader.CcpHdd;
                    int absoluteStartCp = Math.Clamp(headerStoryStartCp + hf.CharacterPosition, headerStoryStartCp, headerStoryEndCp);
                    int absoluteEndCp = Math.Clamp(absoluteStartCp + hf.CharacterLength, absoluteStartCp, headerStoryEndCp);
                    var paragraphs = ParseParagraphsRange(_textReader.Text, absoluteStartCp, absoluteEndCp, _globalChpMap, _globalPapMap, ref _globalImageCounter);
                    hf.Paragraphs = HeaderFooterParagraphsLookReasonable(hf.Text, paragraphs)
                        ? paragraphs
                        : new List<ParagraphModel>();
                }
            }

            _headerFooterReader.Headers.RemoveAll(h => !HeaderFooterContentHelper.HasUsableContent(h));
            _headerFooterReader.Footers.RemoveAll(f => !HeaderFooterContentHelper.HasUsableContent(f));

            Document.HeadersFooters.Headers = _headerFooterReader.Headers;
            Document.HeadersFooters.Footers = _headerFooterReader.Footers;
        }

        // 11. Extract VBA Macros if present
        ExtractVbaProject();

        IsLoaded = true;
    }

    internal static void MergeTextboxShapesIntoTextboxes(DocumentModel document)
    {
        if (document.Textboxes.Count == 0 || document.Shapes.Count == 0)
            return;

        var textboxShapes = document.Shapes
            .Where(shape => shape.Type == ShapeType.Textbox)
            .OrderBy(shape => shape.Anchor?.ParagraphIndex ?? shape.ParagraphIndexHint)
            .ThenBy(shape => shape.Id)
            .ToList();

        if (textboxShapes.Count == 0)
            return;

        var matchedShapeIds = new HashSet<int>();
        foreach (var textbox in document.Textboxes.OrderBy(t => t.AnchorParagraphIndex).ThenBy(t => t.AnchorCharacterPosition).ThenBy(t => t.Index))
        {
            var shape = FindBestTextboxShapeMatch(textbox, textboxShapes, matchedShapeIds);
            if (shape == null)
                continue;

            matchedShapeIds.Add(shape.Id);
            var anchor = shape.Anchor;
            if (anchor != null)
            {
                textbox.Left = anchor.X;
                textbox.Top = anchor.Y;
                if (anchor.Width > 0)
                    textbox.Width = anchor.Width;
                if (anchor.Height > 0)
                    textbox.Height = anchor.Height;
                textbox.WrapMode = MapTextboxWrapMode(anchor.WrapType);
            }

            var firstAlignedParagraph = textbox.Paragraphs.FirstOrDefault(paragraph => paragraph.Properties != null);
            if (firstAlignedParagraph?.Properties != null)
            {
                textbox.HorizontalAlignment = firstAlignedParagraph.Properties.Alignment switch
                {
                    ParagraphAlignment.Center => TextboxHorizontalAlignment.Center,
                    ParagraphAlignment.Right => TextboxHorizontalAlignment.Right,
                    _ => TextboxHorizontalAlignment.Left
                };
            }
        }

        document.Shapes.RemoveAll(shape => matchedShapeIds.Contains(shape.Id));
    }

    internal static void AttachTextboxAnchorHints(DocumentModel document, IReadOnlyList<TextboxAnchorFieldInfo> anchorFields)
    {
        if (document.Textboxes.Count == 0 || anchorFields.Count == 0 || document.Paragraphs.Count == 0)
            return;

        var paragraphsByCp = document.Paragraphs
            .Select(paragraph => new
            {
                Paragraph = paragraph,
                StartCp = paragraph.Runs.Count > 0 ? paragraph.Runs.Min(run => run.CharacterPosition) : int.MaxValue
            })
            .Where(item => item.StartCp != int.MaxValue)
            .OrderBy(item => item.StartCp)
            .ToList();

        if (paragraphsByCp.Count == 0)
            return;

        var orderedTextboxes = document.Textboxes
            .OrderBy(textbox => textbox.StoryStartCharacterPosition)
            .ThenBy(textbox => textbox.Index)
            .ToList();

        var orderedAnchorFields = anchorFields
            .OrderBy(field => field.FieldStartCharacterPosition)
            .ToList();

        int count = Math.Min(orderedTextboxes.Count, orderedAnchorFields.Count);
        for (int i = 0; i < count; i++)
        {
            var anchorField = orderedAnchorFields[i];
            int anchorCp = anchorField.FieldStartCharacterPosition;
            var textbox = orderedTextboxes[i];
            textbox.AnchorCharacterPosition = anchorCp;

            var bestParagraph = paragraphsByCp.LastOrDefault(item => item.StartCp <= anchorCp) ?? paragraphsByCp[0];
            textbox.AnchorParagraphIndex = bestParagraph.Paragraph.Index;
            anchorField.AnchorParagraphIndex = bestParagraph.Paragraph.Index;
        }
    }

    private List<TextboxAnchorFieldInfo> ReadTextboxAnchorFields()
    {
        var plcCharacterPositions = new List<int>();
        if (_fibReader == null || _tableReader == null || _textReader == null)
            return new List<TextboxAnchorFieldInfo>();

        if (_fibReader.FcPlcfFldTxbx == 0 || _fibReader.LcbPlcfFldTxbx < 10)
            return new List<TextboxAnchorFieldInfo>();

        try
        {
            _tableReader.BaseStream.Seek(_fibReader.FcPlcfFldTxbx, SeekOrigin.Begin);
            int n = (int)((_fibReader.LcbPlcfFldTxbx - 4) / 6);
            if (n <= 0)
                return new List<TextboxAnchorFieldInfo>();

            var cpArray = new int[n + 1];
            for (int i = 0; i <= n; i++)
            {
                cpArray[i] = _tableReader.ReadInt32();
            }

            for (int i = 0; i < n; i++)
            {
                if (_tableReader.BaseStream.Position + 2 > _tableReader.BaseStream.Length)
                    break;

                _tableReader.ReadUInt16();
                int cp = cpArray[i];
                if (cp < 0 || cp >= _fibReader.CcpText || cp >= _textReader.Text.Length)
                    continue;

                plcCharacterPositions.Add(cp);
            }
        }
        catch (Exception ex)
        {
            Logger.Warning("Failed to read textbox anchor field PLC; continuing with sequence-based textbox matching.", ex);
            return new List<TextboxAnchorFieldInfo>();
        }

        var anchorFields = BuildTextboxAnchorFields(_textReader.Text, _fibReader.CcpText, plcCharacterPositions);
        if (plcCharacterPositions.Count > 0 && anchorFields.Count == 0)
        {
            Logger.Warning("Textbox anchor field PLC was present but no complete textbox field boundaries could be reconstructed; continuing with sequence-based textbox matching.");
        }

        return anchorFields;
    }

    internal static List<TextboxAnchorFieldInfo> BuildTextboxAnchorFields(string text, int mainDocumentLength, IReadOnlyList<int> plcCharacterPositions)
    {
        var anchorFields = new List<TextboxAnchorFieldInfo>();
        var openFields = new Stack<TextboxAnchorFieldInfo>();

        foreach (int cp in plcCharacterPositions.Distinct().OrderBy(cp => cp))
        {
            if (cp < 0 || cp >= mainDocumentLength || cp >= text.Length)
                continue;

            switch (text[cp])
            {
                case FieldReader.FieldStartChar:
                    openFields.Push(new TextboxAnchorFieldInfo { FieldStartCharacterPosition = cp });
                    break;

                case FieldReader.FieldSeparatorChar:
                    if (openFields.Count > 0 && !openFields.Peek().FieldSeparatorCharacterPosition.HasValue)
                    {
                        openFields.Peek().FieldSeparatorCharacterPosition = cp;
                    }
                    break;

                case FieldReader.FieldEndChar:
                    if (openFields.Count > 0)
                    {
                        var field = openFields.Pop();
                        field.FieldEndCharacterPosition = cp;
                        anchorFields.Add(field);
                    }
                    break;
            }
        }

        anchorFields.Sort((left, right) => left.FieldStartCharacterPosition.CompareTo(right.FieldStartCharacterPosition));
        return anchorFields;
    }

    private static ShapeModel? FindBestTextboxShapeMatch(TextboxModel textbox, IReadOnlyList<ShapeModel> textboxShapes, ISet<int> matchedShapeIds)
    {
        ShapeModel? bestShape = null;
        int bestScore = int.MaxValue;

        foreach (var shape in textboxShapes)
        {
            if (matchedShapeIds.Contains(shape.Id))
                continue;

            int shapeParagraph = shape.Anchor?.ParagraphIndex ?? shape.ParagraphIndexHint;
            int score = 0;

            if (textbox.AnchorParagraphIndex >= 0 && shapeParagraph >= 0)
            {
                score += Math.Abs(textbox.AnchorParagraphIndex - shapeParagraph) * 100;
            }
            else if (textbox.AnchorParagraphIndex >= 0 || shapeParagraph >= 0)
            {
                score += 10_000;
            }

            score += Math.Abs(textbox.Index - shape.Id % 1000);

            if (score < bestScore)
            {
                bestScore = score;
                bestShape = shape;
            }
        }

        return bestShape;
    }

    private static TextboxWrapMode MapTextboxWrapMode(ShapeWrapType wrapType)
    {
        return wrapType switch
        {
            ShapeWrapType.Square => TextboxWrapMode.Square,
            ShapeWrapType.Tight => TextboxWrapMode.Tight,
            ShapeWrapType.Through => TextboxWrapMode.Through,
            ShapeWrapType.TopBottom => TextboxWrapMode.TopBottom,
            ShapeWrapType.BehindText => TextboxWrapMode.Behind,
            ShapeWrapType.InFrontOfText => TextboxWrapMode.InFront,
            _ => TextboxWrapMode.Inline
        };
    }

    private int RepairFootnoteStoryLength(int totalCp)
    {
        if (_fibReader == null || _tableReader == null)
            return totalCp;

        if (_fibReader.CcpFtn != 0 || _fibReader.FcFtn == 0 || _fibReader.LcbFtn == 0)
            return totalCp;

        try
        {
            _tableReader.BaseStream.Seek(_fibReader.FcFtn, SeekOrigin.Begin);
            int n = (int)((_fibReader.LcbFtn - 4) / 6);
            if (n <= 0)
                return totalCp;

            int footCpLen = 0;
            for (int i = 0; i <= n; i++)
            {
                footCpLen = _tableReader.ReadInt32();
            }

            _fibReader.SetDerivedFootnoteCharacterCount(footCpLen);

            int candidate = _fibReader.CcpText + _fibReader.CcpFtn + _fibReader.CcpHdd + _fibReader.CcpAtn + _fibReader.CcpEdn + _fibReader.CcpTxbx + _fibReader.CcpHdrTxbx;
            return Math.Max(totalCp, candidate);
        }
        catch
        {
            return totalCp;
        }
    }

    private static bool HeaderFooterParagraphsLookReasonable(string extractedText, List<ParagraphModel> paragraphs)
    {
        if (paragraphs == null || paragraphs.Count == 0)
            return false;

        var paragraphText = NormalizeHeaderFooterComparisonText(string.Concat(paragraphs.Select(p => p.Text)));
        if (paragraphText.Length == 0)
            return false;

        var extracted = NormalizeHeaderFooterComparisonText(extractedText);
        if (extracted.Length == 0)
            return paragraphText.Length <= 64;

        if (paragraphText.Length > Math.Max(96, extracted.Length * 4))
            return false;

        return paragraphText.Contains(extracted, StringComparison.Ordinal) ||
               extracted.Contains(paragraphText, StringComparison.Ordinal);
    }

    private static string NormalizeHeaderFooterComparisonText(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return string.Empty;

        var sb = new StringBuilder(text.Length);
        foreach (var ch in text)
        {
            if (!char.IsWhiteSpace(ch))
                sb.Append(ch);
        }

        return sb.ToString();
    }

    /// <summary>
    /// Creates placeholder ChartModel instances for any OLE streams whose names
    /// look like chart containers. This gives callers a basic, editable chart
    /// in the resulting DOCX even when we do not yet understand the underlying
    /// binary chart format.
    /// </summary>
    /// <summary>Iterates through ObjectPool storage to extract embedded OLE objects.</summary>
    private void ScanObjectPool(DocumentModel document)
    {
        if (_cfb == null || !_cfb.HasStorage("ObjectPool")) return;

        try
        {
            var opStorage = _cfb.GetStorage("ObjectPool");
            if (opStorage == null) return;

            var children = _cfb.GetChildren(opStorage);
            foreach (var child in children)
            {
                if (child.ObjectType != 1) continue; // OBJ_STORAGE

                var objectId = child.Name;
                if (!document.OleObjects.Any(o => o.ObjectId == objectId))
                {
                    try
                    {
                        var oleData = Writers.CfbBuilder.RepackStorage(_cfb, child);
                        var progId = ExtractProgIdFromOleStorage(child);

                        var oleObj = new OleObjectModel
                        {
                            ObjectId = objectId,
                            ProgId = progId ?? "Unknown",
                            ObjectData = oleData
                        };

                        // Convert Equation.3 to OMML
                        if (oleObj.ProgId == "Equation.3")
                        {
                            var eqNative = _cfb.GetChildren(child).FirstOrDefault(c => c.Name == "Equation Native");
                            if (eqNative != null)
                            {
                                var mtefBytes = _cfb.GetStreamBytes(eqNative);
                                var mtefReader = new MtefReader(mtefBytes);
                                oleObj.MathContent = mtefReader.ConvertToOmml();
                            }
                        }

                        document.OleObjects.Add(oleObj);
                    }
                    catch (Exception ex)
                    {
                        Logger.Warning($"Failed to repack OLE storage '{objectId}' from ObjectPool.", ex);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Warning("Failed to scan ObjectPool for embedded OLE objects.", ex);
        }
    }

    private string? ExtractProgIdFromOleStorage(DirectoryEntry storage)
    {
        if (_cfb == null) return null;
        var children = _cfb.GetChildren(storage);
        
        // Try \x03ObjInfo first (standard OLE storage)
        var objInfoEntry = children.FirstOrDefault(c => c.Name == "\x03ObjInfo");
        if (objInfoEntry != null)
        {
            _ = _cfb.GetStreamBytes(objInfoEntry);
        }

        // Check for common stream names that hint at ProgID
        if (children.Any(c => c.Name == "WordDocument")) return "Word.Document.8";
        if (children.Any(c => c.Name == "Workbook")) return "Excel.Sheet.8";
        if (children.Any(c => c.Name == "PowerPoint Document")) return "PowerPoint.Show.8";
        if (children.Any(c => c.Name == "Equation Native" || c.Name == "\x02OlePres000")) return "Equation.3";
        if (children.Any(c => c.Name == "Package")) return "Package";

        return null;
    }

    private void IdentifyCharts()
    {
        if (_cfb == null) return;

        // Look for streams which strongly suggest embedded charts. In many
        // real-world documents these names come from OLE chart objects.
        var chartLikeStreams = _cfb.StreamNames
            .Where(n => n.IndexOf("Chart", StringComparison.OrdinalIgnoreCase) >= 0)
            .ToList();

        if (chartLikeStreams.Count == 0)
            return;

        int existingCount = Document.Charts.Count;
        for (int i = 0; i < chartLikeStreams.Count; i++)
        {
            var name = chartLikeStreams[i];

            // Try to capture the raw OLE stream bytes for this chart so that
            // future phases (or callers) can recover real series data from the
            // original container (e.g. MS Graph or embedded Excel workbook).
            byte[]? sourceBytes = null;
            try
            {
                sourceBytes = _cfb.GetStreamBytes(name);
            }
            catch
            {
                // best-effort only; leave SourceBytes as null on failure
            }

            var model = new ChartModel
            {
                Index = existingCount + i,
                Title = name,
                Type = ChartType.Column,
                Categories = new List<string> { "Category 1", "Category 2", "Category 3" },
                Series =
                {
                    new ChartSeries
                    {
                        Name = "Series 1",
                        Values = new List<double> { 1, 2, 3 }
                    }
                },
                SourceStreamName = name,
                SourceBytes = sourceBytes,
                ParagraphIndexHint = -1
            };

            // Best-effort attempt to recover real categories/series from the
            // embedded bytes; falls back to the placeholder data above when
            // nothing sensible can be inferred.
            BiffChartScanner.TryPopulateChart(model);

            Document.Charts.Add(model);
        }

        // If we still do not have paragraph hints for charts, distribute them
        // roughly evenly across "normal" paragraphs so that each chart appears
        // near some body text rather than all being appended at the end.
        if (Document.Paragraphs.Count > 0)
        {
            var candidateParagraphIndices = Document.Paragraphs
                .Where(p => p.Type == ParagraphType.Normal)
                .Select(p => p.Index)
                .ToList();

            if (candidateParagraphIndices.Count == 0)
            {
                candidateParagraphIndices = Document.Paragraphs.Select(p => p.Index).ToList();
            }

            if (candidateParagraphIndices.Count > 0)
            {
                var chartsNeedingPlacement = Document.Charts
                    .Where(c => c.ParagraphIndexHint < 0)
                    .ToList();

                for (int i = 0; i < chartsNeedingPlacement.Count; i++)
                {
                    var target = candidateParagraphIndices[
                        (int)((long)i * candidateParagraphIndices.Count / chartsNeedingPlacement.Count)
                    ];
                    chartsNeedingPlacement[i].ParagraphIndexHint = target;
                }
            }
        }
    }

    private Dictionary<int, ChpBase>? _globalChpMap;
    private Dictionary<int, PapBase>? _globalPapMap;
    private int _globalImageCounter = 0;

    /// <summary>
    /// Parses a range of text into paragraphs and runs based on FKP structures.
    /// </summary>
    public List<ParagraphModel> ParseParagraphsRange(string text, int startCp, int endCp, Dictionary<int, ChpBase> chpMap, Dictionary<int, PapBase> papMap, ref int imageCounter)
    {
        var paragraphs = new List<ParagraphModel>();
        if (string.IsNullOrEmpty(text)) return paragraphs;

        var paragraphIndex = 0;
        var paraStart = startCp;
        
        endCp = Math.Min(text.Length, endCp);

        for (int i = startCp; i <= endCp; i++)
        {
            bool isParagraphEnd = (i == endCp) || (text[i] == '\r') || (text[i] == '\x0D');

            if (!isParagraphEnd) continue;

            var paraText = text.Substring(paraStart, i - paraStart);
            var paraStartCp = paraStart;
            paraStart = i + 1; // skip the delimiter
            if (paraStart > endCp) paraStart = endCp;

            // Get PAP for this paragraph (use the first CP of the paragraph; if none, use nearest preceding CP so style/alignment are preserved)
            PapBase? pap = null;
            if (paraStartCp < text.Length && papMap.TryGetValue(paraStartCp, out var foundPap))
                pap = foundPap;
            if (pap == null && papMap.Count > 0)
            {
                // When the paragraph start CP falls in a PLC gap, prefer PAP entries that
                // still belong to the current paragraph before borrowing metadata from the
                // preceding paragraph. This preserves paragraph-local direct formatting such
                // as character-unit first-line indents without relying on text-specific fixes.
                int paragraphEndCp = paraStartCp + paraText.Length;
                for (int cp = paraStartCp + 1; cp <= paragraphEndCp; cp++)
                {
                    if (papMap.TryGetValue(cp, out var nextPap))
                    {
                        pap = nextPap;
                        break;
                    }
                }

                if (pap == null)
                {
                    for (int cp = paraStartCp - 1; cp >= Math.Max(0, startCp - 2000); cp--)
                    {
                        if (papMap.TryGetValue(cp, out var prevPap)) { pap = prevPap; break; }
                    }
                }

                if (pap == null)
                {
                    var firstKey = papMap.Keys.Min();
                    if (firstKey <= paraStartCp + 2000)
                        papMap.TryGetValue(firstKey, out pap);
                }
            }

            var paragraph = new ParagraphModel
            {
                Index = paragraphIndex++,
                StartCp = paraStartCp,
                EndCp = paraText.Length > 0 ? paraStartCp + paraText.Length - 1 : paraStartCp,
                RawText = paraText,
                Type = ParagraphType.Normal,
                Properties = pap != null 
                    ? _fkpParser!.ConvertToParagraphProperties(pap, Document.Styles)
                    : new ParagraphProperties { StyleIndex = 0 },
                // MS-DOC itap is zero-based for in-table paragraphs: 0 = top-level
                // table, 1 = nested one level deeper, etc. Convert it to the
                // one-based levels used by TableReader so nested tables are not
                // collapsed onto the top-level table.
                NestingLevel = pap?.InTable == true ? Math.Max(1, pap.Itap + 1) : 0,
                ListFormatId = pap?.ListFormatId ?? 0,
                ListLevel = pap?.ListLevel ?? 0
            };

            OverlayParagraphPropertiesFromRange(paragraph.Properties, papMap, paraStartCp, paraStartCp + paraText.Length);

            var piecePap = _textReader?.GetPieceParagraphPropertiesAtCp(paraStartCp);
            if (piecePap != null)
            {
                OverlayParagraphProperties(paragraph.Properties, piecePap);
            }

            // Detect special paragraph types
            if (paraText.Contains('\x07') || pap?.InTable == true || (pap?.Itap ?? 0) > 0)
            {
                // Contains cell end mark — this is a table cell paragraph
                paragraph.Type = ParagraphType.TableCell;
            }
            else if (paraText.Contains('\x0C'))
            {
                paragraph.Type = ParagraphType.PageBreak;
            }

            // Split paragraph into runs based on CHP changes; when no direct CHP,
            // inherit run properties (font size, color, etc.) from the paragraph style.
            var runs = ParseRunsInParagraph(paraText, paraStartCp, chpMap, papMap, ref imageCounter);
            paragraph.Runs.AddRange(runs);

            // If no runs were created (no CHP data), create one run using paragraph style so font/color/size are preserved
            if (paragraph.Runs.Count == 0)
            {
                var cleanText = CleanSpecialChars(paraText);
                if (!string.IsNullOrEmpty(cleanText))
                {
                    var runProps = GetRunPropertiesFromParagraphStyleAtCp(papMap, paraStartCp) ?? new RunProperties { FontSize = 24 };
                    paragraph.Runs.Add(new RunModel
                    {
                        Text = cleanText,
                        CharacterPosition = paraStartCp,
                        CharacterLength = paraText.Length,
                        Properties = runProps
                    });
                }
            }

            paragraph.Runs = ApplyBookmarkMarkers(paragraph.Runs, paraStartCp, paraStartCp + paraText.Length);
            ApplyParagraphStyleDefaults(paragraph);

            paragraphs.Add(paragraph);
        }

        ApplyNarrativeFirstLineIndentHeuristics(paragraphs);
        
        return paragraphs;
    }

    private void ApplyNarrativeFirstLineIndentHeuristics(List<ParagraphModel> paragraphs)
    {
        if (paragraphs.Count < 2)
            return;

        int inferredIndentChars = InferNarrativeFirstLineChars(paragraphs);
        if (inferredIndentChars <= 0)
            return;

        for (int index = 1; index < paragraphs.Count; index++)
        {
            var current = paragraphs[index];
            var previous = paragraphs[index - 1];

            if (!ShouldApplyNarrativeFirstLineIndent(previous, current))
                continue;

            current.Properties ??= new ParagraphProperties();
            current.Properties.IndentFirstLineChars = inferredIndentChars;
        }
    }

    private int InferNarrativeFirstLineChars(List<ParagraphModel> paragraphs)
    {
        var explicitIndents = paragraphs
            .Where(paragraph => paragraph.Properties?.IndentFirstLineChars > 0)
            .Select(paragraph => paragraph.Properties!.IndentFirstLineChars)
            .GroupBy(value => value)
            .OrderByDescending(group => group.Count())
            .ThenByDescending(group => group.Key)
            .Select(group => group.Key)
            .ToList();

        if (explicitIndents.Count > 0)
            return explicitIndents[0];

        return 200;
    }

    private bool ShouldApplyNarrativeFirstLineIndent(ParagraphModel previous, ParagraphModel current)
    {
        if (!IsHeadingOrTitleParagraph(previous))
            return false;

        if (current.Type != ParagraphType.Normal || current.Properties == null)
            return false;

        if (current.Properties.ListFormatId > 0 || current.ListFormatId > 0)
            return false;

        if (current.Properties.IndentLeft != 0 || current.Properties.IndentLeftChars != 0 ||
            current.Properties.IndentRight != 0 || current.Properties.IndentRightChars != 0 ||
            current.Properties.IndentFirstLine != 0 || current.Properties.IndentFirstLineChars != 0)
            return false;

        var text = current.Text?.Trim() ?? string.Empty;
        if (text.Length < 60)
            return false;

        int wordCount = text.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries).Length;
        return wordCount >= 8;
    }

    private bool IsHeadingOrTitleParagraph(ParagraphModel paragraph)
    {
        var styleIndex = paragraph.Properties?.StyleIndex ?? 0;
        if (styleIndex == StyleIds.TITLE)
            return true;

        if (styleIndex >= StyleIds.HEADING_1 && styleIndex <= StyleIds.HEADING_9)
            return true;

        var style = Document.Styles?.Styles?.FirstOrDefault(candidate =>
            candidate.Type == StyleType.Paragraph && candidate.StyleId == styleIndex);
        if (style?.Name == null)
            return false;

        return style.Name.StartsWith("Heading", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(style.Name, "Title", StringComparison.OrdinalIgnoreCase);
    }

    private List<RunModel> ApplyBookmarkMarkers(List<RunModel> runs, int paragraphStartCp, int paragraphEndCp)
    {
        if (Document.Bookmarks.Count == 0)
            return runs;

        if (runs.Count == 0)
        {
            var markerOnlyRuns = new List<RunModel>();

            foreach (var bookmark in Document.Bookmarks.Where(bookmark => bookmark.StartCp == paragraphStartCp))
            {
                markerOnlyRuns.Add(CreateBookmarkRun(bookmark, isStart: true));
            }

            foreach (var bookmark in Document.Bookmarks.Where(bookmark => bookmark.EndCp == paragraphStartCp))
            {
                markerOnlyRuns.Add(CreateBookmarkRun(bookmark, isStart: false));
            }

            if (paragraphEndCp != paragraphStartCp)
            {
                foreach (var bookmark in Document.Bookmarks.Where(bookmark => bookmark.StartCp == paragraphEndCp))
                {
                    markerOnlyRuns.Add(CreateBookmarkRun(bookmark, isStart: true));
                }

                foreach (var bookmark in Document.Bookmarks.Where(bookmark => bookmark.EndCp == paragraphEndCp))
                {
                    markerOnlyRuns.Add(CreateBookmarkRun(bookmark, isStart: false));
                }
            }

            return markerOnlyRuns;
        }

        var markerPositions = Document.Bookmarks
            .SelectMany(bookmark => new[] { bookmark.StartCp, bookmark.EndCp })
            .Where(cp => cp > paragraphStartCp && cp < paragraphEndCp)
            .Distinct()
            .OrderBy(cp => cp)
            .ToList();

        var splitRuns = SplitRunsAtBookmarkBoundaries(runs, markerPositions);
        var withMarkers = new List<RunModel>(splitRuns.Count + Document.Bookmarks.Count * 2);

        foreach (var run in splitRuns)
        {
            foreach (var bookmark in Document.Bookmarks.Where(bookmark => bookmark.StartCp == run.CharacterPosition))
            {
                withMarkers.Add(CreateBookmarkRun(bookmark, isStart: true));
            }

            foreach (var bookmark in Document.Bookmarks.Where(bookmark => bookmark.EndCp == run.CharacterPosition))
            {
                withMarkers.Add(CreateBookmarkRun(bookmark, isStart: false));
            }

            withMarkers.Add(run);
        }

        foreach (var bookmark in Document.Bookmarks.Where(bookmark => bookmark.StartCp == paragraphEndCp))
        {
            withMarkers.Add(CreateBookmarkRun(bookmark, isStart: true));
        }

        foreach (var bookmark in Document.Bookmarks.Where(bookmark => bookmark.EndCp == paragraphEndCp))
        {
            withMarkers.Add(CreateBookmarkRun(bookmark, isStart: false));
        }

        return withMarkers;
    }

    private static List<RunModel> SplitRunsAtBookmarkBoundaries(List<RunModel> runs, List<int> boundaries)
    {
        if (boundaries.Count == 0)
            return runs;

        var splitRuns = new List<RunModel>(runs.Count);

        foreach (var run in runs)
        {
            if (!CanSplitRunAtBookmarkBoundary(run))
            {
                splitRuns.Add(run);
                continue;
            }

            int segmentStart = run.CharacterPosition;
            int segmentTextOffset = 0;
            int runEnd = run.CharacterPosition + run.CharacterLength;
            var runBoundaries = boundaries.Where(cp => cp > run.CharacterPosition && cp < runEnd).ToList();

            if (runBoundaries.Count == 0)
            {
                splitRuns.Add(run);
                continue;
            }

            foreach (var boundary in runBoundaries)
            {
                int segmentLength = boundary - segmentStart;
                if (segmentLength > 0)
                {
                    splitRuns.Add(CloneRunSegment(run, run.Text.Substring(segmentTextOffset, segmentLength), segmentStart, segmentLength));
                }

                segmentStart = boundary;
                segmentTextOffset += segmentLength;
            }

            int trailingLength = runEnd - segmentStart;
            if (trailingLength > 0)
            {
                splitRuns.Add(CloneRunSegment(run, run.Text.Substring(segmentTextOffset, trailingLength), segmentStart, trailingLength));
            }
        }

        return splitRuns;
    }

    private static bool CanSplitRunAtBookmarkBoundary(RunModel run)
    {
        return !run.IsPicture &&
               !run.IsField &&
               !run.IsOle &&
               run.CharacterLength > 0 &&
               run.Text.Length == run.CharacterLength;
    }

    private static RunModel CloneRunSegment(RunModel run, string text, int characterPosition, int characterLength)
    {
        return new RunModel
        {
            Text = text,
            Properties = run.Properties,
            IsField = run.IsField,
            FieldCode = run.FieldCode,
            CharacterPosition = characterPosition,
            CharacterLength = characterLength,
            IsPicture = run.IsPicture,
            ImageIndex = run.ImageIndex,
            DisplayWidthTwips = run.DisplayWidthTwips,
            DisplayHeightTwips = run.DisplayHeightTwips,
            FcPic = run.FcPic,
            ImageRelationshipId = run.ImageRelationshipId,
            IsOle = run.IsOle,
            OleObjectId = run.OleObjectId,
            OleProgId = run.OleProgId,
            IsHyperlink = run.IsHyperlink,
            HyperlinkUrl = run.HyperlinkUrl,
            HyperlinkBookmark = run.HyperlinkBookmark,
            HyperlinkRelationshipId = run.HyperlinkRelationshipId,
            CropTop = run.CropTop,
            CropBottom = run.CropBottom,
            CropLeft = run.CropLeft,
            CropRight = run.CropRight,
            FlipHorizontal = run.FlipHorizontal,
            FlipVertical = run.FlipVertical
        };
    }

    private static RunModel CreateBookmarkRun(BookmarkModel bookmark, bool isStart)
    {
        return new RunModel
        {
            Text = string.Empty,
            CharacterPosition = isStart ? bookmark.StartCp : bookmark.EndCp,
            CharacterLength = 0,
            IsBookmark = true,
            IsBookmarkStart = isStart,
            BookmarkName = bookmark.Name
        };
    }

    /// <summary>
    /// Parses runs within a paragraph based on CHP property changes.
    /// When there is no direct CHP at a position, run properties are taken from the paragraph's style so that
    /// font size and color from the .doc are preserved.
    /// </summary>
    private List<RunModel> ParseRunsInParagraph(string paraText, int paraStartCp, Dictionary<int, ChpBase> chpMap, Dictionary<int, PapBase> papMap, ref int imageCounter)
    {
        var runs = new List<RunModel>();
        if (string.IsNullOrEmpty(paraText)) return runs;

        string? activeEmbedProgId = null;
        string? activeOleObjectId = null;
        HyperlinkModel? activeHyperlink = null;
        var activeFieldCode = new StringBuilder();
        bool collectingFieldCode = false;
        bool insideFieldResult = false;

        var runStart = 0;
        ChpBase? currentChp = null;

        for (int i = 0; i <= paraText.Length; i++)
        {
            var cp = paraStartCp + i;
            ChpBase? chpAtCp = null;
            
            if (chpMap.TryGetValue(cp, out var foundChp))
                chpAtCp = foundChp;

            var pieceChp = _textReader?.GetPieceRunPropertiesAtCp(cp);
            if (pieceChp != null)
            {
                if (chpAtCp == null)
                {
                    chpAtCp = pieceChp;
                }
                else
                {
                    var mergedPieceChp = CloneChpBase(chpAtCp);
                    OverlayChpBase(mergedPieceChp, pieceChp);
                    chpAtCp = mergedPieceChp;
                }
            }

            // Isolate control markers so field instructions/results and object anchors do not smear into one visible run.
            bool isSpecialChar = i < paraText.Length && (paraText[i] == '\x01' || paraText[i] == '\x08' || paraText[i] == '\x13' || paraText[i] == '\x14' || paraText[i] == '\x15');
            bool previousWasSpecial = i > 0 && (paraText[i - 1] == '\x01' || paraText[i - 1] == '\x08' || paraText[i - 1] == '\x13' || paraText[i - 1] == '\x14' || paraText[i - 1] == '\x15');
            
            bool chpChanged = i == paraText.Length || !ChpEquals(currentChp, chpAtCp) || isSpecialChar || previousWasSpecial;

            if (chpChanged && runStart < i)
            {
                var runText = paraText.Substring(runStart, i - runStart);
                var cleanText = CleanSpecialChars(runText);
                var hasPictureMarker = runText.Contains('\x01') || runText.Contains('\x08');

                // Even if text is empty (special char stripped), preserve it if it's visual or structural (hyperlink/field)
                if (!string.IsNullOrEmpty(cleanText) || hasPictureMarker || runText.Contains('\x13') || runText.Contains('\x14') || runText.Contains('\x15'))
                {
                    var runProps = GetEffectiveRunPropertiesAtCp(papMap, paraStartCp + runStart, currentChp);
                    bool wasCollectingFieldCode = collectingFieldCode;

                    var run = new RunModel
                    {
                        Text = cleanText,
                        CharacterPosition = paraStartCp + runStart,
                        CharacterLength = runText.Length,
                        Properties = runProps
                    };

                    run.FcPic = currentChp?.FcPic ?? 0;

                    var completedFieldCode = TryAdvanceFieldState(runText, activeFieldCode, ref collectingFieldCode, ref insideFieldResult);
                    var isPicture = hasPictureMarker && !wasCollectingFieldCode && !collectingFieldCode && !insideFieldResult;
                    if ((wasCollectingFieldCode || runText.Contains(FieldReader.FieldStartChar)) && !insideFieldResult)
                    {
                        run.Text = string.Empty;
                    }

                    if (!string.IsNullOrEmpty(completedFieldCode))
                    {
                        run.IsField = true;
                        run.FieldCode = completedFieldCode;

                        if (_fieldReader != null)
                        {
                            var parsedField = _fieldReader.ParseField(run.FieldCode);
                            if (parsedField != null && parsedField.Type == FieldType.Embed)
                            {
                                activeEmbedProgId = parsedField.Arguments;
                            }
                            else if (parsedField != null && parsedField.Type == FieldType.Hyperlink && _hyperlinkReader != null)
                            {
                                activeHyperlink = _hyperlinkReader.ParseHyperlink(run.FieldCode)
                                                ?? _hyperlinkReader.CreateHyperlink(parsedField.Arguments);
                            }
                        }
                    }

                    // Apply active hyperlink state mapped across all split runs within the boundaries
                    if (insideFieldResult && activeHyperlink != null && (!string.IsNullOrEmpty(cleanText) || isPicture) && (!string.IsNullOrEmpty(activeHyperlink.Url) || !string.IsNullOrEmpty(activeHyperlink.Bookmark)))
                    {
                        run.IsHyperlink = true;
                        run.HyperlinkUrl = activeHyperlink.Url;
                        run.HyperlinkBookmark = activeHyperlink.Bookmark;
                        run.HyperlinkRelationshipId = activeHyperlink.RelationshipId;
                        run.IsField = false; // Treat as hyperlink, overriding generic field semantics

                        if (!Document.Hyperlinks.Any(h => string.Equals(h.Url, activeHyperlink.Url, StringComparison.OrdinalIgnoreCase)
                                                            && string.Equals(h.Bookmark, activeHyperlink.Bookmark, StringComparison.Ordinal)))
                        {
                            Document.Hyperlinks.Add(activeHyperlink);
                        }
                    }

                    // Check if this run contains the field separator \x14 which holds the OLE Object ID in FcPic
                    if (runText.Contains('\x14') && run.FcPic != 0 && activeEmbedProgId != null)
                    {
                        activeOleObjectId = $"_{run.FcPic}";
                        if (_cfb != null)
                        {
                            var storage = _cfb.GetStorage($"ObjectPool/{activeOleObjectId}");
                            if (storage != null && !Document.OleObjects.Any(o => o.ObjectId == activeOleObjectId))
                            {
                                try
                                {
                                    byte[] oleData = Writers.CfbBuilder.RepackStorage(_cfb, storage);
                                    var oleObj = new OleObjectModel
                                    {
                                        ObjectId = activeOleObjectId,
                                        ProgId = activeEmbedProgId,
                                        ObjectData = oleData
                                    };
                                    Document.OleObjects.Add(oleObj);
                                }
                                catch (Exception ex)
                                {
                                    Logger.Warning($"Failed to extract OLE object '{activeOleObjectId}' from field data.", ex);
                                }
                            }
                        }
                    }

                    if (isPicture)
                    {
                        run.IsPicture = true;
                        run.ImageIndex = imageCounter++;
                        
                        if (activeOleObjectId != null && activeEmbedProgId != null)
                        {
                            run.IsOle = true;
                            run.OleObjectId = activeOleObjectId;
                            run.OleProgId = activeEmbedProgId;
                        }
                    }

                    if (runText.Contains('\x15'))
                    {
                        // Field end clears persistent modes
                        activeFieldCode.Clear();
                        collectingFieldCode = false;
                        insideFieldResult = false;
                        activeEmbedProgId = null;
                        activeOleObjectId = null;
                        activeHyperlink = null;
                    }

                    runs.Add(run);
                }

                runStart = i;
            }

            currentChp = chpAtCp;
        }

        return runs;
    }

    private static string? TryAdvanceFieldState(string runText, StringBuilder activeFieldCode, ref bool collectingFieldCode, ref bool insideFieldResult)
    {
        string? completedFieldCode = null;
        int index = 0;

        while (index < runText.Length)
        {
            if (!collectingFieldCode)
            {
                var fieldStartIndex = runText.IndexOf(FieldReader.FieldStartChar, index);
                if (fieldStartIndex < 0)
                    break;

                activeFieldCode.Clear();
                collectingFieldCode = true;
                insideFieldResult = false;
                index = fieldStartIndex + 1;
                continue;
            }

            int separatorIndex = runText.IndexOf(FieldReader.FieldSeparatorChar, index);
            int fieldEndIndex = runText.IndexOf(FieldReader.FieldEndChar, index);

            int stopIndex;
            bool foundSeparator;
            if (separatorIndex >= 0 && (fieldEndIndex < 0 || separatorIndex < fieldEndIndex))
            {
                stopIndex = separatorIndex;
                foundSeparator = true;
            }
            else if (fieldEndIndex >= 0)
            {
                stopIndex = fieldEndIndex;
                foundSeparator = false;
            }
            else
            {
                stopIndex = runText.Length;
                foundSeparator = false;
            }

            if (stopIndex > index)
                activeFieldCode.Append(runText, index, stopIndex - index);

            if (foundSeparator)
            {
                completedFieldCode = activeFieldCode.ToString().Trim();
                activeFieldCode.Clear();
                collectingFieldCode = false;
                insideFieldResult = true;
                index = stopIndex + 1;
                continue;
            }

            if (fieldEndIndex >= 0)
            {
                activeFieldCode.Clear();
                collectingFieldCode = false;
                insideFieldResult = false;
                index = fieldEndIndex + 1;
                continue;
            }

            break;
        }

        return completedFieldCode;
    }

    /// <summary>
    /// Compares two CHP objects for equality so we split runs when formatting (including color/size) changes.
    /// </summary>
    private static bool ChpEquals(ChpBase? a, ChpBase? b)
    {
        if (ReferenceEquals(a, b)) return true;
        if (a == null || b == null) return false;

         return a.StyleId == b.StyleId &&
             a.IsBold == b.IsBold &&
             a.IsBoldCs == b.IsBoldCs &&
             a.IsItalic == b.IsItalic &&
             a.HasExplicitItalic == b.HasExplicitItalic &&
             a.IsItalicCs == b.IsItalicCs &&
             a.HasExplicitItalicCs == b.HasExplicitItalicCs &&
                         a.IsSuperscript == b.IsSuperscript &&
                         a.IsSubscript == b.IsSubscript &&
               a.IsStrikeThrough == b.IsStrikeThrough &&
             a.IsDoubleStrikeThrough == b.IsDoubleStrikeThrough &&
               a.IsUnderline == b.IsUnderline &&
             a.Underline == b.Underline &&
               a.FontSize == b.FontSize &&
               a.FontSizeCs == b.FontSizeCs &&
               a.FontIndex == b.FontIndex &&
             a.FontIndexCs == b.FontIndexCs &&
             a.Color == b.Color &&
             a.HighlightColor == b.HighlightColor &&
               a.HasRgbColor == b.HasRgbColor &&
               a.RgbColor == b.RgbColor &&
             a.IsOutline == b.IsOutline &&
             a.IsShadow == b.IsShadow &&
             a.IsEmboss == b.IsEmboss &&
             a.IsImprint == b.IsImprint &&
                         BordersEqual(a.Border, b.Border) &&
             a.Position == b.Position &&
             a.Kerning == b.Kerning &&
                         a.Scale == b.Scale &&
                         a.EastAsianLayoutType == b.EastAsianLayoutType &&
                         a.IsEastAsianVertical == b.IsEastAsianVertical &&
                         a.IsEastAsianVerticalCompress == b.IsEastAsianVerticalCompress;
    }

        private static bool BordersEqual(BorderInfo? a, BorderInfo? b)
        {
                if (ReferenceEquals(a, b)) return true;
                if (a == null || b == null) return false;

                return a.Style == b.Style &&
                             a.Width == b.Width &&
                             a.Color == b.Color &&
                             a.Space == b.Space;
        }

    private static ChpBase CloneChpBase(ChpBase source)
    {
        return new ChpBase
        {
            FontIndex = source.FontIndex,
            StyleId = source.StyleId,
            FontSize = source.FontSize,
            FontSizeCs = source.FontSizeCs,
            IsBold = source.IsBold,
            IsBoldCs = source.IsBoldCs,
            IsItalic = source.IsItalic,
            IsItalicCs = source.IsItalicCs,
            HasExplicitItalic = source.HasExplicitItalic,
            HasExplicitItalicCs = source.HasExplicitItalicCs,
            IsUnderline = source.IsUnderline,
            Underline = source.Underline,
            IsStrikeThrough = source.IsStrikeThrough,
            IsSmallCaps = source.IsSmallCaps,
            IsAllCaps = source.IsAllCaps,
            IsHidden = source.IsHidden,
            IsSuperscript = source.IsSuperscript,
            IsSubscript = source.IsSubscript,
            Color = source.Color,
            FontIndexCs = source.FontIndexCs,
            CharacterSpacingAdjustment = source.CharacterSpacingAdjustment,
            Language = source.Language,
            LanguageId = source.LanguageId,
            IsDoubleStrikeThrough = source.IsDoubleStrikeThrough,
            DxaOffset = source.DxaOffset,
            IsOutline = source.IsOutline,
            Kerning = source.Kerning,
            Position = source.Position,
            FcPic = source.FcPic,
            Scale = source.Scale,
            HighlightColor = source.HighlightColor,
            IsShadow = source.IsShadow,
            IsEmboss = source.IsEmboss,
            IsImprint = source.IsImprint,
            Border = source.Border == null
                ? null
                : new BorderInfo
                {
                    Style = source.Border.Style,
                    Width = source.Border.Width,
                    Color = source.Border.Color,
                    Space = source.Border.Space
                },
            RgbColor = source.RgbColor,
            HasRgbColor = source.HasRgbColor,
            EastAsianLayoutType = source.EastAsianLayoutType,
            IsEastAsianVertical = source.IsEastAsianVertical,
            IsEastAsianVerticalCompress = source.IsEastAsianVerticalCompress,
            IsDeleted = source.IsDeleted,
            IsInserted = source.IsInserted,
            AuthorIndexDel = source.AuthorIndexDel,
            AuthorIndexIns = source.AuthorIndexIns,
            DateDel = source.DateDel,
            DateIns = source.DateIns
        };
    }

    private static void OverlayChpBase(ChpBase target, ChpBase overlay)
    {
        if (overlay.FontIndex != -1) target.FontIndex = overlay.FontIndex;
        if (overlay.StyleId != 0) target.StyleId = overlay.StyleId;
        if (overlay.FontSize != 24) target.FontSize = overlay.FontSize;
        if (overlay.FontSizeCs != 24) target.FontSizeCs = overlay.FontSizeCs;
        target.IsBold |= overlay.IsBold;
        target.IsBoldCs |= overlay.IsBoldCs;
        if (overlay.HasExplicitItalic)
        {
            target.IsItalic = overlay.IsItalic;
            target.HasExplicitItalic = true;
        }
        else
        {
            target.IsItalic |= overlay.IsItalic;
        }

        if (overlay.HasExplicitItalicCs)
        {
            target.IsItalicCs = overlay.IsItalicCs;
            target.HasExplicitItalicCs = true;
        }
        else
        {
            target.IsItalicCs |= overlay.IsItalicCs;
        }
        target.IsUnderline |= overlay.IsUnderline;
        if (overlay.Underline != 0) target.Underline = overlay.Underline;
        target.IsStrikeThrough |= overlay.IsStrikeThrough;
        target.IsDoubleStrikeThrough |= overlay.IsDoubleStrikeThrough;
        target.IsSmallCaps |= overlay.IsSmallCaps;
        target.IsAllCaps |= overlay.IsAllCaps;
        target.IsHidden |= overlay.IsHidden;
        target.IsSuperscript |= overlay.IsSuperscript;
        target.IsSubscript |= overlay.IsSubscript;
        if (overlay.Color != 0) target.Color = overlay.Color;
        if (overlay.FontIndexCs != -1) target.FontIndexCs = overlay.FontIndexCs;
        if (overlay.CharacterSpacingAdjustment != 0) target.CharacterSpacingAdjustment = overlay.CharacterSpacingAdjustment;
        if (overlay.Language != 0) target.Language = overlay.Language;
        if (overlay.LanguageId != 0) target.LanguageId = overlay.LanguageId;
        if (overlay.DxaOffset != 0) target.DxaOffset = overlay.DxaOffset;
        target.IsOutline |= overlay.IsOutline;
        if (overlay.Kerning != 0) target.Kerning = overlay.Kerning;
        if (overlay.Position != 0) target.Position = overlay.Position;
        if (overlay.FcPic != 0) target.FcPic = overlay.FcPic;
        if (overlay.Scale != 100) target.Scale = overlay.Scale;
        if (overlay.HighlightColor != 0) target.HighlightColor = overlay.HighlightColor;
        target.IsShadow |= overlay.IsShadow;
        target.IsEmboss |= overlay.IsEmboss;
        target.IsImprint |= overlay.IsImprint;
        if (overlay.Border != null)
        {
            target.Border = new BorderInfo
            {
                Style = overlay.Border.Style,
                Width = overlay.Border.Width,
                Color = overlay.Border.Color,
                Space = overlay.Border.Space
            };
        }
        if (overlay.HasRgbColor)
        {
            target.RgbColor = overlay.RgbColor;
            target.HasRgbColor = true;
        }
        if (overlay.EastAsianLayoutType != 0) target.EastAsianLayoutType = overlay.EastAsianLayoutType;
        target.IsEastAsianVertical |= overlay.IsEastAsianVertical;
        target.IsEastAsianVerticalCompress |= overlay.IsEastAsianVerticalCompress;
        target.IsDeleted |= overlay.IsDeleted;
        target.IsInserted |= overlay.IsInserted;
        if (overlay.AuthorIndexDel != 0) target.AuthorIndexDel = overlay.AuthorIndexDel;
        if (overlay.AuthorIndexIns != 0) target.AuthorIndexIns = overlay.AuthorIndexIns;
        if (overlay.DateDel != 0) target.DateDel = overlay.DateDel;
        if (overlay.DateIns != 0) target.DateIns = overlay.DateIns;
    }

    private void ParseSections()
    {
        if (_sectionReader == null) return;
        var sectionInfos = _sectionReader.ReadSections();
        if (sectionInfos.Count == 0) return;

        // Map CP to paragraph index
        foreach (var sec in sectionInfos)
        {
            sec.SectionIndex = sectionInfos.IndexOf(sec);
            sec.StartParagraphIndex = GetParagraphIndexAtCp(sec.StartCp);

            // A gutter margin is only meaningful when mirror/facing page
            // layout is enabled. Word's binary section SPRMs can carry a
            // non-zero gutter operand even when the document layout does not
            // use mirrored margins, in which case the effective gutter should
            // be treated as zero.
            if (!Document.Properties.FMirrorMargins && !Document.Properties.FFacingPages)
            {
                sec.Gutter = 0;
            }
        }
        
        Document.Properties.Sections = sectionInfos;
        
        // If there are sections, the first section's properties often override the global document ones
        if (sectionInfos.Count > 0)
        {
            var s = sectionInfos[0];
            if (Document.Properties.SectionStartPageNumber > 1)
                s.PageNumberStart = Document.Properties.SectionStartPageNumber;
            Document.Properties.PageWidth = s.PageWidth;
            Document.Properties.PageHeight = s.PageHeight;
            Document.Properties.MarginTop = s.MarginTop;
            Document.Properties.MarginBottom = s.MarginBottom;
            Document.Properties.MarginLeft = s.MarginLeft;
            Document.Properties.MarginRight = s.MarginRight;
        }
    }

    private int GetParagraphIndexAtCp(int cp)
    {
        if (Document.Paragraphs.Count == 0) return 0;
        // Simple linear search for now
        for (int i = 0; i < Document.Paragraphs.Count; i++)
        {
            var p = Document.Paragraphs[i];
            if (p.Runs.Count > 0 && p.Runs[0].CharacterPosition >= cp)
            {
                return i;
            }
        }
        return Document.Paragraphs.Count - 1;
    }

    /// <summary>
    /// Gets run properties from the paragraph style at the given CP when there is no direct CHP,
    /// so that style-based font size and color from the .doc are preserved.
    /// </summary>
    private RunProperties GetEffectiveRunPropertiesAtCp(Dictionary<int, PapBase> papMap, int cp, ChpBase? directChp)
    {
        var runProps = GetRunPropertiesFromParagraphStyleAtCp(papMap, cp) ?? new RunProperties { FontSize = 24 };
        if (directChp == null)
            return runProps;

        if (directChp.StyleId != 0)
        {
            var characterStyleProps = GetRunPropertiesFromCharacterStyle(directChp.StyleId);
            if (characterStyleProps != null)
            {
                runProps = MergeRunProperties(runProps, characterStyleProps);
            }
        }

        var directProps = _fkpParser!.ConvertToRunProperties(directChp, Document.Styles);
        runProps = MergeRunProperties(runProps, directProps);

        if (directChp.HasExplicitItalic)
        {
            runProps.IsItalic = directProps.IsItalic;
        }

        if (directChp.HasExplicitItalicCs)
        {
            runProps.IsItalicCs = directProps.IsItalicCs;
        }

        return runProps;
    }

    private RunProperties? GetRunPropertiesFromParagraphStyleAtCp(Dictionary<int, PapBase> papMap, int cp)
    {
        PapBase? pap = null;
        if (papMap.TryGetValue(cp, out var exactPap))
        {
            pap = exactPap;
        }
        else
        {
            for (int probe = cp - 1; probe >= Math.Max(0, cp - 2048); probe--)
            {
                if (papMap.TryGetValue(probe, out var prevPap))
                {
                    pap = prevPap;
                    break;
                }
            }
        }

        if (pap == null) return null;
        var styles = Document.Styles;
        if (styles?.Styles == null || styles.Styles.Count == 0) return null;
        var style = FindParagraphStyle(styles, pap);
        var sr = style?.RunProperties;
        if (sr == null) return null;
        return CloneRunProperties(sr);
    }

    private static StyleDefinition? FindParagraphStyle(StyleSheet styles, PapBase pap)
    {
        var styleIndex = pap.StyleId != 0 ? pap.StyleId : pap.Istd;
        if (styleIndex > 0)
        {
            var exact = styles.Styles.FirstOrDefault(s => s.Type == StyleType.Paragraph && s.StyleId == styleIndex);
            if (exact != null)
                return exact;
        }

        return styles.Styles.FirstOrDefault(s => s.Type == StyleType.Paragraph && s.IsPrimary && s.RunProperties != null)
            ?? styles.Styles.FirstOrDefault(s => s.Type == StyleType.Paragraph && s.RunProperties != null);
    }

    private RunProperties? GetRunPropertiesFromCharacterStyle(ushort styleId)
    {
        var styles = Document.Styles;
        if (styles?.Styles == null || styles.Styles.Count == 0)
            return null;

        var style = styles.Styles.FirstOrDefault(s => s.Type == StyleType.Character && s.StyleId == styleId);
        if (style?.RunProperties == null)
            return null;

        return CloneRunProperties(style.RunProperties);
    }

    private static RunProperties MergeRunProperties(RunProperties baseProps, RunProperties directProps)
    {
        var merged = CloneRunProperties(baseProps);

        if (directProps.FontIndex != -1)
            merged.FontIndex = directProps.FontIndex;
        if (!string.IsNullOrEmpty(directProps.FontName))
            merged.FontName = directProps.FontName;
        if (directProps.FontSize != 24)
            merged.FontSize = directProps.FontSize;
        if (directProps.FontSizeCs != 24)
            merged.FontSizeCs = directProps.FontSizeCs;

        merged.IsBold = directProps.IsBold || merged.IsBold;
        merged.IsBoldCs = directProps.IsBoldCs || merged.IsBoldCs;
        merged.IsItalic = directProps.IsItalic || merged.IsItalic;
        merged.IsItalicCs = directProps.IsItalicCs || merged.IsItalicCs;
        merged.IsUnderline = directProps.IsUnderline || merged.IsUnderline;
        if (directProps.UnderlineType != UnderlineType.None)
            merged.UnderlineType = directProps.UnderlineType;
        merged.IsStrikeThrough = directProps.IsStrikeThrough || merged.IsStrikeThrough;
        merged.IsDoubleStrikeThrough = directProps.IsDoubleStrikeThrough || merged.IsDoubleStrikeThrough;
        merged.IsSmallCaps = directProps.IsSmallCaps || merged.IsSmallCaps;
        merged.IsAllCaps = directProps.IsAllCaps || merged.IsAllCaps;
        merged.IsHidden = directProps.IsHidden || merged.IsHidden;
        merged.IsSuperscript = directProps.IsSuperscript || merged.IsSuperscript;
        merged.IsSubscript = directProps.IsSubscript || merged.IsSubscript;
        merged.IsOutline = directProps.IsOutline || merged.IsOutline;
        merged.IsShadow = directProps.IsShadow || merged.IsShadow;
        merged.IsEmboss = directProps.IsEmboss || merged.IsEmboss;
        merged.IsImprint = directProps.IsImprint || merged.IsImprint;
        if (directProps.Border != null)
        {
            merged.Border = new BorderInfo
            {
                Style = directProps.Border.Style,
                Width = directProps.Border.Width,
                Color = directProps.Border.Color,
                Space = directProps.Border.Space
            };
        }

        if (directProps.HasRgbColor)
        {
            merged.RgbColor = directProps.RgbColor;
            merged.HasRgbColor = true;
            merged.Color = directProps.Color;
        }
        else if (directProps.Color != 0)
        {
            merged.Color = directProps.Color;
        }

        if (directProps.BgColor != -1)
            merged.BgColor = directProps.BgColor;
        if (directProps.HighlightColor != 0)
            merged.HighlightColor = directProps.HighlightColor;
        if (directProps.CharacterSpacingAdjustment != 0)
            merged.CharacterSpacingAdjustment = directProps.CharacterSpacingAdjustment;
        if (directProps.Kerning != 0)
            merged.Kerning = directProps.Kerning;
        if (directProps.Position != 0)
            merged.Position = directProps.Position;
        if (directProps.CharacterScale != 100)
            merged.CharacterScale = directProps.CharacterScale;
        if (directProps.EastAsianLayoutType != 0)
            merged.EastAsianLayoutType = directProps.EastAsianLayoutType;
        merged.IsEastAsianVertical = directProps.IsEastAsianVertical || merged.IsEastAsianVertical;
        merged.IsEastAsianVerticalCompress = directProps.IsEastAsianVerticalCompress || merged.IsEastAsianVerticalCompress;
        if (!directProps.SnapToGrid)
            merged.SnapToGrid = false;
        if (directProps.Language != 0)
            merged.Language = directProps.Language;
        if (!string.IsNullOrEmpty(directProps.LanguageAsia))
            merged.LanguageAsia = directProps.LanguageAsia;
        if (!string.IsNullOrEmpty(directProps.LanguageCs))
            merged.LanguageCs = directProps.LanguageCs;
        if (!string.IsNullOrEmpty(directProps.RubyText))
            merged.RubyText = directProps.RubyText;

        merged.IsDeleted = directProps.IsDeleted || merged.IsDeleted;
        merged.IsInserted = directProps.IsInserted || merged.IsInserted;
        if (directProps.AuthorIndexDel != 0)
            merged.AuthorIndexDel = directProps.AuthorIndexDel;
        if (directProps.AuthorIndexIns != 0)
            merged.AuthorIndexIns = directProps.AuthorIndexIns;
        if (directProps.DateDel != 0)
            merged.DateDel = directProps.DateDel;
        if (directProps.DateIns != 0)
            merged.DateIns = directProps.DateIns;

        return merged;
    }

    private static RunProperties CloneRunProperties(RunProperties sr)
    {
        var r = new RunProperties
        {
            FontIndex = sr.FontIndex,
            FontName = sr.FontName,
            FontSize = sr.FontSize,
            FontSizeCs = sr.FontSizeCs,
            IsBold = sr.IsBold,
            IsBoldCs = sr.IsBoldCs,
            IsItalic = sr.IsItalic,
            IsItalicCs = sr.IsItalicCs,
            IsUnderline = sr.IsUnderline,
            UnderlineType = sr.UnderlineType,
            IsStrikeThrough = sr.IsStrikeThrough,
            IsDoubleStrikeThrough = sr.IsDoubleStrikeThrough,
            IsSmallCaps = sr.IsSmallCaps,
            IsAllCaps = sr.IsAllCaps,
            IsHidden = sr.IsHidden,
            IsSuperscript = sr.IsSuperscript,
            IsSubscript = sr.IsSubscript,
            Color = sr.Color,
            BgColor = sr.BgColor,
            CharacterSpacingAdjustment = sr.CharacterSpacingAdjustment,
            Language = sr.Language,
            LanguageAsia = sr.LanguageAsia,
            LanguageCs = sr.LanguageCs,
            HighlightColor = sr.HighlightColor,
            RgbColor = sr.RgbColor,
            HasRgbColor = sr.HasRgbColor,
            IsOutline = sr.IsOutline,
            IsShadow = sr.IsShadow,
            IsEmboss = sr.IsEmboss,
            IsImprint = sr.IsImprint,
            Border = sr.Border == null
                ? null
                : new BorderInfo
                {
                    Style = sr.Border.Style,
                    Width = sr.Border.Width,
                    Color = sr.Border.Color,
                    Space = sr.Border.Space
                },
            Kerning = sr.Kerning,
            Position = sr.Position
            ,CharacterScale = sr.CharacterScale
            ,EastAsianLayoutType = sr.EastAsianLayoutType
            ,IsEastAsianVertical = sr.IsEastAsianVertical
            ,IsEastAsianVerticalCompress = sr.IsEastAsianVerticalCompress
            ,SnapToGrid = sr.SnapToGrid
            ,RubyText = sr.RubyText
            ,IsDeleted = sr.IsDeleted
            ,IsInserted = sr.IsInserted
            ,AuthorIndexDel = sr.AuthorIndexDel
            ,AuthorIndexIns = sr.AuthorIndexIns
            ,DateDel = sr.DateDel
            ,DateIns = sr.DateIns
        };
        return r;
    }

    private void ApplyParagraphStyleDefaults(ParagraphModel paragraph)
    {
        if (paragraph == null) return;
        var paragraphProps = paragraph.Properties;
        if (paragraphProps == null) return;

        var styles = Document.Styles;
        if (styles == null || styles.Styles == null || styles.Styles.Count == 0)
            return;

        var style = ResolveParagraphStyle(styles, paragraphProps.StyleIndex);

        if (style == null)
            return;

        // Paragraph-level properties
        if (style.ParagraphProperties != null)
        {
            var sp = style.ParagraphProperties;

            if (paragraphProps.Alignment == ParagraphAlignment.Left &&
                sp.Alignment != ParagraphAlignment.Left)
            {
                paragraphProps.Alignment = sp.Alignment;
            }

            if (paragraphProps.IndentLeft == 0 && sp.IndentLeft != 0)
                paragraphProps.IndentLeft = sp.IndentLeft;

            if (paragraphProps.IndentLeftChars == 0 && sp.IndentLeftChars != 0)
                paragraphProps.IndentLeftChars = sp.IndentLeftChars;

            if (paragraphProps.IndentRight == 0 && sp.IndentRight != 0)
                paragraphProps.IndentRight = sp.IndentRight;

            if (paragraphProps.IndentRightChars == 0 && sp.IndentRightChars != 0)
                paragraphProps.IndentRightChars = sp.IndentRightChars;

            if (paragraphProps.IndentFirstLine == 0 && sp.IndentFirstLine != 0)
                paragraphProps.IndentFirstLine = sp.IndentFirstLine;

            if (paragraphProps.IndentFirstLineChars == 0 && sp.IndentFirstLineChars != 0)
                paragraphProps.IndentFirstLineChars = sp.IndentFirstLineChars;

            if (paragraphProps.SpaceBefore == 0 && sp.SpaceBefore != 0)
                paragraphProps.SpaceBefore = sp.SpaceBefore;

            if (paragraphProps.SpaceBefore == 0 && paragraphProps.SpaceBeforeLines == 0 && sp.SpaceBeforeLines != 0)
                paragraphProps.SpaceBeforeLines = sp.SpaceBeforeLines;

            if (paragraphProps.SpaceAfter == 0 && sp.SpaceAfter != 0)
                paragraphProps.SpaceAfter = sp.SpaceAfter;

            if (paragraphProps.SpaceAfter == 0 && paragraphProps.SpaceAfterLines == 0 && sp.SpaceAfterLines != 0)
                paragraphProps.SpaceAfterLines = sp.SpaceAfterLines;

            if (!paragraphProps.KeepWithNext && sp.KeepWithNext)
                paragraphProps.KeepWithNext = true;

            if (!paragraphProps.KeepTogether && sp.KeepTogether)
                paragraphProps.KeepTogether = true;

            if (!paragraphProps.PageBreakBefore && sp.PageBreakBefore)
                paragraphProps.PageBreakBefore = true;

            if (paragraphProps.ListFormatId == 0 && sp.ListFormatId != 0)
            {
                paragraphProps.ListFormatId = sp.ListFormatId;
                paragraphProps.ListLevel = sp.ListLevel;
            }
            else if (paragraphProps.ListFormatId != 0 && paragraphProps.ListLevel == 0 && sp.ListFormatId == paragraphProps.ListFormatId)
            {
                paragraphProps.ListLevel = sp.ListLevel;
            }

            paragraphProps.BorderTop ??= sp.BorderTop;
            paragraphProps.BorderBottom ??= sp.BorderBottom;
            paragraphProps.BorderLeft ??= sp.BorderLeft;
            paragraphProps.BorderRight ??= sp.BorderRight;

            if (!paragraphProps.HasExplicitLineSpacing && sp.HasExplicitLineSpacing)
            {
                paragraphProps.LineSpacing = sp.LineSpacing;
                paragraphProps.LineSpacingMultiple = sp.LineSpacingMultiple;
                paragraphProps.HasExplicitLineSpacing = true;
            }
            else if (!paragraphProps.HasExplicitLineSpacing && paragraphProps.LineSpacing == 240 && sp.LineSpacing != 240)
            {
                paragraphProps.LineSpacing = sp.LineSpacing;
                paragraphProps.LineSpacingMultiple = sp.LineSpacingMultiple;
            }
        }

        // Visible text runs already receive paragraph-style defaults during CHP
        // extraction. Reapplying boolean defaults here can incorrectly force
        // style formatting like italic back onto runs that explicitly clear it.
        if (style.RunProperties == null || paragraph.Runs == null || paragraph.Runs.Count == 0)
            return;

        var sr = style.RunProperties;

        if (paragraph.ListFormatId == 0 && paragraphProps.ListFormatId != 0)
            paragraph.ListFormatId = paragraphProps.ListFormatId;
        if (paragraph.ListFormatId != 0)
            paragraph.ListLevel = paragraphProps.ListLevel;

        foreach (var run in paragraph.Runs)
        {
            if (run.Properties == null)
            {
                run.Properties = CloneRunProperties(sr);
            }

            var rp = run.Properties;
            ApplyEastAsiaDefaultFont(run, rp);
        }
    }

    private void ApplyEastAsiaDefaultFont(RunModel run, RunProperties properties)
    {
        if (!string.IsNullOrEmpty(properties.FontName))
            return;

        if (!ContainsEastAsianText(run.Text) && string.IsNullOrEmpty(properties.LanguageAsia) && properties.Language != 0x0804)
            return;

        properties.FontName = Document.Theme.MinorEastAsiaFont
            ?? Document.Theme.MajorEastAsiaFont
            ?? "SimSun";
    }

    private static bool ContainsEastAsianText(string? text)
    {
        if (string.IsNullOrEmpty(text))
            return false;

        foreach (var c in text)
        {
            if ((c >= '\u4E00' && c <= '\u9FFF') ||
                (c >= '\u3400' && c <= '\u4DBF') ||
                (c >= '\u3000' && c <= '\u303F') ||
                (c >= '\u3040' && c <= '\u30FF') ||
                (c >= '\uAC00' && c <= '\uD7AF'))
            {
                return true;
            }
        }

        return false;
    }

    private static StyleDefinition? ResolveParagraphStyle(StyleSheet styles, int styleIndex)
    {
        if (styleIndex == 0)
        {
            return styles.Styles
                       .Where(s =>
                           s.Type == StyleType.Paragraph &&
                           string.Equals(s.Name, "Normal", StringComparison.OrdinalIgnoreCase))
                       .OrderBy(s => s.StyleId == 0 ? 1 : 0)
                       .FirstOrDefault()
                   ?? styles.Styles.FirstOrDefault(s => s.Type == StyleType.Paragraph && s.StyleId == StyleIds.NORMAL)
                   ?? styles.Styles.FirstOrDefault(s => s.Type == StyleType.Paragraph && s.StyleId == 0);
        }

        var exactMatch = styles.Styles.FirstOrDefault(s => s.Type == StyleType.Paragraph && s.StyleId == styleIndex);
        if (exactMatch != null)
            return exactMatch;

        return null;
    }

    /// <summary>
    /// Removes Word special control characters from text for display.
    /// </summary>
    private static string CleanSpecialChars(string text)
    {
        if (string.IsNullOrEmpty(text)) return text;

        var sb = new StringBuilder(text.Length);
        foreach (var ch in text)
        {
            switch (ch)
            {
                case '\x01': // SOH — field begin/end or inline picture
                case '\x07': // BEL — cell mark
                case '\x08': // BS  — drawn object
                case '\x13': // field begin
                case '\x14': // field separator
                case '\x15': // field end
                    break; // skip these

                case '\x0B': // vertical tab → line break
                    sb.Append('\n');
                    break;

                case '\x0C': // form feed → page break (keep as text for now)
                    break;

                case '\x1E': // non-breaking hyphen
                    sb.Append('-');
                    break;

                case '\x1F': // optional hyphen
                    break; // skip

                case '\xA0': // non-breaking space
                    sb.Append(' ');
                    break;

                default:
                    if (!char.IsControl(ch) || ch == '\t')
                        sb.Append(ch);
                    break;
            }
        }
        return sb.ToString();
    }

    /// <summary>Gets the text reader for direct access</summary>
    public TextReader GetTextReader() => _textReader!;

    /// <summary>Gets the FIB reader for direct access</summary>
    public FibReader GetFibReader() => _fibReader!;

    /// <summary>Gets the style reader for direct access</summary>
    public StyleReader GetStyleReader() => _styleReader!;

    /// <summary>Gets the CFB reader for diagnostics</summary>
    public CfbReader GetCfbReader() => _cfb!;

    private void ExtractVbaProject()
    {
        if (_cfb == null) return;

        try
        {
            if (_cfb.HasStorage("Macros"))
            {
                var macrosStorage = _cfb.GetStorage("Macros");
                // Repack the Macros storage into a standalone OLE compound file 
                // representing vbaProject.bin
                if (macrosStorage != null)
                {
                    Document.VbaProject = Writers.CfbBuilder.RepackStorage(_cfb, macrosStorage);
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Warning("Failed to extract VBA project storage.", ex);
        }
    }
    private static void OverlayParagraphProperties(ParagraphProperties target, PapBase overlay)
    {
        if (overlay.StyleId != 0)
            target.StyleIndex = overlay.StyleId;
        else if (target.StyleIndex == 0 && overlay.Istd != 0)
            target.StyleIndex = overlay.Istd;

        if (overlay.Justification != 0)
            target.Alignment = (ParagraphAlignment)overlay.Justification;

        if (overlay.IndentLeft != 0)
            target.IndentLeft = overlay.IndentLeft;
        if (overlay.IndentLeftChars != 0)
            target.IndentLeftChars = overlay.IndentLeftChars;
        if (overlay.IndentRight != 0)
            target.IndentRight = overlay.IndentRight;
        if (overlay.IndentRightChars != 0)
            target.IndentRightChars = overlay.IndentRightChars;
        if (overlay.IndentFirstLine != 0)
            target.IndentFirstLine = overlay.IndentFirstLine;
        if (overlay.IndentFirstLineChars != 0)
            target.IndentFirstLineChars = overlay.IndentFirstLineChars;
        if (overlay.SpaceBefore != 0)
        {
            target.SpaceBefore = overlay.SpaceBefore;
            target.SpaceBeforeLines = 0;
        }
        if (overlay.SpaceBeforeLines != 0)
        {
            target.SpaceBeforeLines = overlay.SpaceBeforeLines;
            target.SpaceBefore = 0;
        }
        if (overlay.SpaceAfter != 0)
        {
            target.SpaceAfter = overlay.SpaceAfter;
            target.SpaceAfterLines = 0;
        }
        if (overlay.SpaceAfterLines != 0)
        {
            target.SpaceAfterLines = overlay.SpaceAfterLines;
            target.SpaceAfter = 0;
        }
        if (overlay.HasExplicitLineSpacing)
        {
            target.LineSpacing = overlay.LineSpacing;
            target.LineSpacingMultiple = overlay.LineSpacingMultiple;
            target.HasExplicitLineSpacing = true;
        }
        else if (!target.HasExplicitLineSpacing && (overlay.LineSpacing != 240 || overlay.LineSpacingMultiple != 1))
        {
            target.LineSpacing = overlay.LineSpacing;
            target.LineSpacingMultiple = overlay.LineSpacingMultiple;
        }
        if (overlay.KeepWithNext)
            target.KeepWithNext = true;
        if (overlay.KeepTogether)
            target.KeepTogether = true;
        if (overlay.PageBreakBefore)
            target.PageBreakBefore = true;
        if (overlay.BorderTop != null)
            target.BorderTop = overlay.BorderTop;
        if (overlay.BorderBottom != null)
            target.BorderBottom = overlay.BorderBottom;
        if (overlay.BorderLeft != null)
            target.BorderLeft = overlay.BorderLeft;
        if (overlay.BorderRight != null)
            target.BorderRight = overlay.BorderRight;
        if (overlay.ListFormatId != 0)
            target.ListFormatId = overlay.ListFormatId;
        if (overlay.ListLevel != 0)
            target.ListLevel = overlay.ListLevel;
        if (overlay.OutlineLevel != 9)
            target.OutlineLevel = overlay.OutlineLevel;
        if (overlay.Shading != null)
            target.Shading = overlay.Shading;
    }

    private static void OverlayParagraphPropertiesFromRange(ParagraphProperties target, Dictionary<int, PapBase> papMap, int paragraphStartCp, int paragraphEndCp)
    {
        if (papMap.Count == 0)
            return;

        foreach (var entry in papMap
            .Where(kvp => kvp.Key >= paragraphStartCp && kvp.Key <= paragraphEndCp)
            .OrderBy(kvp => kvp.Key))
        {
            OverlayParagraphProperties(target, entry.Value);
        }
    }

    private static void ApplyPictureShapeDisplaySizes(DocumentModel document)
    {
        if (document.Shapes == null || document.Shapes.Count == 0)
            return;

        var pictureRuns = EnumeratePictureRuns(document)
            .Where(run => run.ImageIndex >= 0)
            .ToList();
        if (pictureRuns.Count == 0)
            return;

        var shapesByImageIndex = document.Shapes
            .Where(shape =>
                shape.Type == ShapeType.Picture &&
                shape.ImageIndex is not null)
            .GroupBy(shape => shape.ImageIndex!.Value)
            .ToDictionary(
                group => group.Key,
                group => new Queue<ShapeModel>(group
                    .OrderBy(shape => shape.Anchor?.ParagraphIndex >= 0 ? shape.Anchor.ParagraphIndex : shape.ParagraphIndexHint)
                    .ThenBy(shape => shape.Anchor?.ZOrder ?? int.MaxValue)
                    .ThenBy(shape => shape.Id)));

        if (shapesByImageIndex.Count == 0)
            return;

        foreach (var run in pictureRuns)
        {
            if (run.ImageIndex < 0)
                continue;

            if (!shapesByImageIndex.TryGetValue(run.ImageIndex, out var shapes) || shapes.Count == 0)
                continue;

            var shape = shapes.Dequeue();
            if (shape.Anchor != null && shape.Anchor.Width > 0 && shape.Anchor.Height > 0)
            {
                run.DisplayWidthTwips = shape.Anchor.Width;
                run.DisplayHeightTwips = shape.Anchor.Height;
            }

            run.FlipHorizontal = run.FlipHorizontal || shape.FlipHorizontal;
            run.FlipVertical = run.FlipVertical || shape.FlipVertical;
        }

        var orderedShapes = document.Shapes
            .Where(shape => shape.Type == ShapeType.Picture)
            .OrderBy(shape => shape.ParagraphIndexHint >= 0 ? shape.ParagraphIndexHint : int.MaxValue)
            .ThenBy(shape => shape.Anchor?.ParagraphIndex ?? int.MaxValue)
            .ThenBy(shape => shape.Anchor?.ZOrder ?? int.MaxValue)
            .ThenBy(shape => shape.Id)
            .ToList();

        if (orderedShapes.Count != pictureRuns.Count)
            return;

        for (int index = 0; index < pictureRuns.Count; index++)
        {
            var run = pictureRuns[index];
            var shape = orderedShapes[index];

            run.FlipHorizontal = run.FlipHorizontal || shape.FlipHorizontal;
            run.FlipVertical = run.FlipVertical || shape.FlipVertical;

            if ((run.DisplayWidthTwips <= 0 || run.DisplayHeightTwips <= 0) &&
                shape.Anchor != null && shape.Anchor.Width > 0 && shape.Anchor.Height > 0)
            {
                run.DisplayWidthTwips = shape.Anchor.Width;
                run.DisplayHeightTwips = shape.Anchor.Height;
            }
        }
    }

    private static IEnumerable<RunModel> EnumeratePictureRuns(DocumentModel document)
    {
        foreach (var para in document.Paragraphs)
            foreach (var run in para.Runs)
                if (run.IsPicture) yield return run;
        foreach (var table in document.Tables)
            foreach (var row in table.Rows)
                foreach (var cell in row.Cells)
                    foreach (var para in cell.Paragraphs)
                        foreach (var run in para.Runs)
                            if (run.IsPicture) yield return run;
        foreach (var note in document.Footnotes)
            foreach (var para in note.Paragraphs)
                foreach (var run in para.Runs)
                    if (run.IsPicture) yield return run;
        foreach (var note in document.Endnotes)
            foreach (var para in note.Paragraphs)
                foreach (var run in para.Runs)
                    if (run.IsPicture) yield return run;
        foreach (var textbox in document.Textboxes)
        {
            if (textbox.Paragraphs == null) continue;
            foreach (var para in textbox.Paragraphs)
                foreach (var run in para.Runs)
                    if (run.IsPicture) yield return run;
        }
    }

    public void Dispose()
    {
        _wordDocReader?.Dispose();
        _tableReader?.Dispose();
        _dataReader?.Dispose();
        _footnoteStream?.Dispose();
        _endnoteStream?.Dispose();
        _anotStream?.Dispose();
        _txbxStream?.Dispose();
        _wordDocStream?.Dispose();
        _tableStream?.Dispose();
        _dataStream?.Dispose();
        _cfb?.Dispose();
    }
}

/// <summary>
/// Image reader — extracts images from Word documents.
/// Phase 1: stub implementation. Phase 3 will parse OfficeArt/BLIP data.
/// </summary>
public class ImageReader
{
    private readonly BinaryReader _wordDocReader;
    private readonly BinaryReader? _dataReader;
    private readonly FibReader _fib;
    private readonly CfbReader? _cfb;

    public ImageReader(BinaryReader wordDocReader, BinaryReader? dataReader, FibReader fib, CfbReader? cfb = null)
    {
        _wordDocReader = wordDocReader;
        _dataReader = dataReader;
        _fib = fib;
        _cfb = cfb;
    }

    /// <summary>
    /// Extracts images from the document.
    /// First extracts at sprmCPicLocation (FcPic) offsets for inline pictures, then scans for BLIPs.
    /// </summary>
    public void ExtractImages(DocumentModel document)
    {
        document.Images = new List<ImageModel>();

        if (_dataReader == null) return;

        try
        {
            var data = _dataReader.BaseStream;
            if (data.Length < 8) return;

            data.Seek(0, SeekOrigin.Begin);
            var buffer = new byte[data.Length];
            _dataReader.Read(buffer, 0, (int)data.Length);

            var extractedRanges = new HashSet<(int start, int end)>();

            // 1. Extract images at sprmCPicLocation (FcPic) offsets — inline pictures in document order
            ExtractImagesAtFcPicPositions(document, buffer, extractedRanges);

            // 2. Scan for additional BLIPs (floating shapes, etc.), skipping already-extracted ranges
            ScanForImageSignatures(buffer, document.Images, extractedRanges);
            ScanForOfficeArtRecords(buffer, document.Images, extractedRanges);

            // 3. Scan other OLE streams (WordDocument, Table, ObjectPool children) for embedded images
            if (_cfb != null)
                ScanAdditionalStreamsForImages(document.Images);

            // 4. Ensure every picture run has a valid ImageIndex (assign 0,1,2... in document order)
            AssignPictureRunIndices(document);

            // 5. PICF carries per-occurrence display size for repeated inline pictures.
            ApplyPictureRunDisplaySizes(document, buffer);
        }
        catch (Exception ex)
        {
            Logger.Warning("Image extraction failed; continuing with partial image recovery.", ex);
        }
    }

    /// <summary>Extracts images at FcPic positions and assigns run.ImageIndex.</summary>
    private void ExtractImagesAtFcPicPositions(DocumentModel document, byte[] buffer, HashSet<(int start, int end)> extractedRanges)
    {
        // Reset all picture run indices; we will set only for successful FcPic extractions
        foreach (var para in document.Paragraphs)
            foreach (var run in para.Runs)
                if (run.IsPicture) run.ImageIndex = -1;
        foreach (var table in document.Tables)
            foreach (var row in table.Rows)
                foreach (var cell in row.Cells)
                    foreach (var para in cell.Paragraphs)
                        foreach (var run in para.Runs)
                            if (run.IsPicture) run.ImageIndex = -1;
        foreach (var note in document.Footnotes)
            foreach (var para in note.Paragraphs)
                foreach (var run in para.Runs)
                    if (run.IsPicture) run.ImageIndex = -1;
        foreach (var note in document.Endnotes)
            foreach (var para in note.Paragraphs)
                foreach (var run in para.Runs)
                    if (run.IsPicture) run.ImageIndex = -1;
        foreach (var textbox in document.Textboxes)
        {
            if (textbox.Paragraphs == null) continue;
            foreach (var para in textbox.Paragraphs)
                foreach (var run in para.Runs)
                    if (run.IsPicture) run.ImageIndex = -1;
        }

        foreach (var para in document.Paragraphs)
        {
            foreach (var run in para.Runs)
            {
                if (!run.IsPicture || run.FcPic == 0) continue;
                // MS-DOC: sprmCPicLocation is position in Data stream (byte offset). Treat as unsigned.
                if (!TryFcPicToBufferOffset(run.FcPic, buffer.Length, out int offset))
                { run.ImageIndex = -1; continue; }

                var img = TryExtractImageAtOffset(buffer, offset);
                if (img != null)
                {
                    img.PictureOffset = offset;
                    run.ImageIndex = document.Images.Count;
                    document.Images.Add(img);
                    if (img.Data != null)
                        extractedRanges.Add((offset, Math.Min(offset + 68 + img.Data.Length + 512, buffer.Length)));
                }
                else
                    run.ImageIndex = -1;
            }
        }
        foreach (var table in document.Tables)
        {
            foreach (var row in table.Rows)
            {
                foreach (var cell in row.Cells)
                {
                    foreach (var para in cell.Paragraphs)
                    {
                        foreach (var run in para.Runs)
                        {
                            if (!run.IsPicture || run.FcPic == 0) continue;
                            if (!TryFcPicToBufferOffset(run.FcPic, buffer.Length, out int offset))
                            { run.ImageIndex = -1; continue; }

                            var img = TryExtractImageAtOffset(buffer, offset);
                            if (img != null)
                            {
                                img.PictureOffset = offset;
                                run.ImageIndex = document.Images.Count;
                                document.Images.Add(img);
                                if (img.Data != null)
                                    extractedRanges.Add((offset, Math.Min(offset + 68 + img.Data.Length + 512, buffer.Length)));
                            }
                            else
                                run.ImageIndex = -1;
                        }
                    }
                }
            }
        }

        // Footnotes
        foreach (var note in document.Footnotes)
        {
            foreach (var para in note.Paragraphs)
            {
                foreach (var run in para.Runs)
                {
                    if (!run.IsPicture || run.FcPic == 0) continue;
                    if (!TryFcPicToBufferOffset(run.FcPic, buffer.Length, out int offset))
                    { run.ImageIndex = -1; continue; }
                    var img = TryExtractImageAtOffset(buffer, offset);
                    if (img != null)
                    {
                        run.ImageIndex = document.Images.Count;
                        document.Images.Add(img);
                        if (img.Data != null)
                            extractedRanges.Add((offset, Math.Min(offset + 68 + img.Data.Length + 512, buffer.Length)));
                    }
                    else
                        run.ImageIndex = -1;
                }
            }
        }
        // Endnotes
        foreach (var note in document.Endnotes)
        {
            foreach (var para in note.Paragraphs)
            {
                foreach (var run in para.Runs)
                {
                    if (!run.IsPicture || run.FcPic == 0) continue;
                    if (!TryFcPicToBufferOffset(run.FcPic, buffer.Length, out int offset))
                    { run.ImageIndex = -1; continue; }
                    var img = TryExtractImageAtOffset(buffer, offset);
                    if (img != null)
                    {
                        img.PictureOffset = offset;
                        run.ImageIndex = document.Images.Count;
                        document.Images.Add(img);
                        if (img.Data != null)
                            extractedRanges.Add((offset, Math.Min(offset + 68 + img.Data.Length + 512, buffer.Length)));
                    }
                    else
                        run.ImageIndex = -1;
                }
            }
        }
        // Textboxes
        foreach (var textbox in document.Textboxes)
        {
            if (textbox.Paragraphs == null) continue;
            foreach (var para in textbox.Paragraphs)
            {
                foreach (var run in para.Runs)
                {
                    if (!run.IsPicture || run.FcPic == 0) continue;
                    if (!TryFcPicToBufferOffset(run.FcPic, buffer.Length, out int offset))
                    { run.ImageIndex = -1; continue; }
                    var img = TryExtractImageAtOffset(buffer, offset);
                    if (img != null)
                    {
                        img.PictureOffset = offset;
                        run.ImageIndex = document.Images.Count;
                        document.Images.Add(img);
                        if (img.Data != null)
                            extractedRanges.Add((offset, Math.Min(offset + 68 + img.Data.Length + 512, buffer.Length)));
                    }
                    else
                        run.ImageIndex = -1;
                }
            }
        }
    }

    /// <summary>Maps FcPic (signed in spec but stored as uint) to a valid buffer offset. Treats value as unsigned byte offset.</summary>
    private static bool TryFcPicToBufferOffset(uint fcPic, int bufferLength, out int offset)
    {
        if (fcPic == 0) { offset = 0; return false; }
        // Use as unsigned; max Data stream size is 0x7FFFFFFF so valid offset fits in int.
        if (fcPic > (uint)int.MaxValue || fcPic >= (uint)bufferLength) { offset = 0; return false; }
        offset = (int)fcPic;
        return true;
    }

    /// <summary>Tries FcPic as byte offset, then alternate interpretations (FC*2, FC/2, and direct OfficeArt at offset).</summary>
    private ImageModel? TryExtractImageAtOffset(byte[] buffer, int offset)
    {
        // 1) Standard: PICF at offset
        if (offset >= 0 && offset < buffer.Length)
        {
            var img = ExtractImageAtPicfOffset(buffer, offset);
            if (img != null) return img;
        }
        // 2) FcPic in 32-bit FC units (byte offset = FcPic*2)
        long offsetAlt = (long)offset * 2;
        if (offsetAlt > 0 && offsetAlt < buffer.Length)
        {
            var img = ExtractImageAtPicfOffset(buffer, (int)offsetAlt);
            if (img != null) return img;
        }
        // 3) Half-byte offset
        int offsetHalf = offset / 2;
        if (offsetHalf >= 0 && offsetHalf < buffer.Length)
        {
            var img = ExtractImageAtPicfOffset(buffer, offsetHalf);
            if (img != null) return img;
        }
        // 4) Some files: FcPic points directly to OfficeArt (no PICF); try BLIP/signature at offset
        if (offset >= 0 && offset < buffer.Length)
        {
            var img = ExtractBlipFromOfficeArtRegion(buffer, offset, buffer.Length - offset);
            if (img != null) return img;
            var scan = ScanForImageAtPosition(buffer, offset);
            if (scan != null) return scan;
        }
        return null;
    }

    /// <summary>Extracts image from PICFAndOfficeArtData at the given Data stream offset.</summary>
    private ImageModel? ExtractImageAtPicfOffset(byte[] buffer, int offset)
    {
        if (offset + 68 > buffer.Length) return null;
        // PICF: lcb(4)+cbHeader(2)+mfpf(8)+...; mfpf.mm at offset 6
        var mm = offset + 8 <= buffer.Length ? BitConverter.ToUInt16(buffer, offset + 6) : (ushort)0;
        var picfEnd = 68;
        if (mm == 0x0066 && offset + 69 <= buffer.Length)
        {
            var cchPicName = buffer[offset + 68];
            picfEnd = 69 + cchPicName;
        }
        if (offset + picfEnd >= buffer.Length) return null;

        var artStart = offset + picfEnd;
        // OfficeArtInlineSpContainer: shape then rgfb (BLIPs); recurse into containers
        var img = ExtractBlipFromOfficeArtRegion(buffer, artStart, buffer.Length - artStart);
        if (img != null) return img;
        var scan = ScanForImageAtPosition(buffer, artStart);
        if (scan != null) return scan;

        // Fallback: search a window for BLIP or raw signature (alignment/variant PICF)
        for (int delta = 8; delta <= 128 && artStart - delta >= 0; delta += 8)
        {
            int winStart = artStart - delta;
            int winLen = Math.Min(4096, buffer.Length - winStart);
            if (winLen >= 8)
            {
                img = ExtractBlipFromOfficeArtRegion(buffer, winStart, winLen);
                if (img != null) return img;
            }
        }
        int searchEnd = Math.Min(artStart + 1024, buffer.Length - 8);
        for (int p = artStart; p < searchEnd; p++)
        {
            scan = ScanForImageAtPosition(buffer, p);
            if (scan != null) return scan;
        }
        return null;
    }

    /// <summary>Searches for a BLIP or image within an OfficeArt region; recurses into containers.</summary>
    private ImageModel? ExtractBlipFromOfficeArtRegion(byte[] buffer, int start, int maxLen)
    {
        int pos = 0;
        while (pos < maxLen - 8 && start + pos < buffer.Length)
        {
            var recType = BitConverter.ToUInt16(buffer, start + pos + 2);
            var recLen = BitConverter.ToUInt32(buffer, start + pos + 4);
            if (recLen > 50 * 1024 * 1024 ||
                pos + 8 + recLen > (uint)maxLen ||
                start + pos + 8 + recLen > buffer.Length)
            {
                pos++;
                continue;
            }
            // BLIP types per MS-ODRAW 2.2.23
            if (recType >= 0xF018 && recType <= 0xF117)
            {
                var img = ExtractBlipFromOfficeArtRecord(buffer, start + pos + 8, (int)recLen, recType);
                if (img != null) return img;
            }
            // Recurse into containers: DggContainer, DgContainer, SpgrContainer, SpContainer, BStoreContainer
            if (recType == 0xF000 || recType == 0xF001 || recType == 0xF002 || recType == 0xF003 || recType == 0xF004 || recType == 0xF007)
            {
                var innerLen = (int)recLen;
                if (innerLen > 0 && start + pos + 8 + innerLen <= buffer.Length)
                {
                    var img = ExtractBlipFromOfficeArtRegion(buffer, start + pos + 8, innerLen);
                    if (img != null) return img;
                }
            }
            pos += 8 + (int)recLen;
        }
        return null;
    }

    /// <summary>Scans for raw image signature at the given position.</summary>
    private ImageModel? ScanForImageAtPosition(byte[] buffer, int pos)
    {
        if (pos + 8 > buffer.Length) return null;
        ImageType? type = null;
        if (buffer[pos] == 0x89 && buffer[pos + 1] == 0x50 && buffer[pos + 2] == 0x4E && buffer[pos + 3] == 0x47)
            type = ImageType.Png;
        else if (buffer[pos] == 0xFF && buffer[pos + 1] == 0xD8 && buffer[pos + 2] == 0xFF)
            type = ImageType.Jpeg;
        else if (buffer[pos] == 0x47 && buffer[pos + 1] == 0x49 && buffer[pos + 2] == 0x46)
            type = ImageType.Gif;
        else if (buffer[pos] == 0x42 && buffer[pos + 1] == 0x4D)
            type = ImageType.Dib;
        if (!type.HasValue) return null;
        return ExtractImageFromPosition(buffer, pos, type.Value);
    }

    /// <summary>
    /// Scans the Data stream for OfficeArt BLIP records.
    /// </summary>
    private List<ImageModel> ScanForBlipRecords()
    {
        var images = new List<ImageModel>();
        var data = _dataReader!.BaseStream;
        var length = data.Length;

        if (length < 8) return images;

        data.Seek(0, SeekOrigin.Begin);
        var buffer = new byte[length];
        _dataReader.Read(buffer, 0, (int)length);

        ScanForImageSignatures(buffer, images);
        ScanForOfficeArtRecords(buffer, images);

        return images;
    }

    /// <summary>
    /// Scans for raw image signatures (PNG, JPEG, GIF, etc.) in the data.
    /// </summary>
    private void ScanForImageSignatures(byte[] buffer, List<ImageModel> images, HashSet<(int start, int end)>? skipRanges = null)
    {
        int pos = 0;
        while (pos < buffer.Length - 8)
        {
            if (IsInSkipRange(pos, skipRanges)) { pos++; continue; }

            ImageType? type = null;
            int headerLen = 0;

            // Check for PNG signature
            if (pos + 8 <= buffer.Length &&
                buffer[pos] == 0x89 && buffer[pos + 1] == 0x50 &&
                buffer[pos + 2] == 0x4E && buffer[pos + 3] == 0x47 &&
                buffer[pos + 4] == 0x0D && buffer[pos + 5] == 0x0A &&
                buffer[pos + 6] == 0x1A && buffer[pos + 7] == 0x0A)
            {
                type = ImageType.Png;
                headerLen = 8;
            }
            // Check for JPEG signature
            else if (pos + 3 <= buffer.Length &&
                     buffer[pos] == 0xFF && buffer[pos + 1] == 0xD8 && buffer[pos + 2] == 0xFF)
            {
                type = ImageType.Jpeg;
                headerLen = 3;
            }
            // Check for GIF signature
            else if (pos + 6 <= buffer.Length &&
                     buffer[pos] == 0x47 && buffer[pos + 1] == 0x49 && buffer[pos + 2] == 0x46 &&
                     buffer[pos + 3] == 0x38 && (buffer[pos + 4] == 0x37 || buffer[pos + 4] == 0x39) &&
                     buffer[pos + 5] == 0x61)
            {
                type = ImageType.Gif;
                headerLen = 6;
            }
            // Check for BMP signature
            else if (pos + 2 <= buffer.Length && buffer[pos] == 0x42 && buffer[pos + 1] == 0x4D)
            {
                type = ImageType.Dib;
                headerLen = 2;
            }

            if (type.HasValue)
            {
                var image = ExtractImageFromPosition(buffer, pos, type.Value);
                if (image != null)
                {
                    images.Add(image);
                    // Mark range as extracted to avoid duplicate from OfficeArt scan
                    skipRanges?.Add((pos, Math.Min(pos + image.Data.Length, buffer.Length)));
                    pos += image.Data.Length;
                    continue;
                }
            }

            pos++;
        }
    }

    /// <summary>Ensures every picture run has a valid ImageIndex. Preserves indices set by FcPic extraction; for runs that failed extraction, picks the image whose PictureOffset is closest to run.FcPic so the correct image (e.g. first-page background) is shown.</summary>
    private static void AssignPictureRunIndices(DocumentModel document)
    {
        int maxIdx = Math.Max(0, document.Images.Count - 1);
        var assigned = new HashSet<int>();
        foreach (var run in EnumeratePictureRuns(document))
            if (run.ImageIndex >= 0)
                assigned.Add(run.ImageIndex);

        int fallbackIdx = 0;
        foreach (var run in EnumeratePictureRuns(document))
        {
            if (run.ImageIndex >= 0) continue;
            int chosen = -1;
            if (run.FcPic != 0 && document.Images.Count > 0)
            {
                long bestDist = long.MaxValue;
                for (int i = 0; i < document.Images.Count; i++)
                {
                    if (assigned.Contains(i)) continue;
                    int po = document.Images[i].PictureOffset;
                    if (po == 0) continue;
                    long d = Math.Abs((long)po - run.FcPic);
                    if (d < bestDist)
                    {
                        bestDist = d;
                        chosen = i;
                    }
                }
            }
            if (chosen < 0)
            {
                while (fallbackIdx <= maxIdx && assigned.Contains(fallbackIdx)) fallbackIdx++;
                chosen = fallbackIdx <= maxIdx ? fallbackIdx++ : 0;
            }
            run.ImageIndex = chosen;
            assigned.Add(chosen);
        }
    }

    private static IEnumerable<RunModel> EnumeratePictureRuns(DocumentModel document)
    {
        foreach (var para in document.Paragraphs)
            foreach (var run in para.Runs)
                if (run.IsPicture) yield return run;
        foreach (var table in document.Tables)
            foreach (var row in table.Rows)
                foreach (var cell in row.Cells)
                    foreach (var para in cell.Paragraphs)
                        foreach (var run in para.Runs)
                            if (run.IsPicture) yield return run;
        foreach (var note in document.Footnotes)
            foreach (var para in note.Paragraphs)
                foreach (var run in para.Runs)
                    if (run.IsPicture) yield return run;
        foreach (var note in document.Endnotes)
            foreach (var para in note.Paragraphs)
                foreach (var run in para.Runs)
                    if (run.IsPicture) yield return run;
        foreach (var textbox in document.Textboxes)
        {
            if (textbox.Paragraphs == null) continue;
            foreach (var para in textbox.Paragraphs)
                foreach (var run in para.Runs)
                    if (run.IsPicture) yield return run;
        }
    }

    private static void ApplyPictureRunDisplaySizes(DocumentModel document, byte[] buffer)
    {
        int maxContentWidthTwips = 0;
        if (document.Properties != null)
        {
            maxContentWidthTwips = Math.Max(0, document.Properties.PageWidth - document.Properties.MarginLeft - document.Properties.MarginRight);
        }

        foreach (var run in EnumeratePictureRuns(document))
        {
            run.DisplayWidthTwips = 0;
            run.DisplayHeightTwips = 0;

            if (!run.IsPicture || !TryFcPicToBufferOffset(run.FcPic, buffer.Length, out int offset))
                continue;

            if (!TryReadPicfDisplaySize(buffer, offset, maxContentWidthTwips, out int widthTwips, out int heightTwips, out bool flipHorizontal, out bool flipVertical))
                continue;

            run.DisplayWidthTwips = widthTwips;
            run.DisplayHeightTwips = heightTwips;
            run.FlipHorizontal = flipHorizontal;
            run.FlipVertical = flipVertical;
        }
    }

    private static bool TryReadPicfDisplaySize(byte[] buffer, int offset, int maxContentWidthTwips, out int widthTwips, out int heightTwips, out bool flipHorizontal, out bool flipVertical)
    {
        widthTwips = 0;
        heightTwips = 0;
        flipHorizontal = false;
        flipVertical = false;

        if (offset < 0 || offset + 36 > buffer.Length)
            return false;

        ushort cbHeader = BitConverter.ToUInt16(buffer, offset + 4);
        if (cbHeader < 36 || offset + cbHeader > buffer.Length)
            return false;

        int dxaGoal = BitConverter.ToUInt16(buffer, offset + 28);
        int dyaGoal = BitConverter.ToUInt16(buffer, offset + 30);
        int mx = BitConverter.ToInt16(buffer, offset + 32);
        int my = BitConverter.ToInt16(buffer, offset + 34);

        if (dxaGoal <= 0 || dyaGoal <= 0)
            return false;

        double width = dxaGoal;
        double height = dyaGoal;

        if (mx < 0)
        {
            flipHorizontal = true;
        }
        if (my < 0)
        {
            flipVertical = true;
        }

        int scaleX = Math.Abs(mx);
        int scaleY = Math.Abs(my);

        if (scaleX > 0 && scaleX != 1000)
            width = width * scaleX / 1000d;
        if (scaleY > 0 && scaleY != 1000)
            height = height * scaleY / 1000d;

        if (maxContentWidthTwips > 0 && width > maxContentWidthTwips)
            width = maxContentWidthTwips;

        widthTwips = (int)Math.Round(width, MidpointRounding.AwayFromZero);
        heightTwips = (int)Math.Round(height, MidpointRounding.AwayFromZero);
        return widthTwips > 0 && heightTwips > 0;
    }

    /// <summary>Scans WordDocument, Table, and ObjectPool streams for embedded images.</summary>
    private void ScanAdditionalStreamsForImages(List<ImageModel> images)
    {
        var names = new[] { "WordDocument", "0Table", "1Table" };
        foreach (var name in names)
        {
            if (!_cfb!.HasStream(name)) continue;
            try
            {
                var bytes = _cfb.GetStreamBytes(name);
                ScanForImageSignatures(bytes, images);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to scan stream '{name}' for embedded images.", ex);
            }
        }
        // Scan other top-level streams that might contain embedded images (e.g. OLE embeddings)
        foreach (var name in _cfb!.StreamNames)
        {
            if (name is "WordDocument" or "0Table" or "1Table" or "Data") continue;
            if (name.Length < 2) continue;
            try
            {
                var bytes = _cfb.GetStreamBytes(name);
                if (bytes.Length > 8 && bytes.Length < 50 * 1024 * 1024)
                    ScanForImageSignatures(bytes, images);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to scan auxiliary stream '{name}' for embedded images.", ex);
            }
        }
    }

    private static bool IsInSkipRange(int pos, HashSet<(int start, int end)>? ranges)
    {
        if (ranges == null) return false;
        foreach (var (start, end) in ranges)
            if (pos >= start && pos < end) return true;
        return false;
    }


    private void ScanForOfficeArtRecords(byte[] buffer, List<ImageModel> images, HashSet<(int start, int end)>? skipRanges = null)
    {
        // OfficeArt record header format:
        // - recVer (4 bits) + recInstance (12 bits) = 2 bytes
        // - recType (2 bytes)
        // - recLen (4 bytes)

        int pos = 0;
        while (pos < buffer.Length - 8)
        {
            if (IsInSkipRange(pos, skipRanges)) { pos++; continue; }
            try
            {
                var recInfo = BitConverter.ToUInt16(buffer, pos);
                var recVer = recInfo & 0x0F;
                var recInstance = (recInfo >> 4) & 0x0FFF;
                var recType = BitConverter.ToUInt16(buffer, pos + 2);
                var recLen = BitConverter.ToUInt32(buffer, pos + 4);

                // Validate record
                if (recLen > 0 && recLen < 100 * 1024 * 1024 && // Max 100MB
                    pos + 8 + recLen <= buffer.Length)
                {
                    // Check for BLIP types (0xF018-0xF117)
                    if (recType >= 0xF018 && recType <= 0xF117)
                    {
                        var image = ExtractBlipFromOfficeArtRecord(buffer, pos + 8, (int)recLen, recType);
                        if (image != null)
                        {
                            images.Add(image);
                            // Mark range as extracted to avoid duplicate from raw signature scan
                            skipRanges?.Add((pos, Math.Min(pos + 8 + (int)recLen, buffer.Length)));
                        }
                    }

                    pos += 8 + (int)recLen;
                }
                else
                {
                    pos++;
                }
            }
            catch
            {
                pos++;
            }
        }
    }

    /// <summary>
    /// Extracts an image from a specific position in the buffer.
    /// </summary>
    private ImageModel? ExtractImageFromPosition(byte[] buffer, int pos, ImageType type)
    {
        try
        {
            int length = EstimateImageLength(buffer, pos, type);
            if (length <= 0 || pos + length > buffer.Length) return null;

            var data = new byte[length];
            Array.Copy(buffer, pos, data, 0, length);

            // Get dimensions if possible
            var (width, height) = GetImageDimensions(data, type);

            return new ImageModel
            {
                Id = $"image{pos}",
                FileName = $"image{pos}.{GetImageExtension(type)}",
                ContentType = GetContentType(type),
                Data = data,
                Type = type,
                Width = width,
                Height = height,
                WidthEMU = width > 0 ? width * 914400 / 96 : 0,
                HeightEMU = height > 0 ? height * 914400 / 96 : 0,
                PictureOffset = pos
            };
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Extracts a BLIP from an OfficeArt record.
    /// </summary>
    private ImageModel? ExtractBlipFromOfficeArtRecord(byte[] buffer, int offset, int length, ushort recType)
    {
        try
        {
            // BLIP record contains a header followed by the image data
            // Skip BLIP header (typically 16-34 bytes depending on type)
            int headerSize = GetBlipHeaderSize(recType);
            if (offset + headerSize >= buffer.Length) return null;

            var imageType = GetBlipImageType(recType);
            if (imageType == ImageType.Unknown) return null;

            var dataLength = length - headerSize;
            if (dataLength <= 0) return null;

            var data = new byte[dataLength];
            Array.Copy(buffer, offset + headerSize, data, 0, dataLength);

            var (width, height) = GetImageDimensions(data, imageType);

            return new ImageModel
            {
                Id = $"blip{offset}",
                FileName = $"image{offset}.{GetImageExtension(imageType)}",
                ContentType = GetContentType(imageType),
                Data = data,
                Type = imageType,
                Width = width,
                Height = height,
                WidthEMU = width > 0 ? width * 914400 / 96 : 0,
                HeightEMU = height > 0 ? height * 914400 / 96 : 0,
                PictureOffset = offset
            };
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Estimates the length of an image based on its type and structure.
    /// </summary>
    private int EstimateImageLength(byte[] buffer, int pos, ImageType type)
    {
        try
        {
            switch (type)
            {
                case ImageType.Png:
                    // PNG has chunks; look for IEND chunk
                    return FindPngEnd(buffer, pos);

                case ImageType.Jpeg:
                    // JPEG ends with EOI marker FF D9
                    return FindJpegEnd(buffer, pos);

                case ImageType.Gif:
                    // GIF ends with 3B
                    return FindGifEnd(buffer, pos);

                case ImageType.Dib:
                    // BMP has size in header
                    return FindBmpEnd(buffer, pos);

                default:
                    return 0;
            }
        }
        catch
        {
            return 0;
        }
    }

    private int FindPngEnd(byte[] buffer, int start)
    {
        int pos = start + 8; // Skip signature
        while (pos < buffer.Length - 12)
        {
            // PNG chunk length is 4 bytes big-endian (network byte order)
            var chunkLenU = (uint)((buffer[pos] << 24) | (buffer[pos + 1] << 16) | (buffer[pos + 2] << 8) | buffer[pos + 3]);
            if (chunkLenU > 100 * 1024 * 1024) break;
            var chunkLen = (int)chunkLenU;

            var chunkType = Encoding.ASCII.GetString(buffer, pos + 4, 4);
            if (chunkType == "IEND")
                return pos + 12; // Length (4) + Type (4) + Data (0) + CRC (4)

            pos += 12 + chunkLen;
        }
        return buffer.Length - start;
    }

    private int FindJpegEnd(byte[] buffer, int start)
    {
        int pos = start + 2; // Skip SOI
        while (pos < buffer.Length - 1)
        {
            if (buffer[pos] != 0xFF)
            {
                pos++;
                continue;
            }
            var m = buffer[pos + 1];
            if (m == 0xD9) // EOI
                return pos + 2 - start;
            if (m == 0xD8) // Another SOI (invalid)
                break;
            // Skip segments with length: SOF, DHT, DQT, DRI, APP, COM, SOS (has length then scan for next marker)
            if (m == 0xC0 || m == 0xC1 || m == 0xC2 || m == 0xC3 || m == 0xC4 || m == 0xC5 ||
                m == 0xC6 || m == 0xC7 || m == 0xC8 || m == 0xC9 || m == 0xCA || m == 0xCB ||
                m == 0xCC || m == 0xCD || m == 0xCE || m == 0xCF ||
                m == 0xDB || m == 0xDD || m == 0xE0 || m == 0xE1 || m == 0xE2 || m == 0xE3 ||
                m == 0xE4 || m == 0xE5 || m == 0xE6 || m == 0xE7 || m == 0xE8 || m == 0xE9 ||
                m == 0xEA || m == 0xEB || m == 0xEC || m == 0xED || m == 0xEE || m == 0xEF ||
                m == 0xFE || m == 0xDA) // SOS: skip length then scan until next 0xFF
            {
                if (pos + 4 > buffer.Length) break;
                var segLen = (buffer[pos + 2] << 8) | buffer[pos + 3];
                if (m == 0xDA) // SOS: after 2+segLen, scan for 0xFF 0xD9
                {
                    int sosEnd = pos + 2 + segLen;
                    for (int i = sosEnd; i < buffer.Length - 1; i++)
                    {
                        if (buffer[i] == 0xFF && buffer[i + 1] == 0xD9)
                            return i + 2 - start;
                        if (buffer[i] == 0xFF && buffer[i + 1] != 0x00) // Skip escaped 0xFF in scan
                            i++;
                    }
                    break;
                }
                pos += 2 + segLen;
                continue;
            }
            if (m >= 0xD0 && m <= 0xD7) { pos += 2; continue; } // RST: no length
            pos++;
        }
        return buffer.Length - start;
    }

    private int FindGifEnd(byte[] buffer, int start)
    {
        int pos = start;
        while (pos < buffer.Length)
        {
            if (buffer[pos] == 0x3B) // Trailer
                return pos + 1 - start;
            pos++;
        }
        return buffer.Length - start;
    }

    private int FindBmpEnd(byte[] buffer, int start)
    {
        if (start + 14 > buffer.Length) return 0;
        var size = BitConverter.ToInt32(buffer, start + 2);
        return size > 0 && size < buffer.Length - start ? size : 0;
    }

    /// <summary>
    /// Gets the image dimensions from the image data.
    /// </summary>
    private (int width, int height) GetImageDimensions(byte[] data, ImageType type)
    {
        try
        {
            switch (type)
            {
                case ImageType.Png:
                    if (data.Length >= 24)
                    {
                        var width = (data[16] << 24) | (data[17] << 16) | (data[18] << 8) | data[19];
                        var height = (data[20] << 24) | (data[21] << 16) | (data[22] << 8) | data[23];
                        return (width, height);
                    }
                    break;

                case ImageType.Jpeg:
                    return GetJpegDimensions(data);

                case ImageType.Gif:
                    if (data.Length >= 10)
                    {
                        var width = BitConverter.ToUInt16(data, 6);
                        var height = BitConverter.ToUInt16(data, 8);
                        return (width, height);
                    }
                    break;

                case ImageType.Dib:
                    if (data.Length >= 26)
                    {
                        var width = BitConverter.ToInt32(data, 18);
                        var height = BitConverter.ToInt32(data, 22);
                        return (width, height);
                    }
                    break;
            }
        }
        catch (Exception ex)
        {
            Logger.Warning("Failed to read image dimensions from embedded data.", ex);
        }

        return (0, 0);
    }

    private (int width, int height) GetJpegDimensions(byte[] data)
    {
        int pos = 2; // Skip SOI
        while (pos < data.Length - 9)
        {
            if (data[pos] == 0xFF && data[pos + 1] != 0x00 && data[pos + 1] != 0xFF)
            {
                var marker = data[pos + 1];
                if (marker == 0xD9) break; // EOI
                if (marker == 0xC0 || marker == 0xC1 || marker == 0xC2) // SOF markers
                {
                    var height = (data[pos + 5] << 8) | data[pos + 6];
                    var width = (data[pos + 7] << 8) | data[pos + 8];
                    return (width, height);
                }

                var len = (data[pos + 2] << 8) | data[pos + 3];
                pos += 2 + len;
            }
            else
            {
                pos++;
            }
        }
        return (0, 0);
    }

    private int GetBlipHeaderSize(ushort recType)
    {
        // MS-ODRAW 2.2.23+: BLIP header sizes
        return recType switch
        {
            0xF01A => 16, // EMF
            0xF01B => 16, // WMF
            0xF01C => 16, // PICT
            0xF01D => 17, // JPEG
            0xF01E => 17, // PNG
            0xF01F => 17, // DIB
            0xF020 => 17, // TIFF (legacy)
            0xF029 => 17, // TIFF
            0xF02A => 17, // JPEG (alternate)
            _ => 16
        };
    }

    private ImageType GetBlipImageType(ushort recType)
    {
        return recType switch
        {
            0xF01A => ImageType.Emf,
            0xF01B => ImageType.Wmf,
            0xF01D => ImageType.Jpeg,
            0xF01E => ImageType.Png,
            0xF01F => ImageType.Dib,
            0xF020 => ImageType.Tiff,
            0xF029 => ImageType.Tiff,
            0xF02A => ImageType.Jpeg,
            _ => ImageType.Unknown
        };
    }

    private string GetImageExtension(ImageType type)
    {
        return type switch
        {
            ImageType.Png => "png",
            ImageType.Jpeg => "jpg",
            ImageType.Gif => "gif",
            ImageType.Emf => "emf",
            ImageType.Wmf => "wmf",
            ImageType.Dib => "bmp",
            ImageType.Tiff => "tiff",
            _ => "bin"
        };
    }

    private string GetContentType(ImageType type)
    {
        return type switch
        {
            ImageType.Png => "image/png",
            ImageType.Jpeg => "image/jpeg",
            ImageType.Gif => "image/gif",
            ImageType.Emf => "image/x-emf",
            ImageType.Wmf => "image/x-wmf",
            ImageType.Dib => "image/bmp",
            ImageType.Tiff => "image/tiff",
            _ => "application/octet-stream"
        };
    }

    /// <summary>
    /// Determines image type from header bytes.
    /// </summary>
    public static ImageType GetImageType(byte[] header)
    {
        if (header.Length < 4) return ImageType.Unknown;

        // PNG: 89 50 4E 47
        if (header[0] == 0x89 && header[1] == 0x50 && header[2] == 0x4E && header[3] == 0x47)
            return ImageType.Png;

        // JPEG: FF D8 FF
        if (header[0] == 0xFF && header[1] == 0xD8 && header[2] == 0xFF)
            return ImageType.Jpeg;

        // GIF: 47 49 46
        if (header[0] == 0x47 && header[1] == 0x49 && header[2] == 0x46)
            return ImageType.Gif;

        // EMF: 01 00 00 00
        if (header[0] == 0x01 && header[1] == 0x00 && header[2] == 0x00 && header[3] == 0x00)
            return ImageType.Emf;

        // WMF: D7 CD C6 9A
        if (header[0] == 0xD7 && header[1] == 0xCD && header[2] == 0xC6 && header[3] == 0x9A)
            return ImageType.Wmf;

        return ImageType.Unknown;
    }
}
