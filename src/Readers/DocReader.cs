using System.Text;
using Nedev.DocToDocx.Models;
using Nedev.DocToDocx.Utils;

namespace Nedev.DocToDocx.Readers;

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
    private CfbReader? _cfb;
    private BinaryReader? _wordDocReader;
    private BinaryReader? _tableReader;
    private BinaryReader? _dataReader;

    private FibReader? _fibReader;
    private TextReader? _textReader;
    private StyleReader? _styleReader;
    private DocumentPropertiesReader? _dopReader;
    private TableReader? _tableParseReader;
    private ImageReader? _imageReader;
    private FkpParser? _fkpParser;
    private FootnoteReader? _footnoteReader;
    private AnnotationReader? _annotationReader;
    private TextboxReader? _textboxReader;
    private HeaderFooterReader? _headerFooterReader;
    private ListReader? _listReader;
    private FieldReader? _fieldReader;
    private HyperlinkReader? _hyperlinkReader;
    private OfficeArtReader? _officeArtReader;
    private List<FspaInfo> _fspaAnchors = new();

    // Keep streams alive for reader lifetime
    private MemoryStream? _wordDocStream;
    private MemoryStream? _tableStream;
    private MemoryStream? _dataStream;
    private MemoryStream? _footnoteStream;
    private MemoryStream? _endnoteStream;
    private MemoryStream? _anotStream;
    private MemoryStream? _txbxStream;

    public DocumentModel Document { get; private set; } = new();
    public bool IsLoaded { get; private set; }

    public DocReader(string filePath)
    {
        _cfb = new CfbReader(filePath);
        InitializeStreams();
    }

    public DocReader(Stream stream)
    {
        _cfb = new CfbReader(stream, leaveOpen: true);
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

        // If the document is XOR-encrypted, configure the CFB reader with the key
        // and reopen the core streams in decrypted form for all subsequent parsing.
        if (_fibReader.FEncrypted)
        {
            _cfb.SetEncryptionKey(_fibReader.LKey);

            _wordDocReader.Dispose();
            _wordDocStream.Dispose();

            _wordDocStream = _cfb.GetDecryptedStream("WordDocument");
            _wordDocReader = new BinaryReader(_wordDocStream, Encoding.Default, leaveOpen: true);
        }

        // Extract Table stream (0Table or 1Table)
        var tableName = _fibReader.TableStreamName;
        if (!_cfb.HasStream(tableName))
        {
            // Try the other one
            tableName = tableName == "1Table" ? "0Table" : "1Table";
            if (!_cfb.HasStream(tableName))
                throw new InvalidDataException($"Table stream not found. Tried '{_fibReader.TableStreamName}' and '{tableName}'.");
        }

        _tableStream = _fibReader.FEncrypted
            ? _cfb.GetDecryptedStream(tableName)
            : _cfb.GetStream(tableName);
        _tableReader = new BinaryReader(_tableStream, Encoding.Default, leaveOpen: true);

        // Read floating shape anchors from PlcfSpaMom (best-effort).
        try
        {
            _fspaAnchors = FspaReader.ReadPlcSpaMom(_tableReader, _fibReader);
        }
        catch
        {
            _fspaAnchors = new List<FspaInfo>();
        }

        // Extract Data stream (optional — contains pictures, OLE objects)
        if (_cfb.HasStream("Data"))
        {
            _dataStream = _fibReader.FEncrypted
                ? _cfb.GetDecryptedStream("Data")
                : _cfb.GetStream("Data");
            _dataReader = new BinaryReader(_dataStream, Encoding.Default, leaveOpen: true);

            // Initialize OfficeArt/Escher reader on the Data stream (best-effort).
            try
            {
                _officeArtReader = new OfficeArtReader(_dataStream);
            }
            catch
            {
                _officeArtReader = null;
            }
        }

        // Extract footnote/endnote streams (optional)
        if (_cfb.HasStream("Footnote"))
        {
            _footnoteStream = _fibReader.FEncrypted
                ? _cfb.GetDecryptedStream("Footnote")
                : _cfb.GetStream("Footnote");
        }

        if (_cfb.HasStream("Endnote"))
        {
            _endnoteStream = _fibReader.FEncrypted
                ? _cfb.GetDecryptedStream("Endnote")
                : _cfb.GetStream("Endnote");
        }

        // Extract annotation stream (optional)
        if (_cfb.HasStream("Anot"))
        {
            _anotStream = _fibReader.FEncrypted
                ? _cfb.GetDecryptedStream("Anot")
                : _cfb.GetStream("Anot");
        }

        // Extract textbox stream (optional)
        if (_cfb.HasStream("Txbx"))
        {
            _txbxStream = _fibReader.FEncrypted
                ? _cfb.GetDecryptedStream("Txbx")
                : _cfb.GetStream("Txbx");
        }

        // Initialize sub-readers
        _textReader = new TextReader(_wordDocReader!, _tableReader!, _fibReader!);
        _styleReader = new StyleReader(_tableReader!, _fibReader!);
        _dopReader = new DocumentPropertiesReader(_tableReader!, _fibReader!);
        _fkpParser = new FkpParser(_wordDocReader!, _tableReader!, _fibReader!, _textReader!);
        _tableParseReader = new TableReader(_wordDocReader!, _tableReader!, _fibReader!, _fkpParser);
        _imageReader = new ImageReader(_wordDocReader!, _dataReader, _fibReader!);
        _footnoteReader = new FootnoteReader(_fibReader!, _textReader!);
        _annotationReader = new AnnotationReader(
            _anotStream != null ? new BinaryReader(_anotStream, Encoding.Default, leaveOpen: true) : null,
            _fibReader!);
        _textboxReader = new TextboxReader(_tableReader!, _fibReader!, _textReader!);
        _headerFooterReader = new HeaderFooterReader(_tableReader!, _wordDocReader!, _fibReader!, _textReader!);
        _listReader = new ListReader(_tableReader!, _fibReader!);
        _fieldReader = new FieldReader();
        _hyperlinkReader = new HyperlinkReader();
    }

    /// <summary>
    /// Loads and parses the document.
    /// </summary>
    public void Load()
    {
        // Step 1: Read document properties
        Document.Properties = _dopReader!.Read();

        // Step 1.5: Read style sheet
        _styleReader!.Read();
        Document.Styles = _styleReader.Styles;

        // Step 2: Read list definitions
        _listReader!.Read();
        Document.NumberingDefinitions = _listReader.NumberingDefinitions;
        Document.ListFormats = _listReader.ListFormats;

        // Step 3: Read text content via Piece Table
        _textReader!.ReadText();

        // Step 4: Parse paragraphs and runs
        ParseDocumentContent();

        // Step 5: Parse tables
        _tableParseReader!.ParseTables(Document);

        // Step 6: Extract images
        _imageReader!.ExtractImages(Document);

        // Step 6.5: Parse OfficeArt/Escher shapes and map basic anchors
        if (_officeArtReader != null)
        {
            OfficeArtMapper.AttachShapes(Document, _officeArtReader, _fspaAnchors);
        }

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
        }

        // Step 10: Read headers/footers
        if (_headerFooterReader != null)
        {
            _headerFooterReader.Read(Document);
            Document.HeadersFooters.Headers = _headerFooterReader.Headers;
            Document.HeadersFooters.Footers = _headerFooterReader.Footers;
        }

        IsLoaded = true;
    }

    /// <summary>
    /// Parses the document content into paragraphs and runs using FKP-based parsing.
    /// 
    /// This implementation:
    ///   - Reads CHP and PAP properties from FKPs
    ///   - Splits text by paragraph marks (CR = 0x0D)
    ///   - Creates multiple runs per paragraph based on CHP changes
    ///   - Applies paragraph formatting from PAP FKPs
    /// </summary>
    private void ParseDocumentContent()
    {
        var text = _textReader!.Text;
        if (string.IsNullOrEmpty(text)) return;

        // Read CHP and PAP properties from FKPs
        var chpMap = _fkpParser!.ReadChpProperties();
        var papMap = _fkpParser.ReadPapProperties();

        // In Word binary format, paragraphs are delimited by CR (0x0D)
        // Special characters: 0x07 = cell mark, 0x0C = page break,
        // 0x01 = field begin/end or inline picture
        var paragraphIndex = 0;
        var paraStart = 0;
        var imageCounter = 0;

        // Only iterate characters in the main document range [0, ccpText)
        int mainDocumentLength = Math.Min(text.Length, _fibReader!.CcpText);

        for (int i = 0; i <= mainDocumentLength; i++)
        {
            bool isParagraphEnd = (i == mainDocumentLength) || (text[i] == '\r') || (text[i] == '\x0D');

            if (!isParagraphEnd) continue;

            var paraText = text.Substring(paraStart, i - paraStart);
            var paraStartCp = paraStart;
            paraStart = i + 1; // skip the delimiter
            if (paraStart > mainDocumentLength) paraStart = mainDocumentLength;

            // Get PAP for this paragraph (use the first CP of the paragraph)
            PapBase? pap = null;
            if (paraStartCp < text.Length && papMap.TryGetValue(paraStartCp, out var foundPap))
            {
                pap = foundPap;
            }

            var paragraph = new ParagraphModel
            {
                Index = paragraphIndex++,
                Type = ParagraphType.Normal,
                Properties = pap != null 
                    ? _fkpParser.ConvertToParagraphProperties(pap, Document.Styles)
                    : new ParagraphProperties(),
                ListFormatId = pap?.ListFormatId ?? 0,
                ListLevel = pap?.ListLevel ?? 0
            };

            // Detect special paragraph types
            if (paraText.Contains('\x07'))
            {
                // Contains cell end mark — this is a table cell paragraph
                paragraph.Type = ParagraphType.TableCell;
            }
            else if (paraText.Contains('\x0C'))
            {
                paragraph.Type = ParagraphType.PageBreak;
            }

            // Split paragraph into runs based on CHP changes
            var runs = ParseRunsInParagraph(paraText, paraStartCp, chpMap, ref imageCounter);
            paragraph.Runs.AddRange(runs);

            // If no runs were created (no CHP data), create a default run
            if (paragraph.Runs.Count == 0)
            {
                var cleanText = CleanSpecialChars(paraText);
                if (!string.IsNullOrEmpty(cleanText))
                {
                    paragraph.Runs.Add(new RunModel
                    {
                        Text = cleanText,
                        CharacterPosition = paraStartCp,
                        CharacterLength = paraText.Length,
                        Properties = new RunProperties { FontSize = 24 }
                    });
                }
            }

            Document.Paragraphs.Add(paragraph);
        }
    }

    /// <summary>
    /// Parses runs within a paragraph based on CHP property changes.
    /// </summary>
    private List<RunModel> ParseRunsInParagraph(string paraText, int paraStartCp, Dictionary<int, ChpBase> chpMap, ref int imageCounter)
    {
        var runs = new List<RunModel>();
        if (string.IsNullOrEmpty(paraText)) return runs;

        var runStart = 0;
        ChpBase? currentChp = null;

        for (int i = 0; i <= paraText.Length; i++)
        {
            var cp = paraStartCp + i;
            ChpBase? chpAtCp = null;
            
            // Try to get CHP for this character position
            if (chpMap.TryGetValue(cp, out var foundChp))
            {
                chpAtCp = foundChp;
            }

            // Check if CHP changed or we're at the end
            bool chpChanged = i == paraText.Length || !ChpEquals(currentChp, chpAtCp);

            if (chpChanged && runStart < i)
            {
                // Create run for the segment
                var runText = paraText.Substring(runStart, i - runStart);
                var cleanText = CleanSpecialChars(runText);
                var isPicture = runText.Contains('\x01') || runText.Contains('\x08');

                if (!string.IsNullOrEmpty(cleanText) || isPicture)
                {
                    var run = new RunModel
                    {
                        Text = cleanText,
                        CharacterPosition = paraStartCp + runStart,
                        CharacterLength = runText.Length,
                        Properties = currentChp != null 
                            ? _fkpParser!.ConvertToRunProperties(currentChp, Document.Styles)
                            : new RunProperties { FontSize = 24 }
                    };

                    // Check for field characters (0x13 = field begin, 0x14 = separator, 0x15 = end)
                    if (runText.Contains('\x13'))
                    {
                        run.IsField = true;
                        var fieldStart = runText.IndexOf('\x13');
                        var fieldSep = runText.IndexOf('\x14');
                        if (fieldStart >= 0 && fieldSep > fieldStart)
                        {
                            run.FieldCode = runText.Substring(fieldStart + 1, fieldSep - fieldStart - 1).Trim();
                        }

                        // Try to interpret hyperlink fields as true OOXML hyperlinks
                        if (!string.IsNullOrEmpty(run.FieldCode) && _fieldReader != null && _hyperlinkReader != null)
                        {
                            var field = _fieldReader.ParseField(run.FieldCode);
                            if (field != null && field.Type == FieldType.Hyperlink)
                            {
                                var hyperlink = _hyperlinkReader.ParseHyperlink(run.FieldCode)
                                                ?? _hyperlinkReader.CreateHyperlink(field.Arguments);

                                if (!string.IsNullOrEmpty(hyperlink.Url))
                                {
                                    run.IsHyperlink = true;
                                    run.HyperlinkUrl = hyperlink.Url;
                                    run.HyperlinkRelationshipId = hyperlink.RelationshipId;

                                    // Treat this run as a hyperlink rather than a generic field
                                    run.IsField = false;

                                    if (!Document.Hyperlinks.Any(h => string.Equals(h.Url, hyperlink.Url, StringComparison.OrdinalIgnoreCase)
                                                                       && string.Equals(h.Bookmark, hyperlink.Bookmark, StringComparison.Ordinal)))
                                    {
                                        Document.Hyperlinks.Add(hyperlink);
                                    }
                                }
                            }
                        }
                    }

                    if (isPicture)
                    {
                        run.IsPicture = true;
                        run.ImageIndex = imageCounter++;
                    }

                    runs.Add(run);
                }

                runStart = i;
            }

            currentChp = chpAtCp;
        }

        return runs;
    }

    /// <summary>
    /// Compares two CHP objects for equality.
    /// </summary>
    private static bool ChpEquals(ChpBase? a, ChpBase? b)
    {
        if (ReferenceEquals(a, b)) return true;
        if (a == null || b == null) return false;

        return a.IsBold == b.IsBold &&
               a.IsItalic == b.IsItalic &&
               a.IsStrikeThrough == b.IsStrikeThrough &&
               a.IsUnderline == b.IsUnderline &&
               a.FontSize == b.FontSize &&
               a.FontIndex == b.FontIndex &&
               a.Color == b.Color;
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
    /// Table reader — parses table structures from document.
    /// Uses a combination of ParagraphType.TableCell markers and TAP (table
    /// properties) decoded from PAP/FKP data. This gives reasonably faithful
    /// reconstruction of row heights, header rows and horizontal/vertical merges
    /// for the common cases while deliberately avoiding the full generality of
    /// nested tables and exotic merge patterns present in the MS-DOC format.
/// </summary>
public class TableReader
{
    private readonly BinaryReader _wordDocReader;
    private readonly BinaryReader _tableReader;
    private readonly FibReader _fib;
    private readonly FkpParser _fkpParser;

        public TableReader(BinaryReader wordDocReader, BinaryReader tableReader, FibReader fib, FkpParser fkpParser)
        {
            _wordDocReader = wordDocReader;
            _tableReader = tableReader;
            _fib = fib;
            _fkpParser = fkpParser;
        }

    /// <summary>
    /// Parses tables from the document by examining paragraph types.
    /// Groups contiguous ParagraphType.TableCell 段落为一张张独立的表格。
    /// </summary>
        public void ParseTables(DocumentModel document)
        {
            var tables = new List<TableModel>();
            TableModel? currentTable = null;
            var rowIndex = 0;
            var cellsInCurrentRow = new List<TableCellModel>();
            int lastTableParagraphIndex = -1;
            TapBase? currentRowTap = null;
            // 保留每一行对应的 TAP 信息，用于之后根据 TC.grfw 计算单元格纵向/横向合并。
            var rowTaps = new List<TapBase?>();

        foreach (var para in document.Paragraphs.OrderBy(p => p.Index))
        {
            if (para.Type != ParagraphType.TableCell)
            {
                // 离开表格区域：收尾当前表格
                if (currentTable != null)
                {
                    FlushCurrentRow();
                    FinalizeCurrentTable();
                }
                continue;
            }

            // 进入一个新的表格块
            if (currentTable == null)
            {
                currentTable = new TableModel
                {
                    Index = tables.Count,
                    StartParagraphIndex = para.Index
                };
                rowIndex = 0;
                cellsInCurrentRow.Clear();
                currentRowTap = null;
            }

            lastTableParagraphIndex = para.Index;

            // 从 FKP/PAPX 中获取与该段落关联的 TAP 信息
            TapBase? tapForParagraph = null;
            var firstRun = para.Runs.FirstOrDefault();
            if (firstRun != null)
            {
                var pap = _fkpParser.GetPapAtCp(firstRun.CharacterPosition);
                tapForParagraph = pap?.Tap;

                if (tapForParagraph != null && currentTable.Properties == null)
                {
                    // Map TAP‑level table properties into the high‑level model so that
                    // the writer can faithfully reproduce alignment, indent, spacing
                    // table width, and table‑wide borders / shading.
                    currentTable.Properties = new TableProperties
                    {
                        Alignment = tapForParagraph.Justification switch
                        {
                            1 => TableAlignment.Center,
                            2 => TableAlignment.Right,
                            _ => TableAlignment.Left
                        },
                        // Prefer the TAP CellSpacing value; if it is zero, fall back to
                        // 2 * GapHalf (derived from sprmTDxaGapHalf) where available.
                        CellSpacing = tapForParagraph.CellSpacing != 0
                            ? tapForParagraph.CellSpacing
                            : (tapForParagraph.GapHalf != 0 ? tapForParagraph.GapHalf * 2 : 0),
                        Indent = tapForParagraph.IndentLeft,
                        PreferredWidth = tapForParagraph.TableWidth,
                        BorderTop = tapForParagraph.BorderTop,
                        BorderBottom = tapForParagraph.BorderBottom,
                        BorderLeft = tapForParagraph.BorderLeft,
                        BorderRight = tapForParagraph.BorderRight,
                        BorderInsideH = tapForParagraph.BorderInsideH,
                        BorderInsideV = tapForParagraph.BorderInsideV,
                        Shading = tapForParagraph.Shading
                    };
                }

                if (currentRowTap == null && tapForParagraph != null)
                {
                    currentRowTap = tapForParagraph;
                }
            }

            // 检测是否为“行结束”标记段落：只包含单元格结束符 \x07 或完全为空
            var isRowEnd = para.Runs.Count == 0 ||
                           (para.Runs.Count == 1 && string.IsNullOrWhiteSpace(para.Runs[0].Text.Replace("\x07", "")));

            if (isRowEnd)
            {
                // 结束当前行但不创建新的单元格
                FlushCurrentRow();
            }
            else
            {
                // 普通单元格内容段落
                var cellParagraph = new ParagraphModel
                {
                    Index = 0,
                    Type = ParagraphType.Normal,
                    Runs = para.Runs.Select(r => new RunModel
                    {
                        Text = r.Text.Replace("\x07", ""),
                        Properties = r.Properties != null ? new RunProperties
                        {
                            IsBold = r.Properties.IsBold,
                            IsItalic = r.Properties.IsItalic,
                            IsUnderline = r.Properties.IsUnderline,
                            FontSize = r.Properties.FontSize,
                            FontName = r.Properties.FontName,
                            Color = r.Properties.Color,
                            BgColor = r.Properties.BgColor
                        } : null
                    }).ToList()
                };

                // 移除空 run
                cellParagraph.Runs.RemoveAll(r => string.IsNullOrEmpty(r.Text));

                var cellModel = new TableCellModel
                {
                    Index = cellsInCurrentRow.Count,
                    RowIndex = rowIndex,
                    ColumnIndex = cellsInCurrentRow.Count,
                    Paragraphs = new List<ParagraphModel> { cellParagraph }
                };

                // 从 TAP 的 CellWidths 推导单元格宽度（使用第一行的宽度信息）
                if (tapForParagraph?.CellWidths != null &&
                    tapForParagraph.CellWidths.Length > cellModel.ColumnIndex)
                {
                    cellModel.Properties ??= new TableCellProperties();
                    cellModel.Properties.Width = tapForParagraph.CellWidths[cellModel.ColumnIndex];
                }

                cellsInCurrentRow.Add(cellModel);
            }
        }

        // 文档结束时，如仍在表格中，收尾
            if (currentTable != null)
            {
                FlushCurrentRow();
                FinalizeCurrentTable();
            }

            document.Tables = tables;

            // 本地帮助方法：结束当前行
            void FlushCurrentRow()
            {
                if (currentTable == null) return;
                if (cellsInCurrentRow.Count == 0) return;

                var row = new TableRowModel
                {
                    Index = rowIndex++,
                    Cells = new List<TableCellModel>(cellsInCurrentRow)
                };

                if (currentRowTap != null)
                {
                    row.Properties ??= new TableRowProperties();
                    if (currentRowTap.RowHeight > 0)
                    {
                        row.Properties.Height = currentRowTap.RowHeight;
                        row.Properties.HeightIsExact = currentRowTap.HeightIsExact;
                    }
                    if (currentRowTap.IsHeaderRow)
                    {
                        row.Properties.IsHeaderRow = true;
                    }
                    row.Properties.AllowBreakAcrossPages = !currentRowTap.CantSplit;
                }

                currentTable.Rows.Add(row);
                rowTaps.Add(currentRowTap);
                cellsInCurrentRow.Clear();
                currentRowTap = null;
            }

            // 本地帮助方法：计算行列数并添加到集合
            void FinalizeCurrentTable()
            {
                if (currentTable == null) return;
                if (currentTable.Rows.Count == 0)
                {
                    currentTable = null;
                    rowTaps.Clear();
                    return;
                }

                currentTable.EndParagraphIndex = lastTableParagraphIndex;
                currentTable.RowCount = currentTable.Rows.Count;
                currentTable.ColumnCount = currentTable.Rows.Max(r => r.Cells.Count);

                // 默认将每张表的第一行标记为表头行，便于在 Word 中重复显示在每页顶部。
                var firstRow = currentTable.Rows.FirstOrDefault();
                if (firstRow != null)
                {
                    firstRow.Properties ??= new TableRowProperties();
                    firstRow.Properties.IsHeaderRow = true;
                }

                // 1) 基于 TAP / TC 的精确信息推断纵向合并（RowSpan）
                bool hasTapMergeInfo = rowTaps.Any(t => t?.CellMerges != null);
                if (hasTapMergeInfo && currentTable.ColumnCount > 0)
                {
                    for (int col = 0; col < currentTable.ColumnCount; col++)
                    {
                        int row = 0;
                        while (row < currentTable.Rows.Count)
                        {
                            var startCell = GetCell(currentTable, row, col);
                            if (startCell == null)
                            {
                                row++;
                                continue;
                            }

                            var tap = row < rowTaps.Count ? rowTaps[row] : null;
                            var mergeArray = tap?.CellMerges;
                            CellMergeFlags? flags = null;
                            if (mergeArray != null && col < mergeArray.Length)
                            {
                                flags = mergeArray[col];
                            }

                            if (flags == null || !flags.VertFirst)
                            {
                                row++;
                                continue;
                            }

                            int span = 1;
                            int nextRow = row + 1;
                            while (nextRow < currentTable.Rows.Count)
                            {
                                var nextTap = nextRow < rowTaps.Count ? rowTaps[nextRow] : null;
                                var nextArray = nextTap?.CellMerges;
                                CellMergeFlags? nextFlags = null;
                                if (nextArray != null && col < nextArray.Length)
                                {
                                    nextFlags = nextArray[col];
                                }

                                if (nextFlags == null || !nextFlags.VertMerged)
                                {
                                    break;
                                }

                                span++;
                                nextRow++;
                            }

                            if (span > 1)
                            {
                                startCell.RowSpan = span;
                                row += span;
                            }
                            else
                            {
                                row++;
                            }
                        }
                    }
                }
                // 2) 回退：基于内容的启发式纵向合并（与之前版本保持兼容）
                else if (currentTable.ColumnCount > 0)
                {
                    for (int col = 0; col < currentTable.ColumnCount; col++)
                    {
                        int row = 0;
                        while (row < currentTable.Rows.Count)
                        {
                            var startCell = GetCell(currentTable, row, col);
                            if (startCell == null)
                            {
                                row++;
                                continue;
                            }

                            if (!CellHasContent(startCell))
                            {
                                row++;
                                continue;
                            }

                            int span = 1;
                            int nextRow = row + 1;
                            while (nextRow < currentTable.Rows.Count)
                            {
                                var nextCell = GetCell(currentTable, nextRow, col);
                                if (nextCell == null)
                                {
                                    break;
                                }

                                if (CellHasContent(nextCell))
                                {
                                    break;
                                }

                                span++;
                                nextRow++;
                            }

                            if (span > 1)
                            {
                                startCell.RowSpan = span;
                                row += span;
                            }
                            else
                            {
                                row++;
                            }
                        }
                    }
                }

                // 3) 基于 TAP / TC 的精确信息推断横向合并（ColumnSpan）
                if (hasTapMergeInfo && currentTable.ColumnCount > 0)
                {
                    for (int row = 0; row < currentTable.Rows.Count; row++)
                    {
                        var tap = row < rowTaps.Count ? rowTaps[row] : null;
                        var mergeArray = tap?.CellMerges;
                        if (mergeArray == null || mergeArray.Length == 0) continue;

                        int col = 0;
                        while (col < currentTable.ColumnCount)
                        {
                            var cell = GetCell(currentTable, row, col);
                            if (cell == null)
                            {
                                col++;
                                continue;
                            }

                            CellMergeFlags? flags = col < mergeArray.Length ? mergeArray[col] : null;
                            if (flags == null || !flags.HorizFirst)
                            {
                                col++;
                                continue;
                            }

                            int span = 1;
                            int nextCol = col + 1;
                            while (nextCol < currentTable.ColumnCount)
                            {
                                var nextFlags = nextCol < mergeArray.Length ? mergeArray[nextCol] : null;
                                if (nextFlags == null || !nextFlags.HorizMerged)
                                {
                                    break;
                                }
                                span++;
                                nextCol++;
                            }

                            if (span > 1)
                            {
                                cell.ColumnSpan = span;
                                col += span;
                            }
                            else
                            {
                                col++;
                            }
                        }
                    }
                }

                tables.Add(currentTable);
                currentTable = null;
                rowTaps.Clear();

                static TableCellModel? GetCell(TableModel table, int rowIndex, int columnIndex)
                {
                    if (rowIndex < 0 || rowIndex >= table.Rows.Count) return null;
                    var row = table.Rows[rowIndex];
                    if (columnIndex < 0 || columnIndex >= row.Cells.Count) return null;
                    return row.Cells[columnIndex];
                }

                static bool CellHasContent(TableCellModel cell)
                {
                    foreach (var para in cell.Paragraphs)
                    {
                        if (para.Runs.Any(r => !string.IsNullOrEmpty(r.Text)))
                        {
                            return true;
                        }
                    }
                    return false;
                }
            }
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

    public ImageReader(BinaryReader wordDocReader, BinaryReader? dataReader, FibReader fib)
    {
        _wordDocReader = wordDocReader;
        _dataReader = dataReader;
        _fib = fib;
    }

    /// <summary>
    /// Extracts images from the document.
    /// Parses OfficeArt records from the Data stream to extract embedded images.
    /// </summary>
    public void ExtractImages(DocumentModel document)
    {
        document.Images = new List<ImageModel>();

        if (_dataReader == null) return;

        try
        {
            // Scan Data stream for BLIP (Binary Large Image or Picture) records
            var images = ScanForBlipRecords();
            document.Images.AddRange(images);
        }
        catch (Exception ex)
        {
            // Image extraction is best-effort; don't fail the entire conversion
            Console.WriteLine($"Warning: Image extraction failed: {ex.Message}");
        }
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

        // Read entire Data stream into memory for scanning
        data.Seek(0, SeekOrigin.Begin);
        var buffer = new byte[length];
        _dataReader.Read(buffer, 0, (int)length);

        // Scan for common image signatures
        ScanForImageSignatures(buffer, images);

        // Also try to parse OfficeArt records
        ScanForOfficeArtRecords(buffer, images);

        return images;
    }

    /// <summary>
    /// Scans for raw image signatures (PNG, JPEG, GIF, etc.) in the data.
    /// </summary>
    private void ScanForImageSignatures(byte[] buffer, List<ImageModel> images)
    {
        int pos = 0;
        while (pos < buffer.Length - 8)
        {
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
                    pos += image.Data.Length;
                    continue;
                }
            }

            pos++;
        }
    }

    /// <summary>
    /// Scans for OfficeArt container records.
    /// </summary>
    private void ScanForOfficeArtRecords(byte[] buffer, List<ImageModel> images)
    {
        // OfficeArt record header format:
        // - recVer (4 bits) + recInstance (12 bits) = 2 bytes
        // - recType (2 bytes)
        // - recLen (4 bytes)

        int pos = 0;
        while (pos < buffer.Length - 8)
        {
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
                HeightEMU = height > 0 ? height * 914400 / 96 : 0
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
                HeightEMU = height > 0 ? height * 914400 / 96 : 0
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
            var chunkLen = (int)BitConverter.ToUInt32(buffer, pos);
            if (chunkLen > 100 * 1024 * 1024) break; // Sanity check

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
            if (buffer[pos] == 0xFF)
            {
                if (buffer[pos + 1] == 0xD9) // EOI marker
                    return pos + 2 - start;
                if (buffer[pos + 1] == 0xD8) // Another SOI (invalid)
                    break;
            }
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
                        var width = (int)BitConverter.ToUInt32(data, 16);
                        var height = (int)BitConverter.ToUInt32(data, 20);
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
        catch { }

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
        // BLIP header sizes vary by type
        return recType switch
        {
            0xF01A => 16, // EMF
            0xF01B => 16, // WMF
            0xF01C => 16, // PICT
            0xF01D => 17, // JPEG
            0xF01E => 17, // PNG
            0xF01F => 17, // BMP
            0xF020 => 17, // TIFF
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
