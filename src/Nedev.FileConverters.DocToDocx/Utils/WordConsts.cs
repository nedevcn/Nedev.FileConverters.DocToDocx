namespace Nedev.FileConverters.DocToDocx.Utils;

/// <summary>
/// Constants for Microsoft Word Document (.doc) binary format
/// Based on MS-DOC specification
/// </summary>
public static class WordConsts
{
    // FIB (File Information Block) base offsets
    public const int FIB_BASE_SIZE = 24;
    public const int FIB_MAGIC_NUMBER = 0xA5EC;
    public const int FIB_MAGIC_NUMBER_OLD = 0xA5DB;

    // FIB offsets (for older versions)
    public const int FIB_O_WSB = 0x00;     // WSB (word start)
    public const int FIB_O_WSF = 0x01;     // WSF (word start flag)
    public const int FIB_O_MOS = 0x02;     // MOS (major version)
    public const int FIB_O_MLOS = 0x03;    // MLOS (minor version)
    public const int FIB_O_PNFB = 0x04;     // PNFB (first FFK)
    public const int FIB_O_CSW = 0x18;      // CSW (count of words)
    public const int FIB_O_RGW = 0x1A;     // RGW array starts

    // FKP (Formatted Disk Page) constants
    public const int FKP_PAGE_SIZE = 512;
    public const int FKP_CHP = 0x00;
    public const int FKP_PAP = 0x01;
    public const int FKP_FCPA = 0x02;
    public const int FKP_LFO = 0x03;

    // Piece Table constants
    public const int PLC_SIGNATURE = 0xFFFE;
    public const int PLC_MAX_POSITION = 0x3FFFFFFF;

    // Document properties
    public const int DOP_SIGNATURE = 0x078E;
    public const int DOP_UNKNOWN = 0x0C1A;

    // STSH (Style Sheet) constants
    public const int STSHI_STD = 0xFFFE;
    public const int STSHI_STD_NORMAL = 0xFFFD;

    // BTE (Begin Table Entry) markers
    public const byte BTE_BEGIN = 0x01;
    public const byte BTE_END = 0x02;
    public const byte BTE_ROW = 0x06;
    public const byte BTE_CELL = 0x07;
    public const byte BTE_ROW_END = 0x08;
    public const byte BTE_CELL_END = 0x09;

    // FIB flags
    public const int FIB_FLAG_FCOMPLICATED = 0x0001;
    public const int FIB_FLAG_FSHIFTED = 0x0010;
    public const int FIB_FLAG_FSPEECH = 0x0040;
    public const int FIB_FLAG_FCOMPLEX = 0x0100;
    public const int FIB_FLAG_FHIGHLIGHT = 0x0200;

    // CP (Character Position) special values
    public const int CP_DOCUMENT_START = 0;
    public const int CP_DOCUMENT_END = 0x3FFFFFFF;

    // Text encoding identifiers
    public const int ENCODING_ANSI = 0;
    public const int ENCODING_MAC = 1;
    public const int ENCODING_PC = 2;
    public const int ENCODING_PCA = 77;

    // Font families
    public const int FF_ROMAN = 0;
    public const int FF_SWISS = 1;
    public const int FF_MODERN = 2;
    public const int FF_SCRIPT = 3;
    public const int FF_DECORATIVE = 4;

    // Paragraph alignment
    public const int JUSTIFY_LEFT = 0;
    public const int JUSTIFY_CENTER = 1;
    public const int JUSTIFY_RIGHT = 2;
    public const int JUSTIFY_JUSTIFY = 3;

    // Sprm (Single Property Modifier) operation codes
    public const ushort SPRM_PJCN = 0x2400;     // Justification
    public const ushort SPRM_PDHIA = 0x2401;   // Space before
    public const ushort SPRM_PDPIA = 0x2402;   // Space after
    public const ushort SPRM_PDLINE = 0x2406;  // Line spacing
    public const ushort SPRM_PCHTO = 0x2600;   // First line indent
    public const ushort SPRM_PCHTO2 = 0x2601;  // Left indent
    public const ushort SPRM_PCHTO3 = 0x2602;  // Right indent

    // CHP (Character Properties) Sprms
    public const ushort SPRM_CRS = 0x2C00;     // Font size
    public const ushort SPRM_CBRC = 0x2C03;    // Bold
    public const ushort SPRM_CITC = 0x2C04;    // Italic
    public const ushort SPRM_CULS = 0x2C05;    // Underline
    public const ushort SPRM_CSS = 0x2C06;     // Strike through
    public const ushort SPRM_CKCS = 0x2C09;    // Superscript/Subscript
    public const ushort SPRM_CFF = 0x2C0A;     // Font family
    public const ushort SPRM_CFTC = 0x2C0B;    // Font type
    public const ushort SPRM_CHS = 0x2C0C;     // Character set
    public const ushort SPRM_CCC = 0x2C0F;     // Color

    // Table cell borders (TAP)
    public const ushort SPRM_TCWBORDER = 0x5600;
    public const ushort SPRM_TCLBORDER = 0x5601;
    public const ushort SPRM_TCRBORDER = 0x5602;
    public const ushort SPRM_TCTOPBORDERW = 0x5603;
    public const ushort SPRM_TCLEFTBORDERW = 0x5604;
    public const ushort SPRM_TCBOTTOMBORDERW = 0x5605;
    public const ushort SPRM_TCRIGHTBORDERW = 0x5606;

    // Text types
    public const int TEXT_TYPE_MAIN = 0;
    public const int TEXT_TYPE_HEADER = 1;
    public const int TEXT_TYPE_FOOTER = 2;
    public const int TEXT_TYPE_FOOTNOTE = 3;
    public const int TEXT_TYPE_ANNOTATION = 4;

    // Picture types
    public const int PICTYPE_WMETAFILE = 1;
    public const int PICTYPE_PMETAFILE = 2;
    public const int PICTYPE_MACPICT = 3;
    public const int PICTYPE_PNG = 4;
    public const int PICTYPE_JPEG = 5;
    public const int PICTYPE_DIB = 6;

    // Maximum values
    public const int MAX_STYLES = 0x03FF;
    public const int MAX_FONTS = 0x03FF;
    public const int MAX_DOCUMENT_SIZE = 0x7FFFFFFF;
}

/// <summary>
/// FIB (File Information Block) structure offsets for Word 97-2003
/// </summary>
public static class FibOffsets
{
    // Base FIB (24 bytes)
    public const int WSB = 0;
    public const int WSF = 1;
    public const int MOS = 2;
    public const int MLOS = 3;
    public const int PNFB = 4;
    public const int PNCHP = 6;
    public const int PNPAP = 8;
    public const int PNBP = 10;
    public const int PNPLC = 12;
    public const int PNTXT = 14;
    public const int PNDTTM = 16;
    public const int CLW = 18;
    public const int CBS = 20;
    public const int CWDB = 22;

    // Extended FIB (starts at 24)
    public const int CSW = 24;
    public const int FIB_KEY = 26;
    public const int FIB_FLAGS = 28;
    public const int HISTORY = 30;
    public const int CSEC = 32;
    public const int PIDX = 34;

    // Document properties
    public const int DOP = 36;
    public const int STSH = 40;
    public const int FIB_MAX = 44;

    // Character positions for different text types
    public const int CCPTXT = 44;
    public const int CCPFPAP = 48;
    public const int CCPFHEADER = 52;
    public const int CCPFFOOTER = 56;
    public const int CCPFFOOTNOTE = 60;
    public const int CCPFANNOTATION = 64;
    public const int CCPFBOOK = 68;
    public const int CCPFTOX = 72;
    public const int CCPFHDR = 76;

    // Style sheet info
    public const int STSHI_STD = 80;
    public const int STSHI_STD_NORMAL = 84;

    // Piece table
    public const int CLCP = 88;
    public const int CBCLP = 92;
    public const int PNParaSep = 96;
    public const int PnChpFirst = 100;
    public const int PnPapFirst = 102;

    // Complex FIB
    public const int BKCINFIB = 1888;
    public const int FLAS = 1892;
}

/// <summary>
/// Standard style identifiers
/// </summary>
public static class StyleIds
{
    public const ushort NORMAL = 1;
    public const ushort HEADING_1 = 2;
    public const ushort HEADING_2 = 3;
    public const ushort HEADING_3 = 4;
    public const ushort HEADING_4 = 5;
    public const ushort HEADING_5 = 6;
    public const ushort HEADING_6 = 7;
    public const ushort HEADING_7 = 8;
    public const ushort HEADING_8 = 9;
    public const ushort HEADING_9 = 10;
    public const ushort TITLE = 15;
    public const ushort BODY_TEXT = 2; // Often 2, but varies
}
