# Nedev.DocToDocx

A high‑fidelity `.doc` → `.docx` converter for .NET 10 with no third‑party dependencies.

## Features

- **Binary `.doc` reader**: Implements core MS‑DOC structures (CFB, FIB, CLX/Piece Table, CHPX/PAPX FKPs, PLCFs).
- **Rich text & styles**: Fonts, font sizes, bold/italic/underline, colors, highlighting, outline/emboss/shadow, language (`w:lang`), paragraph alignment, spacing, indentation, borders, shading, and numbered/bulleted lists (including many localized formats).
- **Tables**: Multi‑row/column tables with TAP‑driven layout: row height and exact/at‑least rules, header rows, `cantSplit`, per‑cell width, proper vertical merges (`vMerge restart/continue`), horizontal merges (`gridSpan`), table‑level borders and shading mapped to DOCX. Nested tables are now tracked as child `TableModel` instances with parent indexes, allowing the writer to emit them in place. Recent iterations improved usage of TAP metadata (preferred table width, left indent, cell spacing) for more faithful layout under complex merges.
- **Sections & page setup**: Multiple sections with page size/orientation, margins, starting page number, and First/Odd/Even headers/footers mapped to separate DOCX parts.
- **Images**: Extracts embedded images from the `WordDocument` and `Data` streams (PNG/JPEG/GIF/BMP/OfficeArt BLIPs), writes `word/media/*`, generates `w:drawing` with size inferred from image dimensions and auto‑scaled to page width, respects per‑image scale, and attaches basic alt text.
- **OfficeArt pictures & floating anchors**: Parses Escher/OfficeArt records and FSPA anchors from `PlcSpaMom` to recover picture shapes; maps them to `wp:anchor` floating images positioned relative to the page, falling back to inline images when anchors are unavailable.
- **Basic charts (experimental)**: When `.doc` files contain embedded OLE chart-like streams, the converter can emit minimal, editable DOCX chart parts (`word/charts/chartN.xml`) with placeholder data, so charts remain editable in Word even if the original series data is not yet fully understood.
- **Footnotes, endnotes, comments, textboxes**: Reads and writes common note and annotation structures into DOCX footnotes/endnotes parts and DrawingML textboxes.
- **Encryption**:
  - **XOR**: Supports Word’s XOR‑obfuscated streams via `EncryptionHelper` and decrypted CFB streams.
  - **RC4 (Passworded)**: Handles Office 97‑2003 RC4‑encrypted documents. The converter will prompt for a password and verify it against the verifier/hash; incorrect passwords are rejected. Full RC4 decryption is performed on the `WordDocument`, `Table`, and `Data` streams.
- **No external dependencies**: Pure .NET, streaming writers (`XmlWriter`) for high performance and low memory usage.

> Note: While many MS‑DOC features are implemented (including OfficeArt‑based picture extraction and floating anchors), the converter does not yet claim 100% coverage of the full [MS‑DOC] / [MS‑ODRAW] specifications. Complex vector shapes, SmartArt, charts, OLE objects, and some rare formatting cases remain intentionally out of scope for now.

## Enhanced compatibility & current limitations

- **Enhanced table compatibility**:
  - TAP‑level information (table width, left indent, cell spacing, borders, shading, header rows, `cantSplit`) is now decoded and preserved through the `TableModel` so that DOCX output can more closely match the original `.doc` layout, especially for merged cells.
  - Vertical merges (`vMerge restart/continue`) and horizontal merges (`gridSpan`) are inferred using a combination of TAP merge flags and content heuristics, which significantly improves the appearance of common merged‑cell tables.
- **Known limitations (deliberate)**:
  - Deeply nested tables and extremely exotic merge patterns may still be flattened or approximated; the goal is a robust, readable DOCX rather than a byte‑perfect structural clone of the original MS‑DOC.
  - Complex OfficeArt vector shapes, SmartArt, rich chart types (with full Excel-backed data), OLE objects, and other advanced drawing features continue to be out of scope; when encountered they are either ignored or downgraded to simpler picture or placeholder chart representations where possible.

## Usage

### Library

Add a reference to the `Nedev.DocToDocx` assembly and call the static converter API:

```csharp
using Nedev.DocToDocx;

DocToDocxConverter.Convert("input.doc", "output.docx");
```

### Command Line Interface (CLI)

You can also use the included CLI tool to convert documents directly from the command line:

```bash
Nedev.DocToDocx.Cli <input.doc|input.docx|inputDir> <output.docx|outputDir> [-p <password>] [-r]
```

When given a `.doc` file the tool will perform a conversion; a `.docx` input is simply copied to the output location (useful for batch scripts). The CLI also understands password-protected `.doc` files (`-p`), verifying the password before attempting to read.

**Arguments:**
- `<input.doc>` The path to the input MS-DOC file.
- `<output.docx>` The path where the output DOCX file will be saved.

**Options:**
- `-p`, `--password` The password to open an encrypted DOC file.
- `-r`, `--recursive` When a directory is supplied as `<input>`, process `.doc` files recursively and mirror the structure in `<output>`.
- `-h`, `--help` Show this help message and exit.


### Running the tests

The repository now ships with a set of unit and integration tests under `src/Nedev.DocToDocx.Tests`.
Sample `.doc`/`.docx` files are included and copied to the test output directory so that
integration tests can exercise the reader and CLI tool. To run the tests execute:

```bash
cd src/Nedev.DocToDocx.Tests
dotnet test
```

The new tests verify document writer output, reader loading of a sample document and
that the CLI can successfully convert it. A GitHub Actions workflow (`.github/workflows/ci.yml`)
is included to build the solution and run the tests on each push or pull request.
