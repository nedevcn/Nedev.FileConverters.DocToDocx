# Nedev.DocToDocx

A high‑fidelity `.doc` → `.docx` converter for .NET 10 with no third‑party dependencies.

## Features

- **Binary `.doc` reader**: Implements core MS‑DOC structures (CFB, FIB, CLX/Piece Table, CHPX/PAPX FKPs, PLCFs).
- **Rich text & styles**: Fonts, font sizes, bold/italic/underline, colors, highlighting, outline/emboss/shadow, language (`w:lang`), paragraph alignment, spacing, indentation, borders, shading, and numbered/bulleted lists (including many localized formats).
- **Tables**: Multi‑row/column tables with TAP‑driven layout: row height and exact/at‑least rules, header rows, `cantSplit`, per‑cell width, proper vertical merges (`vMerge restart/continue`), horizontal merges (`gridSpan`), table‑level borders and shading mapped to DOCX.
- **Sections & page setup**: Multiple sections with page size/orientation, margins, starting page number, and First/Odd/Even headers/footers mapped to separate DOCX parts.
- **Images**: Extracts embedded images from the `WordDocument` and `Data` streams (PNG/JPEG/GIF/BMP/OfficeArt BLIPs), writes `word/media/*`, generates `w:drawing` with size inferred from image dimensions and auto‑scaled to page width, respects per‑image scale, and attaches basic alt text.
- **OfficeArt pictures & floating anchors**: Parses Escher/OfficeArt records and FSPA anchors from `PlcSpaMom` to recover picture shapes; maps them to `wp:anchor` floating images positioned relative to the page, falling back to inline images when anchors are unavailable.
- **Footnotes, endnotes, comments, textboxes**: Reads and writes common note and annotation structures into DOCX footnotes/endnotes parts and DrawingML textboxes.
- **Encryption (XOR)**: Supports Word’s XOR‑obfuscated streams via `EncryptionHelper` and decrypted CFB streams.
- **No external dependencies**: Pure .NET, streaming writers (`XmlWriter`) for high performance and low memory usage.

> Note: While many MS‑DOC features are implemented (including OfficeArt‑based picture extraction and floating anchors), the converter does not yet claim 100% coverage of the full [MS‑DOC] / [MS‑ODRAW] specifications. Complex vector shapes, SmartArt, charts, OLE objects, and some rare formatting cases remain intentionally out of scope for now.

## Library usage

Add a reference to the `Nedev.DocToDocx` assembly and call the static converter API:

```csharp
using Nedev.DocToDocx;

DocToDocxConverter.Convert("input.doc", "output.docx");
```

> This repository currently exposes the converter as a **library API only**.  
> If you need a CLI or additional validation tooling, you can build it on top of `DocToDocxConverter` according to your own application’s needs.