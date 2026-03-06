using System.Xml;
using Nedev.DocToDocx.Models;
using Nedev.DocToDocx.Utils;

namespace Nedev.DocToDocx.Writers;

/// <summary>
/// Writes footnotes and endnotes XML for DOCX
/// </summary>
public class FootnotesWriter
{
    private readonly XmlWriter _writer;
    private DocumentModel? _document;
    private Dictionary<int, string>? _imageIndexToRelId;
    private const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public FootnotesWriter(XmlWriter writer)
    {
        _writer = writer;
    }

    /// <summary>
    /// Writes footnotes XML (optionally with picture support when document and imageRelMap are provided).
    /// </summary>
    public void WriteFootnotes(List<FootnoteModel> footnotes, DocumentModel? document = null, Dictionary<int, string>? imageIndexToRelId = null)
    {
        _document = document;
        _imageIndexToRelId = imageIndexToRelId;
        _writer.WriteStartDocument();
        _writer.WriteStartElement("w", "footnotes", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteAttributeString("xmlns", "w", null, "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        if (imageIndexToRelId != null && imageIndexToRelId.Count > 0)
        {
            _writer.WriteAttributeString("xmlns", "wp", null, "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            _writer.WriteAttributeString("xmlns", "a", null, "http://schemas.openxmlformats.org/drawingml/2006/main");
            _writer.WriteAttributeString("xmlns", "pic", null, "http://schemas.openxmlformats.org/drawingml/2006/picture");
        }

        // Required separator and continuation separator for better compatibility
        WriteSeparatorFootnote(-1, "separator");
        WriteSeparatorFootnote(0, "continuationSeparator");

        foreach (var footnote in footnotes)
        {
            WriteFootnote(footnote, "footnote");
        }

        _writer.WriteEndElement();
        _writer.WriteEndDocument();
        _document = null;
        _imageIndexToRelId = null;
    }

    private void WriteSeparatorFootnote(int id, string type)
    {
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        _writer.WriteStartElement("w", "footnote", wNs);
        _writer.WriteAttributeString("w", "type", null, type);
        _writer.WriteAttributeString("w", "id", null, id.ToString());

        _writer.WriteStartElement("w", "p", wNs);
        _writer.WriteStartElement("w", "r", wNs);
        _writer.WriteStartElement("w", "separator", wNs);
        _writer.WriteEndElement(); // w:separator
        _writer.WriteEndElement(); // w:r
        _writer.WriteEndElement(); // w:p

        _writer.WriteEndElement(); // w:footnote
    }

    /// <summary>
    /// Writes endnotes XML (optionally with picture support when document and imageRelMap are provided).
    /// </summary>
    public void WriteEndnotes(List<EndnoteModel> endnotes, DocumentModel? document = null, Dictionary<int, string>? imageIndexToRelId = null)
    {
        _document = document;
        _imageIndexToRelId = imageIndexToRelId;
        _writer.WriteStartDocument();
        _writer.WriteStartElement("w", "endnotes", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteAttributeString("xmlns", "w", null, "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        if (imageIndexToRelId != null && imageIndexToRelId.Count > 0)
        {
            _writer.WriteAttributeString("xmlns", "wp", null, "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            _writer.WriteAttributeString("xmlns", "a", null, "http://schemas.openxmlformats.org/drawingml/2006/main");
            _writer.WriteAttributeString("xmlns", "pic", null, "http://schemas.openxmlformats.org/drawingml/2006/picture");
        }

        foreach (var endnote in endnotes)
        {
            WriteFootnote(endnote, "endnote");
        }

        _writer.WriteEndElement();
        _writer.WriteEndDocument();
        _document = null;
        _imageIndexToRelId = null;
    }

    private void WriteFootnote(NoteModelBase note, string type)
    {
        _writer.WriteStartElement("w", type, wNs);
        _writer.WriteAttributeString("w", "id", null, note.Index.ToString());

        // Write paragraph
        foreach (var paragraph in note.Paragraphs)
        {
            WriteParagraph(paragraph);
        }

        _writer.WriteEndElement();
    }

    private void WriteParagraph(ParagraphModel paragraph)
    {
        _writer.WriteStartElement("w", "p", wNs);

        // Write runs
        foreach (var run in paragraph.Runs)
        {
            WriteRun(run);
        }

        _writer.WriteEndElement();
    }

    private void WriteRun(RunModel run)
    {
        // Picture run: write w:drawing when document and image rel map are available
        if (run.IsPicture && run.ImageIndex >= 0 && _document != null && _imageIndexToRelId != null &&
            _imageIndexToRelId.TryGetValue(run.ImageIndex, out var relId) &&
            run.ImageIndex < _document.Images.Count)
        {
            var image = _document.Images[run.ImageIndex];
            if (image.Data != null && image.Data.Length > 0)
            {
                WritePictureRun(image, relId, run.ImageIndex + 1);
                return;
            }
        }

        _writer.WriteStartElement("w", "r", wNs);

        if (run.Properties != null)
        {
            _writer.WriteStartElement("w", "rPr", wNs);
            if (!string.IsNullOrEmpty(run.Properties.FontName))
            {
                _writer.WriteStartElement("w", "rFonts", wNs);
                _writer.WriteAttributeString("w", "ascii", null, run.Properties.FontName);
                _writer.WriteAttributeString("w", "hAnsi", null, run.Properties.FontName);
                _writer.WriteEndElement();
            }
            if (run.Properties.FontSize > 0)
            {
                _writer.WriteStartElement("w", "sz", wNs);
                _writer.WriteAttributeString("w", "val", null, run.Properties.FontSize.ToString());
                _writer.WriteEndElement();
            }
            if (run.Properties.IsBold) { _writer.WriteStartElement("w", "b", wNs); _writer.WriteEndElement(); }
            if (run.Properties.IsItalic) { _writer.WriteStartElement("w", "i", wNs); _writer.WriteEndElement(); }
            var colorHex = ColorHelper.ColorToHex(run.Properties.Color);
            if (colorHex != "auto")
            {
                _writer.WriteStartElement("w", "color", wNs);
                _writer.WriteAttributeString("w", "val", null, colorHex);
                _writer.WriteEndElement();
            }
            _writer.WriteEndElement();
        }

        _writer.WriteStartElement("w", "t", wNs);
        if (!string.IsNullOrEmpty(run.Text))
        {
            // sanitize text to remove any embedded nulls or illegal XML chars (see DocumentWriter)
            var safe = Nedev.DocToDocx.Writers.DocumentWriter.SanitizeXmlString(run.Text);
            if (!string.IsNullOrEmpty(safe))
            {
                if (safe.StartsWith(' ') || safe.EndsWith(' ') || safe.Contains("  "))
                    _writer.WriteAttributeString("xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve");
                _writer.WriteString(safe);
            }
        }
        _writer.WriteEndElement();
        _writer.WriteEndElement();
    }

    private void WritePictureRun(ImageModel image, string relId, int imageId)
    {
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string wpNs = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        const string aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const string picNs = "http://schemas.openxmlformats.org/drawingml/2006/picture";
        const string rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        var widthEmu = image.WidthEMU > 0 ? image.WidthEMU : 5715000;
        var heightEmu = image.HeightEMU > 0 ? image.HeightEMU : 3810000;
        if (image.ScaleX > 0 && image.ScaleX != 100000) widthEmu = (int)(widthEmu * (image.ScaleX / 100000.0));
        if (image.ScaleY > 0 && image.ScaleY != 100000) heightEmu = (int)(heightEmu * (image.ScaleY / 100000.0));

        _writer.WriteStartElement("w", "r", wNs);
        _writer.WriteStartElement("w", "drawing", wNs);
        _writer.WriteStartElement("wp", "inline", wpNs);
        _writer.WriteAttributeString("distT", "0");
        _writer.WriteAttributeString("distB", "0");
        _writer.WriteAttributeString("distL", "0");
        _writer.WriteAttributeString("distR", "0");
        _writer.WriteStartElement("wp", "extent", wpNs);
        _writer.WriteAttributeString("cx", widthEmu.ToString());
        _writer.WriteAttributeString("cy", heightEmu.ToString());
        _writer.WriteEndElement();
        _writer.WriteStartElement("wp", "effectExtent", wpNs);
        _writer.WriteAttributeString("l", "0"); _writer.WriteAttributeString("t", "0"); _writer.WriteAttributeString("r", "0"); _writer.WriteAttributeString("b", "0");
        _writer.WriteEndElement();
        _writer.WriteStartElement("wp", "docPr", wpNs);
        _writer.WriteAttributeString("id", imageId.ToString());
        _writer.WriteAttributeString("name", image.FileName);
        _writer.WriteEndElement();
        _writer.WriteStartElement("wp", "cNvGraphicFramePr", wpNs);
        _writer.WriteStartElement("a", "graphicFrameLocks", aNs);
        _writer.WriteAttributeString("noChangeAspect", "1");
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteStartElement("a", "graphic", aNs);
        _writer.WriteStartElement("a", "graphicData", aNs);
        _writer.WriteAttributeString("uri", "http://schemas.openxmlformats.org/drawingml/2006/picture");
        _writer.WriteStartElement("pic", "pic", picNs);
        _writer.WriteStartElement("pic", "nvPicPr", picNs);
        _writer.WriteStartElement("pic", "cNvPr", picNs);
        _writer.WriteAttributeString("id", "0");
        _writer.WriteAttributeString("name", image.FileName);
        _writer.WriteEndElement();
        _writer.WriteStartElement("pic", "cNvPicPr", picNs);
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteStartElement("pic", "blipFill", picNs);
        _writer.WriteStartElement("a", "blip", aNs);
        _writer.WriteAttributeString("r", "embed", rNs, relId);
        _writer.WriteEndElement();
        _writer.WriteStartElement("a", "stretch", aNs);
        _writer.WriteStartElement("a", "fillRect", aNs);
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteStartElement("pic", "spPr", picNs);
        _writer.WriteStartElement("a", "xfrm", aNs);
        _writer.WriteStartElement("a", "off", aNs);
        _writer.WriteAttributeString("x", "0"); _writer.WriteAttributeString("y", "0");
        _writer.WriteEndElement();
        _writer.WriteStartElement("a", "ext", aNs);
        _writer.WriteAttributeString("cx", widthEmu.ToString());
        _writer.WriteAttributeString("cy", heightEmu.ToString());
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteStartElement("a", "prstGeom", aNs);
        _writer.WriteAttributeString("prst", "rect");
        _writer.WriteStartElement("a", "avLst", aNs);
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteEndElement();
    }
}
