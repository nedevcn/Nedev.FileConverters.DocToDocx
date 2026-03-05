using System.Xml;
using Nedev.DocToDocx.Models;
using Nedev.DocToDocx.Utils;

namespace Nedev.DocToDocx.Writers;

/// <summary>
/// DocumentWriter partial class — shape, vector, and floating picture writing methods.
/// </summary>
public partial class DocumentWriter
{
    /// <summary>
    /// Writes a basic vector shape (non-picture OfficeArt / SmartArt fallback)
    /// as a DrawingML wordprocessingShape rectangle. Position and size are
    /// taken from ShapeAnchor when available; otherwise a reasonable inline
    /// size is used.
    /// </summary>
    private void WriteVectorShape(ShapeModel shape)
    {
        if (_document == null)
            return;

        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string wpNs = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        const string aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const string wpsNs = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

        // Derive size in EMUs from anchor or use a default (~4x3 inches).
        const int emuPerTwip = 635;
        int widthEmu = 3657600;  // 4"
        int heightEmu = 2743200; // 3"
        int xEmu = 0;
        int yEmu = 0;
        bool isFloating = false;

        if (shape.Anchor != null)
        {
            var anchor = shape.Anchor;
            if (anchor.Width > 0)
            {
                widthEmu = anchor.Width * emuPerTwip;
            }
            if (anchor.Height > 0)
            {
                heightEmu = anchor.Height * emuPerTwip;
            }
            xEmu = Math.Max(0, anchor.X * emuPerTwip);
            yEmu = Math.Max(0, anchor.Y * emuPerTwip);
            isFloating = anchor.IsFloating;
        }

        // Clamp width to page width (inside margins) while preserving aspect ratio.
        if (_document.Properties != null)
        {
            var page = _document.Properties;
            var maxWidthTwips = page.PageWidth - page.MarginLeft - page.MarginRight;
            if (maxWidthTwips > 0)
            {
                var maxWidthEmu = maxWidthTwips * emuPerTwip;
                if (widthEmu > maxWidthEmu && widthEmu > 0 && heightEmu > 0)
                {
                    var scale = (double)maxWidthEmu / widthEmu;
                    widthEmu = maxWidthEmu;
                    heightEmu = (int)(heightEmu * scale);
                }
            }
        }

        _writer.WriteStartElement("w", "p", wNs);
        _writer.WriteStartElement("w", "r", wNs);
        _writer.WriteStartElement("w", "drawing", wNs);

        if (isFloating)
        {
            // Floating vector shape using wp:anchor (similar to floating picture).
            _writer.WriteStartElement("wp", "anchor", wpNs);
            _writer.WriteAttributeString("distT", "0");
            _writer.WriteAttributeString("distB", "0");
            _writer.WriteAttributeString("distL", "0");
            _writer.WriteAttributeString("distR", "0");
            _writer.WriteAttributeString("simplePos", "0");
            var relHeight = shape.Anchor?.ZOrder ?? 0;
            if (relHeight < 0) relHeight = 0;
            _writer.WriteAttributeString("relativeHeight", relHeight.ToString());
            _writer.WriteAttributeString("behindDoc", shape.Anchor?.WrapType == ShapeWrapType.BehindText ? "1" : "0");
            _writer.WriteAttributeString("locked", "0");
            _writer.WriteAttributeString("layoutInCell", "1");
            _writer.WriteAttributeString("allowOverlap", "1");

            // Horizontal & vertical position.
            _writer.WriteStartElement("wp", "positionH", wpNs);
            _writer.WriteAttributeString("relativeFrom", GetOOXMLRelativeTo(shape.Anchor?.HorizontalRelativeTo ?? ShapeRelativeTo.Page));
            _writer.WriteStartElement("wp", "posOffset", wpNs);
            _writer.WriteString(xEmu.ToString());
            _writer.WriteEndElement(); // wp:posOffset
            _writer.WriteEndElement(); // wp:positionH

            _writer.WriteStartElement("wp", "positionV", wpNs);
            _writer.WriteAttributeString("relativeFrom", GetOOXMLRelativeTo(shape.Anchor?.VerticalRelativeTo ?? ShapeRelativeTo.Page));
            _writer.WriteStartElement("wp", "posOffset", wpNs);
            _writer.WriteString(yEmu.ToString());
            _writer.WriteEndElement(); // wp:posOffset
            _writer.WriteEndElement(); // wp:positionV

            // Extent
            _writer.WriteStartElement("wp", "extent", wpNs);
            _writer.WriteAttributeString("cx", widthEmu.ToString());
            _writer.WriteAttributeString("cy", heightEmu.ToString());
            _writer.WriteEndElement();

            // Effect extent
            _writer.WriteStartElement("wp", "effectExtent", wpNs);
            _writer.WriteAttributeString("l", "0");
            _writer.WriteAttributeString("t", "0");
            _writer.WriteAttributeString("r", "0");
            _writer.WriteAttributeString("b", "0");
            _writer.WriteEndElement();

            // Text wrapping
            WriteWrapMode(shape.Anchor?.WrapType ?? ShapeWrapType.Square);

            // docPr
            _writer.WriteStartElement("wp", "docPr", wpNs);
            _writer.WriteAttributeString("id", (2000 + shape.Id).ToString());
            _writer.WriteAttributeString("name", $"Shape {shape.Id}");
            _writer.WriteEndElement(); // wp:docPr

            // Graphic frame props
            _writer.WriteStartElement("wp", "cNvGraphicFramePr", wpNs);
            _writer.WriteStartElement("a", "graphicFrameLocks", aNs);
            _writer.WriteAttributeString("noChangeAspect", "1");
            _writer.WriteEndElement();
            _writer.WriteEndElement();

            // Graphic
            _writer.WriteStartElement("a", "graphic", aNs);
            _writer.WriteStartElement("a", "graphicData", aNs);
            _writer.WriteAttributeString("uri", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            // WordprocessingShape (rectangle/ellipse etc.)
            WriteWpsShape(shape, widthEmu, heightEmu);

            _writer.WriteEndElement(); // a:graphicData
            _writer.WriteEndElement(); // a:graphic

            _writer.WriteEndElement(); // wp:anchor
        }
        else
        {
            // Inline vector shape using wp:inline.
            _writer.WriteStartElement("wp", "inline", wpNs);
            _writer.WriteAttributeString("distT", "0");
            _writer.WriteAttributeString("distB", "0");
            _writer.WriteAttributeString("distL", "0");
            _writer.WriteAttributeString("distR", "0");

            _writer.WriteStartElement("wp", "extent", wpNs);
            _writer.WriteAttributeString("cx", widthEmu.ToString());
            _writer.WriteAttributeString("cy", heightEmu.ToString());
            _writer.WriteEndElement(); // wp:extent

            _writer.WriteStartElement("wp", "effectExtent", wpNs);
            _writer.WriteAttributeString("l", "0");
            _writer.WriteAttributeString("t", "0");
            _writer.WriteAttributeString("r", "0");
            _writer.WriteAttributeString("b", "0");
            _writer.WriteEndElement(); // wp:effectExtent

            _writer.WriteStartElement("wp", "docPr", wpNs);
            _writer.WriteAttributeString("id", (2000 + shape.Id).ToString());
            _writer.WriteAttributeString("name", $"Shape {shape.Id}");
            _writer.WriteEndElement(); // wp:docPr

            _writer.WriteStartElement("wp", "cNvGraphicFramePr", wpNs);
            _writer.WriteStartElement("a", "graphicFrameLocks", aNs);
            _writer.WriteAttributeString("noChangeAspect", "1");
            _writer.WriteEndElement();
            _writer.WriteEndElement();

            _writer.WriteStartElement("a", "graphic", aNs);
            _writer.WriteStartElement("a", "graphicData", aNs);
            _writer.WriteAttributeString("uri", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            WriteWpsShape(shape, widthEmu, heightEmu);

            _writer.WriteEndElement(); // a:graphicData
            _writer.WriteEndElement(); // a:graphic

            _writer.WriteEndElement(); // wp:inline
        }

        _writer.WriteEndElement(); // w:drawing
        _writer.WriteEndElement(); // w:r
        _writer.WriteEndElement(); // w:p
    }

    /// <summary>
    /// Writes the inner wps:wsp contents (geometry and basic styling) for a
    /// vector shape.
    /// </summary>
    private void WriteWpsShape(ShapeModel shape, int widthEmu, int heightEmu)
    {
        const string aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const string wpsNs = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

        _writer.WriteStartElement("wps", "wsp", wpsNs);

        // spPr
        _writer.WriteStartElement("wps", "spPr", wpsNs);
        _writer.WriteStartElement("a", "xfrm", aNs);
        _writer.WriteStartElement("a", "off", aNs);
        _writer.WriteAttributeString("x", "0");
        _writer.WriteAttributeString("y", "0");
        _writer.WriteEndElement(); // a:off
        _writer.WriteStartElement("a", "ext", aNs);
        _writer.WriteAttributeString("cx", widthEmu.ToString());
        _writer.WriteAttributeString("cy", heightEmu.ToString());
        _writer.WriteEndElement(); // a:ext
        _writer.WriteEndElement(); // a:xfrm

        if (shape.CustomGeometry != null)
        {
            WriteCustomGeometry(shape.CustomGeometry);
        }
        else
        {
            // Preset geometry (rectangle by default)
            var prst = shape.Type switch
            {
                ShapeType.Ellipse => "ellipse",
                ShapeType.Textbox => "rect",
                _ => "rect"
            };
            _writer.WriteStartElement("a", "prstGeom", aNs);
            _writer.WriteAttributeString("prst", prst);
            _writer.WriteStartElement("a", "avLst", aNs);
            _writer.WriteEndElement();
            _writer.WriteEndElement();
        }

        // Fill and line styling
        WriteShapeStyling(shape);

        _writer.WriteEndElement(); // wps:spPr

        // Optionally, basic textbox for shape text when available.
        if (!string.IsNullOrEmpty(shape.Text))
        {
            _writer.WriteStartElement("wps", "txbx", wpsNs);
            _writer.WriteStartElement("w", "txbxContent", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteStartElement("w", "p", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteStartElement("w", "r", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteStartElement("w", "t", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteAttributeString("xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve");
            _writer.WriteString(shape.Text);
            _writer.WriteEndElement(); // w:t
            _writer.WriteEndElement(); // w:r
            _writer.WriteEndElement(); // w:p
            _writer.WriteEndElement(); // w:txbxContent
            _writer.WriteEndElement(); // wps:txbx
        }

        _writer.WriteEndElement(); // wps:wsp
    }

    private void WriteCustomGeometry(CustomGeometry geom)
    {
        const string aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
        _writer.WriteStartElement("a", "custGeom", aNs);
        _writer.WriteStartElement("a", "avLst", aNs); _writer.WriteEndElement();
        _writer.WriteStartElement("a", "gdLst", aNs); _writer.WriteEndElement();
        _writer.WriteStartElement("a", "ahLst", aNs); _writer.WriteEndElement();
        _writer.WriteStartElement("a", "cxnLst", aNs); _writer.WriteEndElement();
        _writer.WriteStartElement("a", "rect", aNs);
        _writer.WriteAttributeString("l", "0"); _writer.WriteAttributeString("t", "0");
        _writer.WriteAttributeString("r", "0"); _writer.WriteAttributeString("b", "0");
        _writer.WriteEndElement();

        _writer.WriteStartElement("a", "pathLst", aNs);
        _writer.WriteStartElement("a", "path", aNs);
        _writer.WriteAttributeString("w", Math.Max(1, geom.ViewRight - geom.ViewLeft).ToString());
        _writer.WriteAttributeString("h", Math.Max(1, geom.ViewBottom - geom.ViewTop).ToString());

        foreach (var segment in geom.Segments)
        {
            switch (segment.Type)
            {
                case SegmentType.MoveTo:
                    if (segment.VertexIndex < geom.Vertices.Count)
                    {
                        var pt = geom.Vertices[segment.VertexIndex];
                        _writer.WriteStartElement("a", "moveTo", aNs);
                        _writer.WriteStartElement("a", "pt", aNs);
                        _writer.WriteAttributeString("x", (pt.X - geom.ViewLeft).ToString());
                        _writer.WriteAttributeString("y", (pt.Y - geom.ViewTop).ToString());
                        _writer.WriteEndElement(); _writer.WriteEndElement();
                    }
                    break;
                case SegmentType.LineTo:
                    if (segment.VertexIndex < geom.Vertices.Count)
                    {
                        var pt = geom.Vertices[segment.VertexIndex];
                        _writer.WriteStartElement("a", "lnTo", aNs);
                        _writer.WriteStartElement("a", "pt", aNs);
                        _writer.WriteAttributeString("x", (pt.X - geom.ViewLeft).ToString());
                        _writer.WriteAttributeString("y", (pt.Y - geom.ViewTop).ToString());
                        _writer.WriteEndElement(); _writer.WriteEndElement();
                    }
                    break;
                case SegmentType.CurveTo:
                    // Escher curves are usually 3 vertices (Bezier)
                    if (segment.VertexIndex + 2 < geom.Vertices.Count)
                    {
                        _writer.WriteStartElement("a", "cubicBezTo", aNs);
                        for (int i = 0; i < 3; i++)
                        {
                            var pt = geom.Vertices[segment.VertexIndex + i];
                            _writer.WriteStartElement("a", "pt", aNs);
                            _writer.WriteAttributeString("x", (pt.X - geom.ViewLeft).ToString());
                            _writer.WriteAttributeString("y", (pt.Y - geom.ViewTop).ToString());
                            _writer.WriteEndElement();
                        }
                        _writer.WriteEndElement();
                    }
                    break;
                case SegmentType.Close:
                    _writer.WriteStartElement("a", "close", aNs); _writer.WriteEndElement();
                    break;
            }
        }
        _writer.WriteEndElement(); // a:path
        _writer.WriteEndElement(); // a:pathLst
        _writer.WriteEndElement(); // a:custGeom
    }

    private void WriteShapeStyling(ShapeModel shape)
    {
        const string aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
        
        // Background fill
        if (shape.FillColor != 0)
        {
            _writer.WriteStartElement("a", "solidFill", aNs);
            _writer.WriteStartElement("a", "srgbClr", aNs);
            _writer.WriteAttributeString("val", ColorHelper.ColorToHex(shape.FillColor));
            _writer.WriteEndElement(); _writer.WriteEndElement();
        }
        else
        {
            _writer.WriteStartElement("a", "noFill", aNs); _writer.WriteEndElement();
        }

        // Outline
        if (shape.IsLineVisible)
        {
            _writer.WriteStartElement("a", "ln", aNs);
            _writer.WriteAttributeString("w", shape.LineWidth > 0 ? shape.LineWidth.ToString() : "9525"); // Default 1pt
            
            _writer.WriteStartElement("a", "solidFill", aNs);
            _writer.WriteStartElement("a", "srgbClr", aNs);
            _writer.WriteAttributeString("val", ColorHelper.ColorToHex(shape.LineColor != 0 ? shape.LineColor : 0));
            _writer.WriteEndElement(); _writer.WriteEndElement();
            
            _writer.WriteEndElement(); // a:ln
        }
    }

    /// <summary>
    /// Writes a single shape. Picture shapes are rendered either as true
    /// floating images (wp:anchor) when ShapeAnchor.IsFloating is set, or as
    /// inline pictures using the existing run-based logic. Non-picture shapes
    /// (basic OfficeArt vectors and SmartArt fallbacks) are rendered as
    /// DrawingML wordprocessingShape rectangles with simple fill/line styling.
    /// </summary>
    private void WriteInlinePictureShape(ShapeModel shape, DocumentModel document)
    {
        // Picture-backed shapes
        if (shape.Type == ShapeType.Picture)
        {
            if (shape.ImageIndex is null)
                return;

            var imageIndex = shape.ImageIndex.Value;
            if (imageIndex < 0 || imageIndex >= document.Images.Count)
                return;

            // Prefer floating output when we have a valid anchor.
            if (shape.Anchor is { IsFloating: true })
            {
                WriteFloatingPictureShape(shape, document);
                return;
            }

            // Fallback: inline picture using existing run-based logic.
            var run = new RunModel
            {
                IsPicture = true,
                ImageIndex = imageIndex,
                Properties = new RunProperties(),
                CropTop = shape.CropTop,
                CropBottom = shape.CropBottom,
                CropLeft = shape.CropLeft,
                CropRight = shape.CropRight
            };

            _writer.WriteStartElement("w", "p", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteStartElement("w", "r", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            WriteRun(run);
            _writer.WriteEndElement(); // w:r
            _writer.WriteEndElement(); // w:p
            return;
        }

        // Non-picture vector shapes (complex OfficeArt, SmartArt fallbacks)
        WriteVectorShape(shape);
    }
    
    /// <summary>
    /// Writes a floating picture using wp:anchor based on ShapeAnchor coordinates.
    /// </summary>
    private void WriteFloatingPictureShape(ShapeModel shape, DocumentModel document)
    {
        if (shape.ImageIndex is null || _document == null)
            return;

        var imageIndex = shape.ImageIndex.Value;
        if (imageIndex < 0 || imageIndex >= document.Images.Count)
            return;

        var image = document.Images[imageIndex];
        if (image.Data == null || image.Data.Length == 0)
            return;

        var anchor = shape.Anchor!;

        // Relationship ID and doc-level image id
        var ids = RelationshipsWriter.ComputeRelationshipIds(_document);
        var imageId = imageIndex + 1;

        // Compute size in EMUs, preferring anchor size when available.
        const int emuPerTwip = 635;
        int widthEmu;
        int heightEmu;

        if (anchor.Width > 0 && anchor.Height > 0)
        {
            widthEmu = anchor.Width * emuPerTwip;
            heightEmu = anchor.Height * emuPerTwip;
        }
        else
        {
            widthEmu = image.WidthEMU > 0 ? image.WidthEMU : 5715000;
            heightEmu = image.HeightEMU > 0 ? image.HeightEMU : 3810000;
        }

        // Respect per-image scale factors
        if (image.ScaleX > 0 && image.ScaleX != 100000)
        {
            widthEmu = (int)(widthEmu * (image.ScaleX / 100000.0));
        }
        if (image.ScaleY > 0 && image.ScaleY != 100000)
        {
            heightEmu = (int)(heightEmu * (image.ScaleY / 100000.0));
        }

        // Full-page background: first body picture or anchor/size close to page → full page dimensions
        if (_document.Properties != null)
        {
            var page = _document.Properties;
            int pageWidthEmu = page.PageWidth * emuPerTwip;
            int pageHeightEmu = page.PageHeight * emuPerTwip;
            bool forceFirstFullPage = _firstBodyPictureNotYetWritten && pageWidthEmu > 0 && pageHeightEmu > 0;
            bool looksFullPage = !forceFirstFullPage && (pageWidthEmu > 0 && pageHeightEmu > 0) &&
                (widthEmu >= pageWidthEmu * 0.85 || heightEmu >= pageHeightEmu * 0.85);
            if (forceFirstFullPage || looksFullPage)
            {
                widthEmu = pageWidthEmu;
                heightEmu = pageHeightEmu;
                if (forceFirstFullPage) _firstBodyPictureNotYetWritten = false;
            }
            else
            {
                var maxWidthTwips = page.PageWidth - page.MarginLeft - page.MarginRight;
                if (maxWidthTwips > 0)
                {
                    var maxWidthEmu = maxWidthTwips * emuPerTwip;
                    if (widthEmu > maxWidthEmu && widthEmu > 0 && heightEmu > 0)
                    {
                        var scale = (double)maxWidthEmu / widthEmu;
                        widthEmu = maxWidthEmu;
                        heightEmu = (int)(heightEmu * scale);
                    }
                }
            }
        }

        // Convert anchor position from twips to EMUs; clamp to non-negative.
        var xEmu = Math.Max(0, anchor.X * emuPerTwip);
        var yEmu = Math.Max(0, anchor.Y * emuPerTwip);

        _writer.WriteStartElement("w", "p", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteStartElement("w", "r", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteStartElement("w", "drawing", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

        // Floating anchor
        _writer.WriteStartElement("wp", "anchor", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        _writer.WriteAttributeString("distT", "0");
        _writer.WriteAttributeString("distB", "0");
        _writer.WriteAttributeString("distL", "0");
        _writer.WriteAttributeString("distR", "0");
        _writer.WriteAttributeString("simplePos", "0");
        var relHeight = anchor.ZOrder;
        if (relHeight < 0) relHeight = 0;
        _writer.WriteAttributeString("relativeHeight", relHeight.ToString());
        _writer.WriteAttributeString("relativeHeight", relHeight.ToString());
        _writer.WriteAttributeString("behindDoc", anchor.WrapType == ShapeWrapType.BehindText ? "1" : "0");
        _writer.WriteAttributeString("locked", "0");
        _writer.WriteAttributeString("layoutInCell", "1");
        _writer.WriteAttributeString("allowOverlap", "1");

        const string wpNs = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        
        // Position
        _writer.WriteStartElement("wp", "positionH", wpNs);
        _writer.WriteAttributeString("relativeFrom", GetOOXMLRelativeTo(anchor.HorizontalRelativeTo));
        _writer.WriteStartElement("wp", "posOffset", wpNs);
        _writer.WriteString(xEmu.ToString());
        _writer.WriteEndElement(); // wp:posOffset
        _writer.WriteEndElement(); // wp:positionH

        _writer.WriteStartElement("wp", "positionV", wpNs);
        _writer.WriteAttributeString("relativeFrom", GetOOXMLRelativeTo(anchor.VerticalRelativeTo));
        _writer.WriteStartElement("wp", "posOffset", wpNs);
        _writer.WriteString(yEmu.ToString());
        _writer.WriteEndElement(); // wp:posOffset
        _writer.WriteEndElement(); // wp:positionV

        // Extent
        _writer.WriteStartElement("wp", "extent", wpNs);
        _writer.WriteAttributeString("cx", widthEmu.ToString());
        _writer.WriteAttributeString("cy", heightEmu.ToString());
        _writer.WriteEndElement();

        // Effect extent
        _writer.WriteStartElement("wp", "effectExtent", wpNs);
        _writer.WriteAttributeString("l", "0");
        _writer.WriteAttributeString("t", "0");
        _writer.WriteAttributeString("r", "0");
        _writer.WriteAttributeString("b", "0");
        _writer.WriteEndElement();

        // Text wrapping
        WriteWrapMode(anchor.WrapType);

        // Doc properties
        _writer.WriteStartElement("wp", "docPr", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        _writer.WriteAttributeString("id", imageId.ToString());
        var baseName = !string.IsNullOrEmpty(image.FileName) ? image.FileName : $"Picture {imageId}";
        _writer.WriteAttributeString("name", baseName);
        var altText = baseName;
        var dotIndex = baseName.LastIndexOf('.');
        if (dotIndex > 0)
        {
            altText = baseName.Substring(0, dotIndex);
        }
        _writer.WriteAttributeString("descr", altText);
        _writer.WriteEndElement(); // wp:docPr

        // Non-visual graphic frame properties
        _writer.WriteStartElement("wp", "cNvGraphicFramePr", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        _writer.WriteStartElement("a", "graphicFrameLocks", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteAttributeString("noChangeAspect", "1");
        _writer.WriteEndElement();
        _writer.WriteEndElement();

        // Graphic
        _writer.WriteStartElement("a", "graphic", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteStartElement("a", "graphicData", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteAttributeString("uri", "http://schemas.openxmlformats.org/drawingml/2006/picture");

        // Picture
        _writer.WriteStartElement("pic", "pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");

        // Non-visual picture properties
        _writer.WriteStartElement("pic", "nvPicPr", "http://schemas.openxmlformats.org/drawingml/2006/picture");
        _writer.WriteStartElement("pic", "cNvPr", "http://schemas.openxmlformats.org/drawingml/2006/picture");
        _writer.WriteAttributeString("id", "0");
        _writer.WriteAttributeString("name", image.FileName);
        _writer.WriteEndElement();
        _writer.WriteStartElement("pic", "cNvPicPr", "http://schemas.openxmlformats.org/drawingml/2006/picture");
        _writer.WriteEndElement();
        _writer.WriteEndElement(); // pic:nvPicPr

        // Blip fill
        _writer.WriteStartElement("pic", "blipFill", "http://schemas.openxmlformats.org/drawingml/2006/picture");
        _writer.WriteStartElement("a", "blip", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteAttributeString("r", "embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", $"rId{ids.FirstImageRId + imageIndex}");
        _writer.WriteEndElement();

        // Cropping
        if (shape.CropTop != 0 || shape.CropBottom != 0 || shape.CropLeft != 0 || shape.CropRight != 0)
        {
            _writer.WriteStartElement("a", "srcRect", "http://schemas.openxmlformats.org/drawingml/2006/main");
            if (shape.CropTop != 0) _writer.WriteAttributeString("t", ((long)shape.CropTop * 100000 / 65536).ToString());
            if (shape.CropBottom != 0) _writer.WriteAttributeString("b", ((long)shape.CropBottom * 100000 / 65536).ToString());
            if (shape.CropLeft != 0) _writer.WriteAttributeString("l", ((long)shape.CropLeft * 100000 / 65536).ToString());
            if (shape.CropRight != 0) _writer.WriteAttributeString("r", ((long)shape.CropRight * 100000 / 65536).ToString());
            _writer.WriteEndElement();
        }

        _writer.WriteStartElement("a", "stretch", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteStartElement("a", "fillRect", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteEndElement(); // pic:blipFill

        // Shape properties
        _writer.WriteStartElement("pic", "spPr", "http://schemas.openxmlformats.org/drawingml/2006/picture");
        _writer.WriteStartElement("a", "xfrm", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteStartElement("a", "off", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteAttributeString("x", "0");
        _writer.WriteAttributeString("y", "0");
        _writer.WriteEndElement();
        _writer.WriteStartElement("a", "ext", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteAttributeString("cx", widthEmu.ToString());
        _writer.WriteAttributeString("cy", heightEmu.ToString());
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteStartElement("a", "prstGeom", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteAttributeString("prst", "rect");
        _writer.WriteStartElement("a", "avLst", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteEndElement();
        _writer.WriteEndElement();

        // Line/Border
        if (shape.IsLineVisible && (shape.LineWidth > 0 || shape.LineColor != 0))
        {
            _writer.WriteStartElement("a", "ln", "http://schemas.openxmlformats.org/drawingml/2006/main");
            if (shape.LineWidth > 0) _writer.WriteAttributeString("w", shape.LineWidth.ToString());
            _writer.WriteStartElement("a", "solidFill", "http://schemas.openxmlformats.org/drawingml/2006/main");
            _writer.WriteStartElement("a", "srgbClr", "http://schemas.openxmlformats.org/drawingml/2006/main");
            _writer.WriteAttributeString("val", ColorHelper.ColorToHex(shape.LineColor == 0 ? 0 : shape.LineColor));
            _writer.WriteEndElement();
            _writer.WriteEndElement();
            _writer.WriteEndElement(); // a:ln
        }

        _writer.WriteEndElement(); // pic:spPr

        _writer.WriteEndElement(); // pic:pic
        _writer.WriteEndElement(); // a:graphicData
        _writer.WriteEndElement(); // a:graphic
        _writer.WriteEndElement(); // wp:anchor
        _writer.WriteEndElement(); // w:drawing
        _writer.WriteEndElement(); // w:r
        _writer.WriteEndElement(); // w:p
    }

    private static string GetOOXMLRelativeTo(ShapeRelativeTo relativeTo)
    {
        return relativeTo switch
        {
            ShapeRelativeTo.Margin => "margin",
            ShapeRelativeTo.Column => "column",
            ShapeRelativeTo.Paragraph => "paragraph",
            _ => "page"
        };
    }

    private void WriteWrapMode(ShapeWrapType wrapType)
    {
        const string wpNs = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        switch (wrapType)
        {
            case ShapeWrapType.Square:
                _writer.WriteStartElement("wp", "wrapSquare", wpNs);
                _writer.WriteAttributeString("wrapText", "bothSides");
                _writer.WriteEndElement();
                break;
            case ShapeWrapType.Tight:
                _writer.WriteStartElement("wp", "wrapTight", wpNs);
                _writer.WriteAttributeString("wrapText", "bothSides");
                _writer.WriteStartElement("wp", "wrapPolygon", wpNs);
                _writer.WriteAttributeString("edited", "0");
                _writer.WriteStartElement("wp", "start", wpNs); _writer.WriteAttributeString("x", "0"); _writer.WriteAttributeString("y", "0"); _writer.WriteEndElement();
                _writer.WriteStartElement("wp", "lineTo", wpNs); _writer.WriteAttributeString("x", "0"); _writer.WriteAttributeString("y", "21600"); _writer.WriteEndElement();
                _writer.WriteStartElement("wp", "lineTo", wpNs); _writer.WriteAttributeString("x", "21600"); _writer.WriteAttributeString("y", "21600"); _writer.WriteEndElement();
                _writer.WriteStartElement("wp", "lineTo", wpNs); _writer.WriteAttributeString("x", "21600"); _writer.WriteAttributeString("y", "0"); _writer.WriteEndElement();
                _writer.WriteStartElement("wp", "lineTo", wpNs); _writer.WriteAttributeString("x", "0"); _writer.WriteAttributeString("y", "0"); _writer.WriteEndElement();
                _writer.WriteEndElement();
                _writer.WriteEndElement();
                break;
            case ShapeWrapType.Through:
                _writer.WriteStartElement("wp", "wrapThrough", wpNs);
                _writer.WriteAttributeString("wrapText", "bothSides");
                _writer.WriteStartElement("wp", "wrapPolygon", wpNs);
                _writer.WriteAttributeString("edited", "0");
                _writer.WriteStartElement("wp", "start", wpNs); _writer.WriteAttributeString("x", "0"); _writer.WriteAttributeString("y", "0"); _writer.WriteEndElement();
                _writer.WriteStartElement("wp", "lineTo", wpNs); _writer.WriteAttributeString("x", "0"); _writer.WriteAttributeString("y", "21600"); _writer.WriteEndElement();
                _writer.WriteStartElement("wp", "lineTo", wpNs); _writer.WriteAttributeString("x", "21600"); _writer.WriteAttributeString("y", "21600"); _writer.WriteEndElement();
                _writer.WriteStartElement("wp", "lineTo", wpNs); _writer.WriteAttributeString("x", "21600"); _writer.WriteAttributeString("y", "0"); _writer.WriteEndElement();
                _writer.WriteStartElement("wp", "lineTo", wpNs); _writer.WriteAttributeString("x", "0"); _writer.WriteAttributeString("y", "0"); _writer.WriteEndElement();
                _writer.WriteEndElement();
                _writer.WriteEndElement();
                break;
            case ShapeWrapType.TopBottom:
                _writer.WriteStartElement("wp", "wrapTopAndBottom", wpNs);
                _writer.WriteEndElement();
                break;
            default:
                _writer.WriteStartElement("wp", "wrapNone", wpNs);
                _writer.WriteEndElement();
                break;
        }
    }
}
