using System.Xml;
using Nedev.DocToDocx.Models;

namespace Nedev.DocToDocx.Writers;

/// <summary>
/// Writes the word/comments.xml part for standard track-changes and annotations support.
/// </summary>
public class CommentsWriter
{
    private readonly XmlWriter _writer;

    public CommentsWriter(XmlWriter writer)
    {
        _writer = writer;
    }

    public void WriteComments(DocumentModel document)
    {
        _writer.WriteStartDocument();
        _writer.WriteStartElement("w", "comments", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        
        // Ensure necessary namespaces for drawing/relationships if comments have pictures
        _writer.WriteAttributeString("xmlns", "w", null, "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        int commentId = 0;
        foreach (var annotation in document.Annotations)
        {
            _writer.WriteStartElement("w", "comment", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteAttributeString("w", "id", null, commentId.ToString());
            
            if (!string.IsNullOrEmpty(annotation.Author))
            {
                _writer.WriteAttributeString("w", "author", null, annotation.Author);
            }
            if (!string.IsNullOrEmpty(annotation.Initials))
            {
                _writer.WriteAttributeString("w", "initials", null, annotation.Initials);
            }
            if (annotation.Date != default && annotation.Date > new System.DateTime(1900, 1, 1))
            {
                _writer.WriteAttributeString("w", "date", null, annotation.Date.ToString("yyyy-MM-ddTHH:mm:ssZ"));
            }

            // Write paragraphs with proper formatting
            if (annotation.Paragraphs.Count > 0)
            {
                foreach (var paragraph in annotation.Paragraphs)
                {
                    _writer.WriteStartElement("w", "p", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                    _writer.WriteStartElement("w", "pPr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                    _writer.WriteStartElement("w", "pStyle", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                    _writer.WriteAttributeString("w", "val", null, "CommentText");
                    _writer.WriteEndElement();
                    _writer.WriteEndElement();
                    
                    foreach (var run in paragraph.Runs)
                    {
                        if (string.IsNullOrEmpty(run.Text)) continue;
                        
                        _writer.WriteStartElement("w", "r", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                        
                        // Write run properties if available
                        WriteCommentRunProperties(run);
                        
                        // write text
                        _writer.WriteStartElement("w", "t", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                        _writer.WriteAttributeString("xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve");
                        _writer.WriteString(run.Text);
                        _writer.WriteEndElement();
                        
                        _writer.WriteEndElement(); // w:r
                    }
                    _writer.WriteEndElement(); // w:p
                }
            }
            else
            {
                // Must have at least one empty paragraph
                _writer.WriteStartElement("w", "p", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteStartElement("w", "pPr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteStartElement("w", "pStyle", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteAttributeString("w", "val", null, "CommentText");
                _writer.WriteEndElement();
                _writer.WriteEndElement();
                _writer.WriteEndElement(); // w:p
            }

            _writer.WriteEndElement(); // w:comment
            
            // Set mapping ID on the annotation model so DocumentWriter knows which ID to use
            annotation.Id = commentId.ToString();
            commentId++;
        }

        _writer.WriteEndElement(); // w:comments
        _writer.WriteEndDocument();
    }

    private void WriteCommentRunProperties(RunModel run)
    {
        var props = run.Properties;
        if (props == null) return;

        bool hasProps = props.IsBold || props.IsItalic || props.IsUnderline ||
                       props.IsStrikeThrough || props.IsSuperscript || props.IsSubscript ||
                       props.IsSmallCaps || props.IsAllCaps || props.Color != 0 ||
                       props.HasRgbColor || props.HighlightColor != 0 ||
                       !string.IsNullOrEmpty(props.FontName) || props.FontSize != 24;

        if (!hasProps) return;

        _writer.WriteStartElement("w", "rPr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

        if (!string.IsNullOrEmpty(props.FontName))
        {
            _writer.WriteStartElement("w", "rFonts", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteAttributeString("w", "ascii", null, props.FontName);
            _writer.WriteAttributeString("w", "hAnsi", null, props.FontName);
            _writer.WriteEndElement();
        }
        if (props.IsBold)
        {
            _writer.WriteStartElement("w", "b", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteEndElement();
        }
        if (props.IsItalic)
        {
            _writer.WriteStartElement("w", "i", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteEndElement();
        }
        if (props.IsUnderline)
        {
            _writer.WriteStartElement("w", "u", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteAttributeString("w", "val", null, "single");
            _writer.WriteEndElement();
        }
        if (props.IsStrikeThrough)
        {
            _writer.WriteStartElement("w", "strike", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteEndElement();
        }
        if (props.IsSmallCaps)
        {
            _writer.WriteStartElement("w", "smallCaps", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteEndElement();
        }
        if (props.IsAllCaps)
        {
            _writer.WriteStartElement("w", "caps", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteEndElement();
        }
        if (props.HasRgbColor)
        {
            _writer.WriteStartElement("w", "color", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteAttributeString("w", "val", null, Utils.ColorHelper.RgbToHex(props.RgbColor));
            _writer.WriteEndElement();
        }
        else if (props.Color > 0)
        {
            _writer.WriteStartElement("w", "color", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteAttributeString("w", "val", null, Utils.ColorHelper.ColorToHex(props.Color));
            _writer.WriteEndElement();
        }
        if (props.FontSize != 24)
        {
            _writer.WriteStartElement("w", "sz", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteAttributeString("w", "val", null, props.FontSize.ToString());
            _writer.WriteEndElement();
        }

        _writer.WriteEndElement(); // w:rPr
    }
}
