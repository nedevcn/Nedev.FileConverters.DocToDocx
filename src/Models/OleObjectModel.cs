using System;
using System.Collections.Generic;

namespace Nedev.DocToDocx.Models;

/// <summary>
/// Represents an OLE Embedded Object extracted from the document.
/// </summary>
public class OleObjectModel
{
    public string ObjectId { get; set; } = string.Empty;
    public string ProgId { get; set; } = string.Empty;
    public byte[] ObjectData { get; set; } = Array.Empty<byte>();
    public int ImageIndex { get; set; } = -1; // Visual representation
    public string? RelationshipId { get; set; }
    public string? MathContent { get; set; } // Converted OMML
}
