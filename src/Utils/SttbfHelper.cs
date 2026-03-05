using System.Text;

namespace Nedev.DocToDocx.Utils;

/// <summary>
/// Helper for reading STTBF (String Table with Binary Form) structures from Word documents.
/// </summary>
public static class SttbfHelper
{
    /// <summary>
    /// Reads an STTBF structure from the given reader at the specified offset and length.
    /// Handles both legacy (byte length) and extended (ushort length, Unicode) formats.
    /// </summary>
    public static List<string> ReadSttbf(BinaryReader reader, uint fc, uint lcb)
    {
        var strings = new List<string>();
        if (fc == 0 || lcb == 0) return strings;

        long originalPos = reader.BaseStream.Position;
        try
        {
            reader.BaseStream.Seek(fc, SeekOrigin.Begin);
            
            // Read fExtend (2 bytes)
            ushort fExtend = reader.ReadUInt16();
            bool isUnicode = (fExtend == 0xFFFF);
            
            // Read cData (2 bytes if extended, else it was the first 2 bytes)
            ushort cData = isUnicode ? reader.ReadUInt16() : fExtend;
            
            // cbExtra (2 bytes)
            ushort cbExtra = isUnicode ? reader.ReadUInt16() : (ushort)0;
            if (!isUnicode)
            {
                // In non-extended format, the header is only 2 bytes (cData).
                // We already read it as fExtend.
            }

            for (int i = 0; i < cData; i++)
            {
                if (reader.BaseStream.Position >= fc + lcb) break;

                int cch = isUnicode ? reader.ReadUInt16() : reader.ReadByte();
                if (cch == 0)
                {
                    strings.Add(string.Empty);
                    if (cbExtra > 0) reader.BaseStream.Seek(cbExtra, SeekOrigin.Current);
                    continue;
                }

                byte[] bytes = reader.ReadBytes(isUnicode ? cch * 2 : cch);
                string str = isUnicode 
                    ? Encoding.Unicode.GetString(bytes) 
                    : Encoding.GetEncoding(1252).GetString(bytes); // Fallback to Western European
                
                strings.Add(str.TrimEnd('\0'));

                if (cbExtra > 0)
                {
                    reader.BaseStream.Seek(cbExtra, SeekOrigin.Current);
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Warning("Failed to read STTBF at 0x" + fc.ToString("X"), ex);
        }
        finally
        {
            reader.BaseStream.Seek(originalPos, SeekOrigin.Begin);
        }

        return strings;
    }
}
