using System.Text;
using Nedev.DocToDocx.Models;
using Nedev.DocToDocx.Utils;

namespace Nedev.DocToDocx.Readers;

public class FootnoteReader
{
    private readonly FibReader _fib;
    private readonly TextReader _textReader;
    private readonly FkpParser? _fkpParser;
    private readonly StyleSheet? _styles;

    public FootnoteReader(FibReader fib, TextReader textReader,
                          FkpParser? fkpParser = null, StyleSheet? styles = null)
    {
        _fib = fib;
        _textReader = textReader;
        _fkpParser = fkpParser;
        _styles = styles;
    }

    public List<FootnoteModel> ReadFootnotes()
    {
        var footnotes = new List<FootnoteModel>();

            // We normally expect the FIB to advertise the total CP count of the
            // footnote story (CcpFtn) but some buggy files set this field to 0
            // even though a valid PLCF is present and the text lives in the
            // WordDocument stream.  Earlier versions of the library simply
            // returned an empty list in that case; instead we now treat a zero
            // CcpFtn as "unknown" and fall through to the reader logic.  the
            // downstream code already handles the missing count by consulting the
            // piece table when reconstructing the text (see TextReader.SetTextFromPieces).
            if (_fib.FcFtn == 0 || _fib.LcbFtn == 0)
                return footnotes;

            if (_fib.CcpFtn == 0)
            {
                Logger.Info("FIB reports CcpFtn=0; will attempt to read footnotes anyway using PLCF data");
            }
        }
        catch (Exception ex)
        {
            Logger.Warning("Failed to read footnotes", ex);
        }

        return footnotes;
    }

    // helper used by fallback logic to decide if a decoded string still
    // looks unsuitable for use.
    private static bool LooksGarbled(string text, int expectedLength)
    {
        if (string.IsNullOrEmpty(text)) return true;
        if (text.Any(char.IsSurrogate)) return true;
        if (text.Count(c => c == '\0') > expectedLength / 2) return true;
        return false;
    }

    public List<EndnoteModel> ReadEndnotes()
    {
        var endnotes = new List<EndnoteModel>();

        if (_fib.FcEnd == 0 || _fib.LcbEnd == 0 || _fib.CcpEdn == 0)
            return endnotes;

        try
        {
            endnotes = ReadEndnotesWithOffset();
        }
        catch (Exception ex)
        {
            Logger.Warning("Failed to read endnotes", ex);
        }

        return endnotes;
    }

    private List<T> ReadNotes<T>(uint fc, uint lcb, int ccp, int storyOffset, bool isEndnote) where T : NoteModelBase, new()
    {
        var notes = new List<T>();

        // PLCF Structure: (n+1) CPs + n references
        // MS-DOC §2.8.2: Each entry in the PLCF is a FRD (2 bytes).
        // PLC structure: (n+1)*4 + n*2 = 6n + 4
        if (lcb < 10) return notes;

        var tableReader = _textReader.TableReader;
        tableReader.BaseStream.Seek(fc, SeekOrigin.Begin);

        var n = (int)((lcb - 4) / 6);
        if (n <= 0) return notes;

        var cpArray = new int[n + 1];
        for (int i = 0; i <= n; i++) cpArray[i] = tableReader.ReadInt32();

        for (int i = 0; i < n; i++)
        {
            var relStart = cpArray[i];
            var relEnd = cpArray[i + 1];
            var length = relEnd - relStart;

            if (length <= 0) continue;

            var note = new T { Index = i + 1 };
            note.CharacterPosition = relStart;
            note.CharacterLength = length;

            // Extract text from global stream using absolute CP
            var absoluteStartCp = storyOffset + relStart;
            var noteText = _textReader.GetText(absoluteStartCp, length);

            // if the string seems to contain junk (unpaired surrogates, many
            // nulls, etc.) try a second pass using the FKP parser to find the
            // corresponding FC offset and re-decode the bytes directly.  Some
            // documents (including our troublesome sample) store footnote text
            // in the Table stream instead of WordDocument; in those cases the
            // normal reader will return mojibake, so we attempt both streams and
            // pick the result with the higher quality score.
            bool isGarbled = string.IsNullOrEmpty(noteText) ||
                             noteText.Any(char.IsSurrogate) ||
                             noteText.Count(c => c == '\0') > length / 2;
            if (isGarbled && _fkpParser != null)
            {
                var fcOffset = _fkpParser.CpToFc(absoluteStartCp);
                if (fcOffset.HasValue)
                {
                    Logger.Info($"FKP fallback cp={absoluteStartCp} length={length} -> fc={fcOffset.Value}");
                    // first try the usual WordDocument stream
                    var alt = _textReader.DecodeRangeFromFc(fcOffset.Value, length, _fib.Lid);
                    if (!string.IsNullOrEmpty(alt) && !LooksGarbled(alt, length))
                    {
                        noteText = alt;
                    }
                    else
                    {
                        // try Table stream decoder
                        var alt2 = _textReader.DecodeRangeFromTableFc(fcOffset.Value, length, _fib.Lid);
                        if (!string.IsNullOrEmpty(alt2) && !LooksGarbled(alt2, length))
                        {
                            Logger.Info("Successfully decoded using table stream fallback");
                            noteText = alt2;
                        }
                    }
                }
            }

            if (string.IsNullOrEmpty(noteText) && length > 0)
            {
                // debug aid: log positions that produced no text; useful when
                // investigating missing footnote content in sample documents.
                Logger.Warning($"Empty footnote text (relStart={relStart} length={length} absoluteCp={absoluteStartCp} storyOffset={storyOffset})");
            }

            if (!string.IsNullOrEmpty(noteText))
            {
                // Try to get actual CHP formatting from FkpParser
                RunProperties? runProps = null;
                if (_fkpParser != null && _styles != null)
                {
                    try
                    {
                        var chp = _fkpParser.GetChpAtCp(absoluteStartCp);
                        if (chp != null)
                        {
                            runProps = _fkpParser.ConvertToRunProperties(chp, _styles);
                        }
                    }
                    catch
                    {
                        // Fall through to default
                    }
                }

                var run = new RunModel
                {
                    Text = noteText,
                    CharacterPosition = relStart,
                    CharacterLength = noteText.Length,
                    Properties = runProps ?? new RunProperties()
                };
                note.Runs.Add(run);

                var paragraph = new ParagraphModel
                {
                    Index = 0,
                    Type = ParagraphType.Normal
                };
                paragraph.Runs.Add(run);
                note.Paragraphs.Add(paragraph);
            }

            notes.Add(note);
        }

        return notes;
    }

    public List<FootnoteModel> ReadFootnotesWithOffset()
    {
        // Story offset: Body
        int footnoteStoryOffset = _fib.CcpText;
        return ReadNotes<FootnoteModel>(_fib.FcFtn, _fib.LcbFtn, _fib.CcpFtn, footnoteStoryOffset, false);
    }

    public List<EndnoteModel> ReadEndnotesWithOffset()
    {
        // Story offset: Body + Footnotes + Headers + Annotations
        int endnoteStoryOffset = _fib.CcpText + _fib.CcpFtn + _fib.CcpHdd + _fib.CcpAtn;
        var notes = ReadNotes<EndnoteModel>(_fib.FcEnd, _fib.LcbEnd, _fib.CcpEdn, endnoteStoryOffset, true);
        return notes.Cast<EndnoteModel>().ToList();
    }
}
