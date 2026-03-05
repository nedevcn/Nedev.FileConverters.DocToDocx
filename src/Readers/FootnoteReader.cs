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

        if (_fib.FcFtn == 0 || _fib.LcbFtn == 0 || _fib.CcpFtn == 0)
            return footnotes;

        try
        {
            footnotes = ReadFootnotesWithOffset();
        }
        catch (Exception ex)
        {
            Logger.Warning("Failed to read footnotes", ex);
        }

        return footnotes;
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
