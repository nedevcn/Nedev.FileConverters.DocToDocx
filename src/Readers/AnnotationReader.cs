using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Nedev.DocToDocx.Models;
using Nedev.DocToDocx.Utils;

namespace Nedev.DocToDocx.Readers;

public class AnnotationReader
{
    private readonly BinaryReader _tableReader;
    private readonly FibReader _fib;
    private readonly TextReader _textReader;

    public AnnotationReader(BinaryReader tableReader, FibReader fib, TextReader textReader)
    {
        _tableReader = tableReader;
        _fib = fib;
        _textReader = textReader;
    }

    public List<AnnotationModel> ReadAnnotations()
    {
        var annotations = new List<AnnotationModel>();

        if (_fib.LcbPlcfandRef == 0 || _fib.LcbPlcfandTxt == 0 || _tableReader == null)
            return annotations;

        try
        {
            annotations = ReadAnnotationsInternal();
        }
        catch (Exception ex)
        {
            Logger.Warning("Failed to read annotations", ex);
        }

        return annotations;
    }

    private List<AnnotationModel> ReadAnnotationsInternal()
    {
        var annotations = new List<AnnotationModel>();

        // 1. Determine number of comments from PlcfandTxt (array of CPs, no data)
        // PLCF of size LcbPlcfandTxt contains (n + 1) * 4 bytes.
        int n = (int)(_fib.LcbPlcfandTxt / 4) - 1;
        if (n <= 0) return annotations;

        // 2. Read CPs from PlcfandRef (anchors in main document)
        _tableReader.BaseStream.Seek(_fib.FcPlcfandRef, SeekOrigin.Begin);
        int[] anchorCps = new int[n + 1];
        for (int i = 0; i <= n; i++)
        {
            anchorCps[i] = _tableReader.ReadInt32();
        }

        // Calculate element size of ATRDPre10
        int atrdSize = (int)((_fib.LcbPlcfandRef - (n + 1) * 4) / n);

        // Read ATRD structs
        var dttms = new uint[n];
        var authorIndices = new short[n];
        for (int i = 0; i < n; i++)
        {
            long startPos = _tableReader.BaseStream.Position;
            _tableReader.ReadInt16(); // lsr
            _tableReader.ReadInt16(); // irsibcatn
            _tableReader.ReadInt16(); // cchAnBkmk
            dttms[i] = _tableReader.ReadUInt32(); // dttm
            authorIndices[i] = _tableReader.ReadInt16(); // ibst (author index)
            
            _tableReader.BaseStream.Seek(startPos + atrdSize, SeekOrigin.Begin);
        }

        // 3. Read bounds of comment text in ATN space
        _tableReader.BaseStream.Seek(_fib.FcPlcfandTxt, SeekOrigin.Begin);
        int[] txtCps = new int[n + 1];
        for (int i = 0; i <= n; i++)
        {
            txtCps[i] = _tableReader.ReadInt32();
        }

        // Calculate global offset for ATN text (at the end of main text + footnotes + headers)
        int atnGlobalStart = _fib.CcpText + _fib.CcpFtn + _fib.CcpHdd;

        // 4. Extract author names from SttbfAtnMod
        var authors = ReadAuthors();

        // 5. Build models
        for (int i = 0; i < n; i++)
        {
            var annotation = new AnnotationModel
            {
                Id = i.ToString(),
                StartCharacterPosition = anchorCps[i],
                EndCharacterPosition = anchorCps[i], // For now, single-point anchor
                Date = DttmHelper.ParseDttm(dttms[i]),
                Author = (authorIndices[i] >= 0 && authorIndices[i] < authors.Count) ? authors[authorIndices[i]] : "Unknown"
            };

            int len = txtCps[i + 1] - txtCps[i];
            if (len > 0)
            {
                // The comment text is inside the global text space handled by TextReader
                string text = _textReader.GetText(atnGlobalStart + txtCps[i], len);
                
                // Clean control chars (Word puts \r for paragraph breaks)
                text = text.Replace("\r", "\n").TrimEnd('\n', '\x05', '\r', '\0');
                
                var run = new RunModel
                {
                    Text = text,
                    CharacterPosition = atnGlobalStart + txtCps[i],
                    CharacterLength = len
                };
                annotation.Runs.Add(run);

                var paragraph = new ParagraphModel { Index = 0, Type = ParagraphType.Normal };
                paragraph.Runs.Add(run);
                annotation.Paragraphs.Add(paragraph);
            }

            annotations.Add(annotation);
        }

        // 6. Attempt to refine the text range if PlcfAtnbkf/PlcfAtnbkl exist
        RefineAnnotationRanges(annotations);

        return annotations;
    }

    private List<string> ReadAuthors()
    {
        return SttbfHelper.ReadSttbf(_tableReader, _fib.FcSttbfAtnMod, _fib.LcbSttbfAtnMod);
    }

    private void RefineAnnotationRanges(List<AnnotationModel> annotations)
    {
        if (_fib.FcPlcfAtnbkf == 0 || _fib.FcPlcfAtnbkl == 0) return;

        // PlcfAtnbkf: Array of CPs (starts) + Array of short (index)
        // PLCF element size = 2 bytes.
        int nStarts = (int)((_fib.LcbPlcfAtnbkf - 4) / 6);
        if (nStarts > 0)
        {
            _tableReader.BaseStream.Seek(_fib.FcPlcfAtnbkf, SeekOrigin.Begin);
            int[] startCps = new int[nStarts + 1];
            for (int i = 0; i <= nStarts; i++) startCps[i] = _tableReader.ReadInt32();
            
            for (int i = 0; i < nStarts; i++)
            {
                short annotIdx = _tableReader.ReadInt16();
                if (annotIdx >= 0 && annotIdx < annotations.Count)
                {
                    annotations[annotIdx].StartCharacterPosition = startCps[i];
                }
            }
        }

        int nEnds = (int)((_fib.LcbPlcfAtnbkl - 4) / 6);
        if (nEnds > 0)
        {
            _tableReader.BaseStream.Seek(_fib.FcPlcfAtnbkl, SeekOrigin.Begin);
            int[] endCps = new int[nEnds + 1];
            for (int i = 0; i <= nEnds; i++) endCps[i] = _tableReader.ReadInt32();
            
            for (int i = 0; i < nEnds; i++)
            {
                short annotIdx = _tableReader.ReadInt16();
                if (annotIdx >= 0 && annotIdx < annotations.Count)
                {
                    annotations[annotIdx].EndCharacterPosition = endCps[i];
                }
            }
        }
    }

}

