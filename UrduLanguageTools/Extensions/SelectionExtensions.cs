using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System;

namespace UrduLanguageTools.Extensions
{
    public static class SelectionExtensions
    {
        public static IReadOnlyList<Range> InsertLines(this Selection selection, IReadOnlyList<string> lines, int? linesPerGroup = null, ParagraphEnding paragraphEnding = ParagraphEnding.None)
        {
            var lineRanges = new List<Range>();
            for (var i = 0; i < lines.Count; i++)
            {
                var isLastLine = i == lines.Count - 1;
                var line = lines[i];
                var start = selection.Start;
                selection.TypeText(line);
                selection.InsertBreak(WdBreakType.wdLineBreak);

                if (linesPerGroup != null && i > 0 && (i + 1) % linesPerGroup == 0)
                {
                    var emptyLineStart = selection.Start;
                    selection.TypeText(CharCode.BraillePatternBlank.ToString());
                    selection.TypeParagraph();
                    var emptyLineEnd = selection.End;
                    var emptyLineRange = selection.Document.Range(emptyLineStart, emptyLineEnd);
                    emptyLineRange.Font.Size = 1;
                }

                if (isLastLine)
                {
                    switch (paragraphEnding)
                    {
                        case ParagraphEnding.Page:
                            selection.InsertBreak(WdBreakType.wdPageBreak);
                            break;
                        case ParagraphEnding.Section:
                            selection.InsertBreak(WdBreakType.wdSectionBreakNextPage);
                            break;
                        case ParagraphEnding.None:
                            break;
                        default:
                            throw new ArgumentOutOfRangeException(nameof(paragraphEnding), paragraphEnding, $"Unknown paragraph ending type: {paragraphEnding}");
                    }
                }

                var end = selection.End;
                var range = selection.Document.Range(start, end);
                lineRanges.Add(range);
            }

            return lineRanges;
        }
    }
}