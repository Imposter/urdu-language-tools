using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System;

namespace UrduLanguageTools.Extensions
{
    public class LineGroupOptions
    {
        public int? LinesPerGroup { get; set; }
        
        public ParagraphEnding ParagraphEnding { get; set; } = ParagraphEnding.None;
        
        public string ParagraphEndingPattern { get; set; } = string.Empty;
    }
    
    public static class SelectionExtensions
    {
        public static IReadOnlyList<Range> InsertLines(this Selection selection, IReadOnlyList<string> lines, LineGroupOptions options)
        {
            var lineRanges = new List<Range>();
            for (var i = 0; i < lines.Count; i++)
            {
                var isLastLine = i == lines.Count - 1;
                var line = lines[i];
                var start = selection.Start;
                selection.TypeText(line);
                selection.InsertBreak(WdBreakType.wdLineBreak);

                if (options.LinesPerGroup != null && i > 0 && (i + 1) % options.LinesPerGroup == 0)
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
                    // Insert the paragraph ending if specified
                    if (!string.IsNullOrEmpty(options.ParagraphEndingPattern))
                    {
                        // Align text to the center
                        var originalAlignment = selection.ParagraphFormat.Alignment;
                        selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                        selection.TypeText(options.ParagraphEndingPattern);
                        selection.ParagraphFormat.Alignment = originalAlignment;
                    }
                    
                    switch (options.ParagraphEnding)
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
                            throw new ArgumentOutOfRangeException(nameof(options.ParagraphEnding), options.ParagraphEnding, $"Unknown paragraph ending type: {options.ParagraphEnding}");
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