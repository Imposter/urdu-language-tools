using System;
using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Linq;
using UrduLanguageTools.Extensions;

namespace UrduLanguageTools
{
    public sealed class GhazalOptions
    {
        public Style ParagraphStyle { get; set; }

        public bool AddToTableOfContents { get; set; }

        public ParagraphEnding ParagraphEnding { get; set; }

        public int LinesPerVerse { get; set; }
    }
    
    public static class GhazalExtensions
    {
        public static IReadOnlyList<Range> InsertGhazal(
            this Selection selection,
            IReadOnlyList<string> lines,
            GhazalOptions options)
        {
            // Remove all the existing content and replace it with some content we're going to write
            selection.InsertParagraph();
            selection.set_Style(options.ParagraphStyle);
            selection.ParagraphFormat.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
            selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;

            // Insert line by line
            var lineRanges = new List<Range>();
            for (var i = 0; i < lines.Count; i++)
            {
                var isEndOfVerse = i > 0 && (i + 1) % options.LinesPerVerse == 0;
                var isLastLine = i == lines.Count - 1;
                var line = lines[i];
                var start = selection.Start;
                selection.TypeText(line);
                selection.InsertBreak(WdBreakType.wdLineBreak);
                var end = selection.End;

                if (isEndOfVerse)
                {
                    var emptyLineStart = selection.Start;
                    selection.TypeText(CharCode.BraillePatternBlank.ToString());
                    selection.TypeParagraph();
                    var emptyLineEnd = selection.End;
                    var emptyLineRange = selection.Document.Range(emptyLineStart, emptyLineEnd);
                    emptyLineRange.Font.Size = 1;

                    end = emptyLineEnd;
                }

                if (isLastLine)
                {
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

                    end = selection.End;
                }

                var range = selection.Document.Range(start, end);
                lineRanges.Add(range);
            }

            if (options.AddToTableOfContents)
            {
                // Go to the first line and add the ToC entry
                var firstLine = lineRanges.First();
                firstLine.AddTableOfContentsEntry();
            }
            
            return lineRanges;
        }
    }
}
