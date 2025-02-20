using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Word;
using UrduLanguageTools.Extensions;

namespace UrduLanguageTools
{
    public sealed class NazamOptions
    {
        public Style ParagraphStyle { get; set; }
        
        public bool AddToTableOfContents { get; set; }

        public ParagraphEnding ParagraphEnding { get; set; }
    }
    
    public static class NazamExtensions
    {
        public static IReadOnlyList<Range> InsertNazam(
            this Selection selection,
            IReadOnlyList<string> lines,
            NazamOptions options)
        {
            // Remove all the existing content and replace it with some content we're going to write
            selection.InsertParagraph();
            selection.set_Style(options.ParagraphStyle);
            selection.ParagraphFormat.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
            selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

            // Insert line by line
            var lineRanges = new List<Range>();
            for (var i = 0; i < lines.Count; i++)
            {
                var isLastLine = i == lines.Count - 1;
                var line = lines[i];
                var start = selection.Start;
                selection.TypeText(line);
                if (!isLastLine)
                {
                    selection.InsertBreak(WdBreakType.wdLineBreak);
                }
                else
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
                }

                var end = selection.End;
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