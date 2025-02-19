using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Word;

namespace UrduLanguageTools
{
    public sealed class NasarOptions
    {
        public Style ParagraphStyle { get; set; }
        
        public bool AddToTableOfContents { get; set; }

        public ParagraphEnding ParagraphEnding { get; set; }
    }
    
    public static class NasarExtensions
    {
        public static IReadOnlyList<Range> InsertNasar(
            this Selection selection,
            IReadOnlyList<string> lines,
            NasarOptions options)
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
                var line = lines[i];
                var start = selection.Start;
                selection.TypeText(i == lines.Count - 1
                    ? $"{line}{CharCode.ParagraphBreak}"
                    : $"{line}{CharCode.LineBreak}");

                if (i == lines.Count - 1)
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
                var entryRange = selection.Document.Range(firstLine.Start, firstLine.Start);
                selection.Document.Fields.Add(entryRange, WdFieldType.wdFieldTOCEntry, $"\"{firstLine.Text.Trim()}\"");
            }
            
            return lineRanges;
        }
    }
}