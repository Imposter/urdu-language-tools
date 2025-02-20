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

            var lineRanges = selection.InsertLines(lines, options.LinesPerVerse, options.ParagraphEnding);

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
