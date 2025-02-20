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

            var lineRanges = selection.InsertLines(lines, paragraphEnding: options.ParagraphEnding);

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