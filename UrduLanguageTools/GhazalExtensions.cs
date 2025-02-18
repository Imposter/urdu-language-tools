using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Linq;

namespace UrduLanguageTools
{    
    public static class GhazalExtensions
    {
        public static IReadOnlyList<Range> InsertGhazal(
            this Selection selection,
            IReadOnlyList<string> lines,
            GhazalOptions options)
        {
            var style = selection.Document.Styles[options.ParagraphStyle];

            // Remove all the existing content and replace it with some content we're going to write
            selection.InsertParagraph();
            selection.set_Style(style);
            selection.ParagraphFormat.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;

            // Insert line by line
            var lineRanges = new List<Range>();
            for (var i = 0; i < lines.Count; i++)
            {
                var isEndOfVerse = i > 0 && (i + 1) % options.LinesPerVerse == 0;
                var line = lines[i];
                var start = selection.Start;
                selection.TypeText($"{line}{CharCode.LineBreak}");
                var end = selection.End;

                if (isEndOfVerse)
                {
                    var emptyLineStart = selection.Start;
                    selection.TypeText($"{options.EmptyLineCharacter}{CharCode.ParagraphBreak}");
                    var emptyLineEnd = selection.End;
                    var emptyLineRange = selection.Document.Range(emptyLineStart, emptyLineEnd);
                    emptyLineRange.Font.Size = 1;

                    end = emptyLineEnd;
                }

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
