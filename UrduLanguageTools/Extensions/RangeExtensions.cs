using Microsoft.Office.Interop.Word;

namespace UrduLanguageTools.Extensions
{
    public static class RangeExtensions
    {
        public static string GetText(this Range range)
        {
            var includeHiddenText = range.TextRetrievalMode.IncludeHiddenText;
            try
            {
                range.TextRetrievalMode.IncludeHiddenText = false;
                return range.Text;
            }
            finally
            {
                range.TextRetrievalMode.IncludeHiddenText = includeHiddenText;
            }
        }

        public static void AddTableOfContentsEntry(this Range range, string text = "", int level = 0)
        {
            range.Document.Fields.Add(
                range.Document.Range(range.Start, range.Start),
                WdFieldType.wdFieldTOCEntry, 
                $"\"{(string.IsNullOrEmpty(text) ? range.Text.Trim() : text)}\"");
        }
    }
}
