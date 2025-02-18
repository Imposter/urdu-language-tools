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
    }
}
