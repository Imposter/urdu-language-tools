using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using UrduLanguageTools.Extensions;

namespace UrduLanguageTools
{
    public static class RangeExtensions
    {
        public static IReadOnlyList<string> GetLines(this Range range, params char[] newlineChars)
        {
            var includeHiddenText = range.TextRetrievalMode.IncludeHiddenText;
            try
            {
                range.TextRetrievalMode.IncludeHiddenText = false;
                return range.Text.GetLines(newlineChars);
            }
            finally
            {
                range.TextRetrievalMode.IncludeHiddenText = includeHiddenText;
            }
        }
    }
}
