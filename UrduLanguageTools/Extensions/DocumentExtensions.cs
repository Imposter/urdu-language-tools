using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Linq;

namespace UrduLanguageTools
{
    public static class DocumentExtensions
    {
        public static bool TryGetStyle(this Document document, string styleName, out Style style)
        {
            style = null;
            try
            {
                style = document.Styles[styleName];
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static IReadOnlyCollection<string> GetStyleNames(this Document document)
        {
            return document.Styles.Cast<Style>().Select(s => s.NameLocal).ToList();
        }
    }
}
