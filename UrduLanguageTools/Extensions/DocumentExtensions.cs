﻿using Microsoft.Office.Interop.Word;

namespace UrduLanguageTools.Extensions
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
    }
}
