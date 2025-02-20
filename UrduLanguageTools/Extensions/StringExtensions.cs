using System.Collections.Generic;
using System.Linq;

namespace UrduLanguageTools.Extensions
{
    public static class StringExtensions
    {
        private static readonly char[] SplitChars =
        {
            '\u000a', // Line feed
            '\u000b', // Vertical tab
            '\u000d'  // Carriage return
        };

        public static IReadOnlyList<string> GetLines(this string text, params char[] newlineChars)
        {
            return text
                   .Split(newlineChars.Length != 0 ? SplitChars.Concat(newlineChars).ToArray() : SplitChars)
                   .Select(s => s.Trim())
                   .Where(s => s.Length != 0)
                   .ToList();
        }

        public static string RemoveMultipleSpaces(this string text)
        {
            string modifiedText;

            if (string.IsNullOrEmpty(text))
                return string.Empty;

            do
            {
                modifiedText = text.Replace("  ", " ");
                if (modifiedText == text)
                    break;
                text = modifiedText;
            } while (true);

            return modifiedText;
        }
    }
}
