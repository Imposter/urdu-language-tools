using System.Collections.Generic;
using System.Linq;

namespace UrduLanguageTools.Extensions
{
    public static class StringExtensions
    {
        private static char[] SplitChars = new char[] { '\u000a', '\u000b', '\u000d' };

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
