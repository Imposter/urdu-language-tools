using System.Collections;
using System.Collections.Generic;

namespace UrduLanguageTools
{
    public static class EnumerableExtensions
    {
        public static IEnumerable<T> Cast<T>(this IEnumerable enumerable)
        {
            foreach (T item in enumerable)
            {
                yield return item;
            }
        }
    }
}
