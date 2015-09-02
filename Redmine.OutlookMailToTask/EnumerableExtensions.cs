using System;
using System.Collections.Generic;

namespace Redmine.OutlookMailToTask
{
    public static class EnumerableExtensions
    {
        public static void ForEach<T>(this IEnumerable<T> enumerable, Action<T> action)
        {
            foreach (var e in enumerable)
            {
                action(e);
            }
        }
    }
}
