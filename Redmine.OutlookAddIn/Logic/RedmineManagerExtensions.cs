using System.Collections.Generic;
using System.Collections.Specialized;

namespace Redmine.Net.Api.Extensions
{
    public static class RedmineManagerExtensions
    {
        public static List<T> GetAllObjectList<T>(this RedmineManager manager, NameValueCollection parameters = null) where T : class, new()
        {
            if (parameters == null)
            {
                parameters = new NameValueCollection();
            }

            IList<T> redmineAllRecordsResult;
            List<T> allRecords = new List<T>();

            int limit = 100;
            int offset = 0;
            do
            {
                parameters["offset"] = offset.ToString();
                parameters["limit"] = limit.ToString();

                redmineAllRecordsResult = manager.GetObjectList<T>(parameters);
                allRecords.AddRange(redmineAllRecordsResult);

                offset += redmineAllRecordsResult.Count;
            }
            while (redmineAllRecordsResult.Count == limit);

            return allRecords;
        }
    }
}
