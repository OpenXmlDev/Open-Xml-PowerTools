using System.Collections.Generic;

namespace OpenXmlPowerTools
{
    internal static class HtmlConverterExtensions
    {
        public static void AddIfMissing(this Dictionary<string, string> style, string propName, string value)
        {
            if (style.ContainsKey(propName))
            {
                return;
            }

            style.Add(propName, value);
        }
    }
}