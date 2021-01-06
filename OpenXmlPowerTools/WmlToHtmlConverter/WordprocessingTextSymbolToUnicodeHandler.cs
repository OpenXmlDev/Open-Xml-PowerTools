using System.Collections.Generic;

namespace OpenXmlPowerTools
{
    /// <summary>
    /// Replaces any char of wingdings with the Unicode equivalent
    /// </summary>
    public class WordprocessingTextSymbolToUnicodeHandler : IWordprocessingTextHandler
    {
        private static readonly Dictionary<char, char> WingdingsToUnicode = new Dictionary<char, char>
        {
            // TODO add more
            { '','•' }
        };

        /// <summary>
        /// Replaces any char of wingdings with the Unicode equivalent
        /// </summary>
        public string TransformText(string text, Dictionary<string, string> fontFamily)
        {
            if (fontFamily.TryGetValue("font-family", out var currentFontFamily) && currentFontFamily == "Symbol")
            {
                foreach (var item in WingdingsToUnicode)
                {
                    text = text.Replace(item.Key, item.Value);
                }
            }
            return text;
        }
    }
}