using System.Collections.Generic;

namespace OpenXmlPowerTools
{
    /// <summary>
    /// Is a handler that does not temper with value in
    /// </summary>
    public class WordprocessingTextDummyHandler : IWordprocessingTextHandler
    {
        /// <summary>
        /// Is a handler that does not temper with value in
        /// </summary>
        public string TransformText(string text, Dictionary<string, string> fontFamily)
        {
            return text;
        }
    }
}