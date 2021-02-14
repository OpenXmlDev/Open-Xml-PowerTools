using System.Collections.Generic;

namespace OpenXmlPowerTools.OpenXMLWordprocessingMLToHtmlConverter
{
    /// <summary>
    /// Is a handler that does not temper with value in W.t elements
    /// </summary>
    public class TextDummyHandler : ITextHandler
    {
        /// <summary>
        /// Is a handler that does not temper with values in W.t elements
        /// </summary>
        public string TransformText(string text, Dictionary<string, string> fontFamily)
        {
            return text;
        }
    }
}