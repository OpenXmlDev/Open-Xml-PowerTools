using System.Collections.Generic;

namespace OpenXmlPowerTools.OpenXMLWordprocessingMLToHtmlConverter
{
    /// <summary>
    /// Handler that transforms every text in w:t
    /// </summary>
    public interface ITextHandler
    {
        /// <summary>
        /// Returns some kind of changed text, that will be used instead of the original in w:t elements
        /// </summary>
        /// <param name="text"></param>
        /// <param name="fontFamily">fontFamilily of current run</param>
        /// <returns>transformed text</returns>
        public string TransformText(string text, Dictionary<string, string> fontFamily);
    }
}