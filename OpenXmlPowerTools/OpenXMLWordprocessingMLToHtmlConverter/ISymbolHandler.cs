using System.Collections.Generic;
using System.Xml.Linq;

namespace OpenXmlPowerTools.OpenXMLWordprocessingMLToHtmlConverter
{    /// <summary>
     /// Handler that transforms every symbol from w:sym
     /// </summary>
    public interface ISymbolHandler
    {
        /// <summary>
        /// Returns some kind of changed symbol, that will be used instead of the original in a w:sym element
        /// </summary>
        /// <param name="element"></param>
        /// <param name="fontFamily">fontFamilily of current run</param>
        /// <returns>transformed symbol</returns>
        public XElement TransformSymbol(XElement element, Dictionary<string, string> fontFamily);
    }
}