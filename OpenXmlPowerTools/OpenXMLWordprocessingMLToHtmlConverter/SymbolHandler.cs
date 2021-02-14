using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OpenXmlPowerTools.OpenXMLWordprocessingMLToHtmlConverter
{
    /// <summary>
    /// Default handler that transforms every symbol into some html encoded font specific char
    /// </summary>
    public class SymbolHandler : ISymbolHandler
    {
        /// <summary>
        /// Default handler that transforms every symbol into some html encoded font specific char
        /// </summary>
        /// <param name="element"></param>
        /// <param name="fontFamily"></param>
        /// <returns></returns>
        public XElement TransformSymbol(XElement element, Dictionary<string, string> fontFamily)
        {
            var cs = (string)element.Attribute(W._char);
            var c = Convert.ToInt32(cs, 16);
            return new XElement(Xhtml.span, new XEntity(string.Format("#{0}", c)));
        }
    }
}