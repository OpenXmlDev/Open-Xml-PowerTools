using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace OpenXmlPowerTools.OpenXMLWordprocessingMLToHtmlConverter
{
    /// <summary>
    /// Default handler that transforms OpenXml breaks into some HTML specific equivalent
    /// </summary>
    public class BreakHandler : IBreakHandler
    {
        /// <summary>
        /// Default handler that transforms breaks into some HTML specific equivalent
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        public IEnumerable<XNode> TransformBreak(XElement element)
        {
            XElement span = default!;
            var tabWidth = (decimal?)element.Attribute(PtOpenXml.TabWidth);
            if (tabWidth != null)
            {
                span = new XElement(Xhtml.span);
                span.AddAnnotation(new Dictionary<string, string>
                {
                    { "margin", string.Format(NumberFormatInfo.InvariantInfo, "0 0 0 {0:0.00}in", tabWidth) },
                    { "padding", "0 0 0 0" }
                });
            }

            var paragraph = element.Ancestors(W.p).FirstOrDefault();

            var isBidi = paragraph != null && paragraph.Elements(W.pPr).Elements(W.bidi).Any(b => b.Attribute(W.val) == null || b.Attribute(W.val).ToBoolean() == true);

            var zeroWidthChar = isBidi ? new XEntity("#x200f") : new XEntity("#x200e");
            return new XNode[] { new XElement(Xhtml.br), zeroWidthChar, span };
        }
    }
}