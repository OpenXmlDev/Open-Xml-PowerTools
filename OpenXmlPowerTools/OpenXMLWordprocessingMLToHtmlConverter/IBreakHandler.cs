using System.Collections.Generic;
using System.Xml.Linq;

namespace OpenXmlPowerTools.OpenXMLWordprocessingMLToHtmlConverter
{
    public interface IBreakHandler
    {
        IEnumerable<XNode> TransformBreak(XElement element);
    }
}