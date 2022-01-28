using System.Collections.Generic;
using System.Xml.Linq;

namespace Codeuctivity.OpenXMLWordprocessingMLToHtmlConverter
{
    public interface IBreakHandler
    {
        IEnumerable<XNode> TransformBreak(XElement element);
    }
}