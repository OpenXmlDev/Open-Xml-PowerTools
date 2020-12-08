using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class XDR14
    {
        public static readonly XNamespace xdr14 = "http://schemas.microsoft.com/office/excel/2010/spreadsheetDrawing";

        public static readonly XName cNvContentPartPr = xdr14 + "cNvContentPartPr";
        public static readonly XName cNvPr = xdr14 + "cNvPr";
        public static readonly XName nvContentPartPr = xdr14 + "nvContentPartPr";
        public static readonly XName nvPr = xdr14 + "nvPr";
        public static readonly XName xfrm = xdr14 + "xfrm";
    }
}