using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class CUSTPRO
    {
        public static readonly XNamespace custpro = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";

        public static readonly XName Properties = custpro + "Properties";
        public static readonly XName property = custpro + "property";
    }
}