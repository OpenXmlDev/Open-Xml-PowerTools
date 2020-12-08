using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class DS
    {
        public static readonly XNamespace ds = "http://schemas.openxmlformats.org/officeDocument/2006/customXml";

        public static readonly XName datastoreItem = ds + "datastoreItem";
        public static readonly XName itemID = ds + "itemID";
        public static readonly XName schemaRef = ds + "schemaRef";
        public static readonly XName schemaRefs = ds + "schemaRefs";
        public static readonly XName uri = ds + "uri";
    }
}