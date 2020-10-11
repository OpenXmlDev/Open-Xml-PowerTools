using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class MDSSI
    {
        public static readonly XNamespace mdssi = "http://schemas.openxmlformats.org/package/2006/digital-signature";

        public static readonly XName Format = mdssi + "Format";
        public static readonly XName RelationshipReference = mdssi + "RelationshipReference";
        public static readonly XName SignatureTime = mdssi + "SignatureTime";
        public static readonly XName Value = mdssi + "Value";
    }
}