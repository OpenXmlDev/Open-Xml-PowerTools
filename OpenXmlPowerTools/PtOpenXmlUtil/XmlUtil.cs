using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class XmlUtil
    {
        public static XAttribute GetXmlSpaceAttribute(string value)
        {
            return value.Length > 0 && (value[0] == ' ' || value[value.Length - 1] == ' ')
                ? new XAttribute(XNamespace.Xml + "space", "preserve")
                : null;
        }

        public static XAttribute GetXmlSpaceAttribute(char value)
        {
            return value == ' ' ? new XAttribute(XNamespace.Xml + "space", "preserve") : null;
        }
    }
}
