using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class VT
    {
        public static readonly XNamespace vt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

        public static readonly XName _bool = vt + "bool";
        public static readonly XName filetime = vt + "filetime";
        public static readonly XName i4 = vt + "i4";
        public static readonly XName lpstr = vt + "lpstr";
        public static readonly XName lpwstr = vt + "lpwstr";
        public static readonly XName r8 = vt + "r8";
        public static readonly XName variant = vt + "variant";
        public static readonly XName vector = vt + "vector";
    }
}