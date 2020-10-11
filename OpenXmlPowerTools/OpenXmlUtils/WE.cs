using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class WE
    {
        public static readonly XNamespace we = "http://schemas.microsoft.com/office/webextensions/webextension/2010/11";
        public static readonly XName alternateReferences = we + "alternateReferences";
        public static readonly XName binding = we + "binding";
        public static readonly XName bindings = we + "bindings";
        public static readonly XName extLst = we + "extLst";
        public static readonly XName properties = we + "properties";
        public static readonly XName property = we + "property";
        public static readonly XName reference = we + "reference";
        public static readonly XName snapshot = we + "snapshot";
        public static readonly XName web_extension = we + "web-extension";
        public static readonly XName webextension = we + "webextension";
        public static readonly XName webextensionref = we + "webextensionref";
    }
}