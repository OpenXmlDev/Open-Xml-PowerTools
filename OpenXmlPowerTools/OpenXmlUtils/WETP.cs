using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class WETP
    {
        public static readonly XNamespace wetp = "http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11";
        public static readonly XName extLst = wetp + "extLst";
        public static readonly XName taskpane = wetp + "taskpane";
        public static readonly XName taskpanes = wetp + "taskpanes";
        public static readonly XName web_extension_taskpanes = wetp + "web-extension-taskpanes";
        public static readonly XName webextensionref = wetp + "webextensionref";
    }
}