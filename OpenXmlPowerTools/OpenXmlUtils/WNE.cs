using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class WNE
    {
        public static readonly XNamespace wne = "http://schemas.microsoft.com/office/word/2006/wordml";

        public static readonly XName acd = wne + "acd";
        public static readonly XName acdEntry = wne + "acdEntry";
        public static readonly XName acdManifest = wne + "acdManifest";
        public static readonly XName acdName = wne + "acdName";
        public static readonly XName acds = wne + "acds";
        public static readonly XName active = wne + "active";
        public static readonly XName argValue = wne + "argValue";
        public static readonly XName fci = wne + "fci";
        public static readonly XName fciBasedOn = wne + "fciBasedOn";
        public static readonly XName fciIndexBasedOn = wne + "fciIndexBasedOn";
        public static readonly XName fciName = wne + "fciName";
        public static readonly XName hash = wne + "hash";
        public static readonly XName kcmPrimary = wne + "kcmPrimary";
        public static readonly XName kcmSecondary = wne + "kcmSecondary";
        public static readonly XName keymap = wne + "keymap";
        public static readonly XName keymaps = wne + "keymaps";
        public static readonly XName macro = wne + "macro";
        public static readonly XName macroName = wne + "macroName";
        public static readonly XName mask = wne + "mask";
        public static readonly XName recipientData = wne + "recipientData";
        public static readonly XName recipients = wne + "recipients";
        public static readonly XName swArg = wne + "swArg";
        public static readonly XName tcg = wne + "tcg";
        public static readonly XName toolbarData = wne + "toolbarData";
        public static readonly XName toolbars = wne + "toolbars";
        public static readonly XName val = wne + "val";
        public static readonly XName wch = wne + "wch";
    }
}