using System.Linq;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class Util
    {
        public static readonly string[] WordprocessingExtensions =
        {
            ".docx",
            ".docm",
            ".dotx",
            ".dotm"
        };

        public static readonly string[] SpreadsheetExtensions =
        {
            ".xlsx",
            ".xlsm",
            ".xltx",
            ".xltm",
            ".xlam"
        };

        public static readonly string[] PresentationExtensions =
        {
            ".pptx",
            ".potx",
            ".ppsx",
            ".pptm",
            ".potm",
            ".ppsm",
            ".ppam"
        };

        public static bool IsWordprocessingML(string ext)
        {
            return WordprocessingExtensions.Contains(ext.ToLower());
        }

        public static bool IsSpreadsheetML(string ext)
        {
            return SpreadsheetExtensions.Contains(ext.ToLower());
        }

        public static bool IsPresentationML(string ext)
        {
            return PresentationExtensions.Contains(ext.ToLower());
        }

        public static bool? GetBoolProp(XElement rPr, XName propertyName)
        {
            XElement propAtt = rPr.Element(propertyName);
            if (propAtt == null) return null;

            XAttribute val = propAtt.Attribute(W.val);
            if (val == null) return true;

            string s = ((string) val).ToLower();
            switch (s)
            {
                case "1":
                    return true;
                case "0":
                    return false;
                case "true":
                    return true;
                case "false":
                    return false;
                case "on":
                    return true;
                case "off":
                    return false;
                default:
                    return (bool) propAtt.Attribute(W.val);
            }
        }
    }
}
