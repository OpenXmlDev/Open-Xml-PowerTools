using System.Linq;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public class Util
    {
        public static readonly string[] WordprocessingExtensions = new[] {
            ".docx",
            ".docm",
            ".dotx",
            ".dotm",
        };

        public static bool IsWordprocessingML(string ext)
        {
            return WordprocessingExtensions.Contains(ext.ToLower());
        }

        public static readonly string[] SpreadsheetExtensions = new[] {
            ".xlsx",
            ".xlsm",
            ".xltx",
            ".xltm",
            ".xlam",
        };

        public static bool IsSpreadsheetML(string ext)
        {
            return SpreadsheetExtensions.Contains(ext.ToLower());
        }

        public static readonly string[] PresentationExtensions = new[] {
            ".pptx",
            ".potx",
            ".ppsx",
            ".pptm",
            ".potm",
            ".ppsm",
            ".ppam",
        };

        public static bool IsPresentationML(string ext)
        {
            return PresentationExtensions.Contains(ext.ToLower());
        }

        public static bool? GetBoolProp(XElement rPr, XName propertyName)
        {
            var propAtt = rPr.Element(propertyName);
            if (propAtt == null)
            {
                return null;
            }

            var val = propAtt.Attribute(W.val);
            if (val == null)
            {
                return true;
            }

            var s = ((string)val).ToLower();
            if (s == "1")
            {
                return true;
            }

            if (s == "0")
            {
                return false;
            }

            if (s == "true")
            {
                return true;
            }

            if (s == "false")
            {
                return false;
            }

            if (s == "on")
            {
                return true;
            }

            if (s == "off")
            {
                return false;
            }

            return (bool)propAtt.Attribute(W.val);
        }
    }
}