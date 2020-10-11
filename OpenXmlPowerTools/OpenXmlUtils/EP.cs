using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class EP
    {
        public static readonly XNamespace ep = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";

        public static readonly XName Application = ep + "Application";
        public static readonly XName AppVersion = ep + "AppVersion";
        public static readonly XName Characters = ep + "Characters";
        public static readonly XName CharactersWithSpaces = ep + "CharactersWithSpaces";
        public static readonly XName Company = ep + "Company";
        public static readonly XName DocSecurity = ep + "DocSecurity";
        public static readonly XName HeadingPairs = ep + "HeadingPairs";
        public static readonly XName HiddenSlides = ep + "HiddenSlides";
        public static readonly XName HLinks = ep + "HLinks";
        public static readonly XName HyperlinkBase = ep + "HyperlinkBase";
        public static readonly XName HyperlinksChanged = ep + "HyperlinksChanged";
        public static readonly XName Lines = ep + "Lines";
        public static readonly XName LinksUpToDate = ep + "LinksUpToDate";
        public static readonly XName Manager = ep + "Manager";
        public static readonly XName MMClips = ep + "MMClips";
        public static readonly XName Notes = ep + "Notes";
        public static readonly XName Pages = ep + "Pages";
        public static readonly XName Paragraphs = ep + "Paragraphs";
        public static readonly XName PresentationFormat = ep + "PresentationFormat";
        public static readonly XName Properties = ep + "Properties";
        public static readonly XName ScaleCrop = ep + "ScaleCrop";
        public static readonly XName SharedDoc = ep + "SharedDoc";
        public static readonly XName Slides = ep + "Slides";
        public static readonly XName Template = ep + "Template";
        public static readonly XName TitlesOfParts = ep + "TitlesOfParts";
        public static readonly XName TotalTime = ep + "TotalTime";
        public static readonly XName Words = ep + "Words";
    }
}