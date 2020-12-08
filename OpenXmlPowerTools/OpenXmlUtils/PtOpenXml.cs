using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class PtOpenXml
    {
        public static XNamespace ptOpenXml = "http://powertools.codeplex.com/documentbuilder/2011/insert";
        public static XName Insert = ptOpenXml + "Insert";
        public static XName Id = "Id";

        public static XNamespace pt = "http://powertools.codeplex.com/2011";
        public static XName Uri = pt + "Uri";
        public static XName Unid = pt + "Unid";
        public static XName SHA1Hash = pt + "SHA1Hash";
        public static XName CorrelatedSHA1Hash = pt + "CorrelatedSHA1Hash";
        public static XName StructureSHA1Hash = pt + "StructureSHA1Hash";
        public static XName CorrelationSet = pt + "CorrelationSet";
        public static XName Status = pt + "Status";

        public static XName Level = pt + "Level";
        public static XName IndentLevel = pt + "IndentLevel";
        public static XName ContentType = pt + "ContentType";

        public static XName trPr = pt + "trPr";
        public static XName tcPr = pt + "tcPr";
        public static XName rPr = pt + "rPr";
        public static XName pPr = pt + "pPr";
        public static XName tblPr = pt + "tblPr";
        public static XName style = pt + "style";

        public static XName FontName = pt + "FontName";
        public static XName LanguageType = pt + "LanguageType";
        public static XName AbstractNumId = pt + "AbstractNumId";
        public static XName StyleName = pt + "StyleName";
        public static XName TabWidth = pt + "TabWidth";
        public static XName Leader = pt + "Leader";

        public static XName ListItemRun = pt + "ListItemRun";

        public static XName HtmlToWmlCssWidth = pt + "HtmlToWmlCssWidth";
    }
}