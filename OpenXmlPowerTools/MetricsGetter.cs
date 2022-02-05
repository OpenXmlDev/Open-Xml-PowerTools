using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Codeuctivity.OpenXmlPowerTools
{
    public class MetricsGetterSettings
    {
        public bool IncludeTextInContentControls { get; set; }
        public bool IncludeXlsxTableCellData { get; set; }
        public bool RetrieveNamespaceList { get; set; }
        public bool RetrieveContentTypeList { get; set; }
    }

    public class MetricsGetter
    {
        private const string UriString = "http://broken-link/";

        public static XElement? GetMetrics(string fileName, MetricsGetterSettings settings)
        {
            var fi = new FileInfo(fileName);
            if (!fi.Exists)
            {
                throw new FileNotFoundException("{0} does not exist.", fi.FullName);
            }

            if (Util.IsWordprocessingML(fi.Extension))
            {
                var wmlDoc = new WmlDocument(fi.FullName, true);
                return GetDocxMetrics(wmlDoc, settings);
            }
            if (Util.IsSpreadsheetML(fi.Extension))
            {
                var smlDoc = new SmlDocument(fi.FullName, true);
                return GetXlsxMetrics(smlDoc, settings);
            }
            if (Util.IsPresentationML(fi.Extension))
            {
                var pmlDoc = new PmlDocument(fi.FullName, true);
                return GetPptxMetrics(pmlDoc, settings);
            }
            return null;
        }

        public static XElement GetDocxMetrics(WmlDocument wmlDoc, MetricsGetterSettings settings)
        {
            try
            {
                using var ms = new MemoryStream();
                ms.Write(wmlDoc.DocumentByteArray, 0, wmlDoc.DocumentByteArray.Length);
                using var document = WordprocessingDocument.Open(ms, true);
                var hasTrackedRevisions = RevisionAccepter.HasTrackedRevisions(document);
                if (hasTrackedRevisions)
                {
                    RevisionAccepter.AcceptRevisions(document);
                }

                var metrics1 = GetWmlMetrics(wmlDoc.FileName, false, document, settings);
                if (hasTrackedRevisions)
                {
                    metrics1.Add(new XElement(H.RevisionTracking, new XAttribute(H.Val, true)));
                }

                return metrics1;
            }
            catch (OpenXmlPowerToolsException e)
            {
                if (e.ToString().Contains("Invalid Hyperlink"))
                {
                    using (var ms = new MemoryStream())
                    {
                        ms.Write(wmlDoc.DocumentByteArray, 0, wmlDoc.DocumentByteArray.Length);
                        UriFixer.FixInvalidUri(ms, brokenUri => FixUri(brokenUri));
                        wmlDoc = new WmlDocument("dummy.docx", ms.ToArray());
                    }
                    using (var ms = new MemoryStream())
                    {
                        ms.Write(wmlDoc.DocumentByteArray, 0, wmlDoc.DocumentByteArray.Length);
                        using var document = WordprocessingDocument.Open(ms, true);
                        var hasTrackedRevisions = RevisionAccepter.HasTrackedRevisions(document);
                        if (hasTrackedRevisions)
                        {
                            RevisionAccepter.AcceptRevisions(document);
                        }

                        var metrics2 = GetWmlMetrics(wmlDoc.FileName, true, document, settings);
                        if (hasTrackedRevisions)
                        {
                            metrics2.Add(new XElement(H.RevisionTracking, new XAttribute(H.Val, true)));
                        }

                        return metrics2;
                    }
                }
            }
            var metrics = new XElement(H.Metrics,
                new XAttribute(H.FileName, wmlDoc.FileName),
                new XAttribute(H.FileType, "WordprocessingML"),
                new XAttribute(H.Error, "Unknown error, metrics not determined"));
            return metrics;
        }

        private static int _getTextWidth(SixLabors.Fonts.FontFamily ff, SixLabors.Fonts.FontStyle fs, decimal sz, string text)
        {
            try
            {
                var font = new SixLabors.Fonts.Font(ff, (float)sz / 2f, fs);
                var size = SixLabors.Fonts.TextMeasurer.Measure(text, new SixLabors.Fonts.RendererOptions(font));

                return (int)size.Width;
            }
            catch
            {
                return 0;
            }
        }

        public static int GetTextWidth(SixLabors.Fonts.FontFamily ff, SixLabors.Fonts.FontStyle fs, decimal sz, string text)
        {
            try
            {
                return _getTextWidth(ff, fs, sz, text);
            }
            catch (ArgumentException)
            {
                try
                {
                    const SixLabors.Fonts.FontStyle fs2 = SixLabors.Fonts.FontStyle.Regular;
                    return _getTextWidth(ff, fs2, sz, text);
                }
                catch (ArgumentException)
                {
                    const SixLabors.Fonts.FontStyle fs2 = SixLabors.Fonts.FontStyle.Bold;
                    try
                    {
                        return _getTextWidth(ff, fs2, sz, text);
                    }
                    catch (ArgumentException)
                    {
                        // if both regular and bold fail, then get metrics for Times New Roman the original FontStyle (in fs)
                        var ff2 = SixLabors.Fonts.SystemFonts.Families.Single(font => font.Name == "Times New Roman");
                        return _getTextWidth(ff2, fs, sz, text);
                    }
                }
            }
            catch (OverflowException)
            {
                // This happened on Azure but interestingly enough not while testing locally.
                return 0;
            }
        }

        private static Uri FixUri(string brokenUri)
        {
            return new Uri(UriString);
        }

        private static XElement GetWmlMetrics(string fileName, bool invalidHyperlink, WordprocessingDocument wDoc, MetricsGetterSettings settings)
        {
            var parts = new XElement(H.Parts,
                wDoc.GetAllParts().Select(part =>
                {
                    return GetMetricsForWmlPart(part, settings);
                }));
            if (!parts.HasElements)
            {
                parts = null;
            }

            var metrics = new XElement(H.Metrics,
                new XAttribute(H.FileName, fileName),
                new XAttribute(H.FileType, "WordprocessingML"),
                GetStyleHierarchy(wDoc),
                GetMiscWmlMetrics(wDoc, invalidHyperlink),
                parts,
                settings.RetrieveNamespaceList ? RetrieveNamespaceList(wDoc) : null,
                settings.RetrieveContentTypeList ? RetrieveContentTypeList(wDoc) : null
                );
            return metrics;
        }

        private static XElement RetrieveContentTypeList(OpenXmlPackage oxPkg)
        {
            var pkg = oxPkg.Package;

            var nonRelationshipParts = pkg.GetParts().Cast<ZipPackagePart>().Where(p => p.ContentType != "application/vnd.openxmlformats-package.relationships+xml");
            var contentTypes = nonRelationshipParts
                .Select(p => p.ContentType)
                .OrderBy(t => t)
                .Distinct();
            var xe = new XElement(H.ContentTypes,
                contentTypes.Select(ct => new XElement(H.ContentType, new XAttribute(H.Val, ct))));
            return xe;
        }

        private static XElement RetrieveNamespaceList(OpenXmlPackage oxPkg)
        {
            var pkg = oxPkg.Package;

            var nonRelationshipParts = pkg.GetParts().Cast<ZipPackagePart>().Where(p => p.ContentType != "application/vnd.openxmlformats-package.relationships+xml");
            var xmlParts = nonRelationshipParts
                .Where(p => p.ContentType.ToLower().EndsWith("xml"));

            var uniqueNamespaces = new HashSet<string>();
            foreach (var xp in xmlParts)
            {
                using var st = xp.GetStream();
                try
                {
                    var xdoc = XDocument.Load(st);
                    var namespaces = xdoc
                        .Descendants()
                        .Attributes()
                        .Where(a => a.IsNamespaceDeclaration)
                        .Select(a => string.Format("{0}|{1}", a.Name.LocalName, a.Value))
                        .OrderBy(t => t)
                        .Distinct()
                        .ToList();
                    foreach (var item in namespaces)
                    {
                        uniqueNamespaces.Add(item);
                    }
                }
                catch (Exception)
                {
                    // if catch exception, forget about it.  Just trying to get a most complete survey possible of all namespaces in all documents. If caught exception, chances are the document is bad anyway.
                    continue;
                }
            }
            var xe = new XElement(H.Namespaces,
                uniqueNamespaces.OrderBy(t => t).Select(n =>
                {
                    var spl = n.Split('|');
                    return new XElement(H.Namespace,
                        new XAttribute(H.NamespacePrefix, spl[0]),
                        new XAttribute(H.NamespaceName, spl[1]));
                }));
            return xe;
        }

        private static List<XElement> GetMiscWmlMetrics(WordprocessingDocument document, bool invalidHyperlink)
        {
            var metrics = new List<XElement>();
            var notes = new List<string>();
            var elementCountDictionary = new Dictionary<XName, int>();

            if (invalidHyperlink)
            {
                metrics.Add(new XElement(H.InvalidHyperlink, new XAttribute(H.Val, invalidHyperlink)));
            }

            var valid = ValidateWordprocessingDocument(document, metrics, notes, elementCountDictionary);
            if (invalidHyperlink)
            {
                valid = false;
            }

            return metrics;
        }

        private static bool ValidateWordprocessingDocument(WordprocessingDocument wDoc, List<XElement> metrics, List<string> notes, Dictionary<XName, int> metricCountDictionary)
        {
            var valid = ValidateAgainstSpecificVersion(wDoc, metrics, DocumentFormat.OpenXml.FileFormatVersions.Office2007, H.SdkValidationError2007);
            valid |= ValidateAgainstSpecificVersion(wDoc, metrics, DocumentFormat.OpenXml.FileFormatVersions.Office2010, H.SdkValidationError2010);
            valid |= ValidateAgainstSpecificVersion(wDoc, metrics, DocumentFormat.OpenXml.FileFormatVersions.Office2013, H.SdkValidationError2013);

            var elementCount = 0;
            var paragraphCount = 0;
            var textCount = 0;
            foreach (var part in wDoc.ContentParts())
            {
                var xDoc = part.GetXDocument();
                foreach (var e in xDoc.Descendants())
                {
                    if (e.Name == W.txbxContent)
                    {
                        IncrementMetric(metricCountDictionary, H.TextBox);
                    }
                    else if (e.Name == W.sdt)
                    {
                        IncrementMetric(metricCountDictionary, H.ContentControl);
                    }
                    else if (e.Name == W.customXml)
                    {
                        IncrementMetric(metricCountDictionary, H.CustomXmlMarkup);
                    }
                    else if (e.Name == W.fldChar)
                    {
                        IncrementMetric(metricCountDictionary, H.ComplexField);
                    }
                    else if (e.Name == W.fldSimple)
                    {
                        IncrementMetric(metricCountDictionary, H.SimpleField);
                    }
                    else if (e.Name == W.altChunk)
                    {
                        IncrementMetric(metricCountDictionary, H.AltChunk);
                    }
                    else if (e.Name == W.tbl)
                    {
                        IncrementMetric(metricCountDictionary, H.Table);
                    }
                    else if (e.Name == W.hyperlink)
                    {
                        IncrementMetric(metricCountDictionary, H.Hyperlink);
                    }
                    else if (e.Name == W.framePr)
                    {
                        IncrementMetric(metricCountDictionary, H.LegacyFrame);
                    }
                    else if (e.Name == W.control)
                    {
                        IncrementMetric(metricCountDictionary, H.ActiveX);
                    }
                    else if (e.Name == W.subDoc)
                    {
                        IncrementMetric(metricCountDictionary, H.SubDocument);
                    }
                    else if (e.Name == VML.imagedata || e.Name == VML.fill || e.Name == VML.stroke || e.Name == A.blip)
                    {
                        var relId = (string)e.Attribute(R.embed);
                        if (relId != null)
                        {
                            ValidateImageExists(part, relId, metricCountDictionary);
                        }

                        relId = (string)e.Attribute(R.pict);
                        if (relId != null)
                        {
                            ValidateImageExists(part, relId, metricCountDictionary);
                        }

                        relId = (string)e.Attribute(R.id);
                        if (relId != null)
                        {
                            ValidateImageExists(part, relId, metricCountDictionary);
                        }
                    }

                    if (part.Uri == wDoc.MainDocumentPart.Uri)
                    {
                        elementCount++;
                        if (e.Name == W.p)
                        {
                            paragraphCount++;
                        }

                        if (e.Name == W.t)
                        {
                            textCount += ((string)e).Length;
                        }
                    }
                }
            }

            foreach (var item in metricCountDictionary)
            {
                metrics.Add(
                    new XElement(item.Key, new XAttribute(H.Val, item.Value)));
            }

            metrics.Add(new XElement(H.ElementCount, new XAttribute(H.Val, elementCount)));
            metrics.Add(new XElement(H.AverageParagraphLength, new XAttribute(H.Val, (int)(textCount / (double)paragraphCount))));

            if (wDoc.GetAllParts().Any(part => part.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
            {
                metrics.Add(new XElement(H.EmbeddedXlsx, new XAttribute(H.Val, true)));
            }

            NumberingFormatListAssembly(wDoc, metrics);

            var wxDoc = wDoc.MainDocumentPart.GetXDocument();

            foreach (var d in wxDoc.Descendants())
            {
                if (d.Name == W.saveThroughXslt)
                {
                    var rid = (string)d.Attribute(R.id);
                    var tempExternalRelationship = wDoc
                        .MainDocumentPart
                        .DocumentSettingsPart
                        .ExternalRelationships
                        .FirstOrDefault(h => h.Id == rid);
                    if (tempExternalRelationship == null)
                    {
                        metrics.Add(new XElement(H.InvalidSaveThroughXslt, new XAttribute(H.Val, true)));
                    }

                    valid = false;
                }
                else if (d.Name == W.trackRevisions)
                {
                    metrics.Add(new XElement(H.TrackRevisionsEnabled, new XAttribute(H.Val, true)));
                }
                else if (d.Name == W.documentProtection)
                {
                    metrics.Add(new XElement(H.DocumentProtection, new XAttribute(H.Val, true)));
                }
            }

            FontAndCharSetAnalysis(wDoc, metrics, notes);

            return valid;
        }

        private static bool ValidateAgainstSpecificVersion(WordprocessingDocument wDoc, List<XElement> metrics, DocumentFormat.OpenXml.FileFormatVersions versionToValidateAgainst, XName versionSpecificMetricName)
        {
            var validator = new OpenXmlValidator(versionToValidateAgainst);
            var errors = validator.Validate(wDoc);
            var valid = errors.Count() == 0;
            if (!valid)
            {
                if (!metrics.Any(e => e.Name == H.SdkValidationError))
                {
                    metrics.Add(new XElement(H.SdkValidationError, new XAttribute(H.Val, true)));
                }

                metrics.Add(new XElement(versionSpecificMetricName, new XAttribute(H.Val, true),
                    errors.Take(3).Select(err =>
                    {
                        var sb = new StringBuilder();
                        if (err.Description.Length > 300)
                        {
                            sb.Append(PtUtils.MakeValidXml(err.Description.Substring(0, 300) + " ... elided ...") + Environment.NewLine);
                        }
                        else
                        {
                            sb.Append(PtUtils.MakeValidXml(err.Description) + Environment.NewLine);
                        }

                        sb.Append("  in part " + PtUtils.MakeValidXml(err.Part.Uri.ToString()) + Environment.NewLine);
                        sb.Append("  at " + PtUtils.MakeValidXml(err.Path.XPath) + Environment.NewLine);
                        return sb.ToString();
                    })));
            }
            return valid;
        }

        private static bool ValidateAgainstSpecificVersion(SpreadsheetDocument sDoc, List<XElement> metrics, DocumentFormat.OpenXml.FileFormatVersions versionToValidateAgainst, XName versionSpecificMetricName)
        {
            var validator = new OpenXmlValidator(versionToValidateAgainst);
            var errors = validator.Validate(sDoc);
            var valid = errors.Count() == 0;
            if (!valid)
            {
                if (!metrics.Any(e => e.Name == H.SdkValidationError))
                {
                    metrics.Add(new XElement(H.SdkValidationError, new XAttribute(H.Val, true)));
                }

                metrics.Add(new XElement(versionSpecificMetricName, new XAttribute(H.Val, true),
                    errors.Take(3).Select(err =>
                    {
                        var sb = new StringBuilder();
                        if (err.Description.Length > 300)
                        {
                            sb.Append(PtUtils.MakeValidXml(err.Description.Substring(0, 300) + " ... elided ...") + Environment.NewLine);
                        }
                        else
                        {
                            sb.Append(PtUtils.MakeValidXml(err.Description) + Environment.NewLine);
                        }

                        sb.Append("  in part " + PtUtils.MakeValidXml(err.Part.Uri.ToString()) + Environment.NewLine);
                        sb.Append("  at " + PtUtils.MakeValidXml(err.Path.XPath) + Environment.NewLine);
                        return sb.ToString();
                    })));
            }
            return valid;
        }

        private static bool ValidateAgainstSpecificVersion(PresentationDocument pDoc, List<XElement> metrics, DocumentFormat.OpenXml.FileFormatVersions versionToValidateAgainst, XName versionSpecificMetricName)
        {
            var validator = new OpenXmlValidator(versionToValidateAgainst);
            var errors = validator.Validate(pDoc);
            var valid = errors.Count() == 0;
            if (!valid)
            {
                if (!metrics.Any(e => e.Name == H.SdkValidationError))
                {
                    metrics.Add(new XElement(H.SdkValidationError, new XAttribute(H.Val, true)));
                }

                metrics.Add(new XElement(versionSpecificMetricName, new XAttribute(H.Val, true),
                    errors.Take(3).Select(err =>
                    {
                        var sb = new StringBuilder();
                        if (err.Description.Length > 300)
                        {
                            sb.Append(PtUtils.MakeValidXml(err.Description.Substring(0, 300) + " ... elided ...") + Environment.NewLine);
                        }
                        else
                        {
                            sb.Append(PtUtils.MakeValidXml(err.Description) + Environment.NewLine);
                        }

                        sb.Append("  in part " + PtUtils.MakeValidXml(err.Part.Uri.ToString()) + Environment.NewLine);
                        sb.Append("  at " + PtUtils.MakeValidXml(err.Path.XPath) + Environment.NewLine);
                        return sb.ToString();
                    })));
            }
            return valid;
        }

        private static void IncrementMetric(Dictionary<XName, int> metricCountDictionary, XName xName)
        {
            if (metricCountDictionary.ContainsKey(xName))
            {
                metricCountDictionary[xName] = metricCountDictionary[xName] + 1;
            }
            else
            {
                metricCountDictionary.Add(xName, 1);
            }
        }

        private static void ValidateImageExists(OpenXmlPart part, string relId, Dictionary<XName, int> metrics)
        {
            var imagePart = part.Parts.FirstOrDefault(ipp => ipp.RelationshipId == relId);
            if (imagePart == null)
            {
                IncrementMetric(metrics, H.ReferenceToNullImage);
            }
        }

        private static void NumberingFormatListAssembly(WordprocessingDocument wDoc, List<XElement> metrics)
        {
            var numFmtList = new List<string>();
            foreach (var part in wDoc.ContentParts())
            {
                var xDoc = part.GetXDocument();
                numFmtList = numFmtList.Concat(xDoc
                    .Descendants(W.p)
                    .Select(p =>
                    {
                        ListItemRetriever.RetrieveListItem(wDoc, p, null);
                        var lif = p.Annotation<ListItemRetriever.ListItemInfo>();
                        if (lif != null && lif.IsListItem && lif.Lvl(ListItemRetriever.GetParagraphLevel(p)) != null)
                        {
                            var numFmtForLevel = (string)lif.Lvl(ListItemRetriever.GetParagraphLevel(p)).Elements(W.numFmt).Attributes(W.val).FirstOrDefault();
                            if (numFmtForLevel == null)
                            {
                                var numFmtElement = lif.Lvl(ListItemRetriever.GetParagraphLevel(p)).Elements(MC.AlternateContent).Elements(MC.Choice).Elements(W.numFmt).FirstOrDefault();
                                if (numFmtElement != null && (string)numFmtElement.Attribute(W.val) == "custom")
                                {
                                    numFmtForLevel = (string)numFmtElement.Attribute(W.format);
                                }
                            }
                            return numFmtForLevel;
                        }
                        return null;
                    })
                    .Where(s => s != null)
                    .Distinct())
                    .ToList();
            }
            if (numFmtList.Any())
            {
                var nfls = numFmtList.StringConcatenate(s => s + ",").TrimEnd(',');
                metrics.Add(new XElement(H.NumberingFormatList, new XAttribute(H.Val, PtUtils.MakeValidXml(nfls))));
            }
        }

        private class FormattingMetrics
        {
            public int RunCount;
            public int RunWithoutRprCount;
            public int ZeroLengthText;
            public int MultiFontRun;

            public int AsciiCharCount;
            public int CSCharCount;
            public int EastAsiaCharCount;
            public int HAnsiCharCount;

            public int AsciiRunCount;
            public int CSRunCount;
            public int EastAsiaRunCount;
            public int HAnsiRunCount;

            public List<string> Languages;

            public FormattingMetrics()
            {
                Languages = new List<string>();
            }
        }

        private static void FontAndCharSetAnalysis(WordprocessingDocument wDoc, List<XElement> metrics, List<string> notes)
        {
            var settings = new FormattingAssemblerSettings
            {
                RemoveStyleNamesFromParagraphAndRunProperties = false,
                ClearStyles = true,
                RestrictToSupportedNumberingFormats = false,
                RestrictToSupportedLanguages = false,
            };
            FormattingAssembler.AssembleFormatting(wDoc, settings);
            var formattingMetrics = new FormattingMetrics();

            foreach (var part in wDoc.ContentParts())
            {
                var xDoc = part.GetXDocument();
                foreach (var run in xDoc.Descendants(W.r))
                {
                    formattingMetrics.RunCount++;
                    AnalyzeRun(run, metrics, notes, formattingMetrics, part.Uri.ToString());
                }
            }

            metrics.Add(new XElement(H.RunCount, new XAttribute(H.Val, formattingMetrics.RunCount)));
            if (formattingMetrics.RunWithoutRprCount > 0)
            {
                metrics.Add(new XElement(H.RunWithoutRprCount, new XAttribute(H.Val, formattingMetrics.RunWithoutRprCount)));
            }

            if (formattingMetrics.ZeroLengthText > 0)
            {
                metrics.Add(new XElement(H.ZeroLengthText, new XAttribute(H.Val, formattingMetrics.ZeroLengthText)));
            }

            if (formattingMetrics.MultiFontRun > 0)
            {
                metrics.Add(new XElement(H.MultiFontRun, new XAttribute(H.Val, formattingMetrics.MultiFontRun)));
            }

            if (formattingMetrics.AsciiCharCount > 0)
            {
                metrics.Add(new XElement(H.AsciiCharCount, new XAttribute(H.Val, formattingMetrics.AsciiCharCount)));
            }

            if (formattingMetrics.CSCharCount > 0)
            {
                metrics.Add(new XElement(H.CSCharCount, new XAttribute(H.Val, formattingMetrics.CSCharCount)));
            }

            if (formattingMetrics.EastAsiaCharCount > 0)
            {
                metrics.Add(new XElement(H.EastAsiaCharCount, new XAttribute(H.Val, formattingMetrics.EastAsiaCharCount)));
            }

            if (formattingMetrics.HAnsiCharCount > 0)
            {
                metrics.Add(new XElement(H.HAnsiCharCount, new XAttribute(H.Val, formattingMetrics.HAnsiCharCount)));
            }

            if (formattingMetrics.AsciiRunCount > 0)
            {
                metrics.Add(new XElement(H.AsciiRunCount, new XAttribute(H.Val, formattingMetrics.AsciiRunCount)));
            }

            if (formattingMetrics.CSRunCount > 0)
            {
                metrics.Add(new XElement(H.CSRunCount, new XAttribute(H.Val, formattingMetrics.CSRunCount)));
            }

            if (formattingMetrics.EastAsiaRunCount > 0)
            {
                metrics.Add(new XElement(H.EastAsiaRunCount, new XAttribute(H.Val, formattingMetrics.EastAsiaRunCount)));
            }

            if (formattingMetrics.HAnsiRunCount > 0)
            {
                metrics.Add(new XElement(H.HAnsiRunCount, new XAttribute(H.Val, formattingMetrics.HAnsiRunCount)));
            }

            if (formattingMetrics.Languages.Any())
            {
                var uls = formattingMetrics.Languages.StringConcatenate(s => s + ",").TrimEnd(',');
                metrics.Add(new XElement(H.Languages, new XAttribute(H.Val, PtUtils.MakeValidXml(uls))));
            }
        }

        private static void AnalyzeRun(XElement run, List<XElement> attList, List<string> notes, FormattingMetrics formattingMetrics, string uri)
        {
            var runText = run.Elements()
                .Where(e => e.Name == W.t || e.Name == W.delText)
                .Select(t => (string)t)
                .StringConcatenate();
            if (runText.Length == 0)
            {
                formattingMetrics.ZeroLengthText++;
                return;
            }
            var rPr = run.Element(W.rPr);
            if (rPr == null)
            {
                formattingMetrics.RunWithoutRprCount++;
                notes.Add(PtUtils.MakeValidXml(string.Format("Error in part {0}: run without rPr at {1}", uri, run.GetXPath())));
                rPr = new XElement(W.rPr);
            }
            var csa = new FormattingAssembler.CharStyleAttributes(null, rPr);
            var fontTypeArray = runText
                .Select(ch => FormattingAssembler.DetermineFontTypeFromCharacter(ch, csa))
                .ToArray();
            var distinctFontTypeArray = fontTypeArray
                .Distinct()
                .ToArray();
            var distinctFonts = distinctFontTypeArray
                .Select(ft =>
                {
                    return GetFontFromFontType(csa, ft);
                })
                .Distinct();
            var languages = distinctFontTypeArray
                .Select(ft =>
                {
                    if (ft == FormattingAssembler.FontType.Ascii)
                    {
                        return csa.LatinLang;
                    }

                    if (ft == FormattingAssembler.FontType.CS)
                    {
                        return csa.BidiLang;
                    }

                    if (ft == FormattingAssembler.FontType.EastAsia)
                    {
                        return csa.EastAsiaLang;
                    }
                    //if (ft == FormattingAssembler.FontType.HAnsi)
                    return csa.LatinLang;
                })
                .Select(l =>
                {
                    if (l == "" || l == null)
                    {
                        return /* "Dflt:" + */ CultureInfo.CurrentCulture.Name;
                    }

                    return l;
                })
                //.Where(l => l != null && l != "")
                .Distinct();
            if (languages.Any(l => !formattingMetrics.Languages.Contains(l)))
            {
                formattingMetrics.Languages = formattingMetrics.Languages.Concat(languages).Distinct().ToList();
            }

            var multiFontRun = distinctFonts.Count() > 1;
            if (multiFontRun)
            {
                formattingMetrics.MultiFontRun++;

                formattingMetrics.AsciiCharCount += fontTypeArray.Where(ft => ft == FormattingAssembler.FontType.Ascii).Count();
                formattingMetrics.CSCharCount += fontTypeArray.Where(ft => ft == FormattingAssembler.FontType.CS).Count();
                formattingMetrics.EastAsiaCharCount += fontTypeArray.Where(ft => ft == FormattingAssembler.FontType.EastAsia).Count();
                formattingMetrics.HAnsiCharCount += fontTypeArray.Where(ft => ft == FormattingAssembler.FontType.HAnsi).Count();
            }
            else
            {
                switch (fontTypeArray[0])
                {
                    case FormattingAssembler.FontType.Ascii:
                        formattingMetrics.AsciiCharCount += runText.Length;
                        formattingMetrics.AsciiRunCount++;
                        break;

                    case FormattingAssembler.FontType.CS:
                        formattingMetrics.CSCharCount += runText.Length;
                        formattingMetrics.CSRunCount++;
                        break;

                    case FormattingAssembler.FontType.EastAsia:
                        formattingMetrics.EastAsiaCharCount += runText.Length;
                        formattingMetrics.EastAsiaRunCount++;
                        break;

                    case FormattingAssembler.FontType.HAnsi:
                        formattingMetrics.HAnsiCharCount += runText.Length;
                        formattingMetrics.HAnsiRunCount++;
                        break;
                }
            }
        }

        private static string? GetFontFromFontType(FormattingAssembler.CharStyleAttributes csa, FormattingAssembler.FontType ft)
        {
            switch (ft)
            {
                case FormattingAssembler.FontType.Ascii:
                    return csa.AsciiFont;

                case FormattingAssembler.FontType.CS:
                    return csa.CsFont;

                case FormattingAssembler.FontType.EastAsia:
                    return csa.EastAsiaFont;

                case FormattingAssembler.FontType.HAnsi:
                    return csa.HAnsiFont;

                default: // dummy
                    return csa.AsciiFont;
            }
        }

        public static XElement GetXlsxMetrics(SmlDocument smlDoc, MetricsGetterSettings settings)
        {
            using var streamDoc = new OpenXmlMemoryStreamDocument(smlDoc);
            using var sDoc = streamDoc.GetSpreadsheetDocument();
            var metrics = new List<XElement>();

            var valid = ValidateAgainstSpecificVersion(sDoc, metrics, DocumentFormat.OpenXml.FileFormatVersions.Office2007, H.SdkValidationError2007);
            valid |= ValidateAgainstSpecificVersion(sDoc, metrics, DocumentFormat.OpenXml.FileFormatVersions.Office2010, H.SdkValidationError2010);
            valid |= ValidateAgainstSpecificVersion(sDoc, metrics, DocumentFormat.OpenXml.FileFormatVersions.Office2013, H.SdkValidationError2013);

            return new XElement(H.Metrics,
                new XAttribute(H.FileName, smlDoc.FileName),
                new XAttribute(H.FileType, "SpreadsheetML"),
                metrics,
                GetTableInfoForWorkbook(sDoc, settings),
                settings.RetrieveNamespaceList ? RetrieveNamespaceList(sDoc) : null,
                settings.RetrieveContentTypeList ? RetrieveContentTypeList(sDoc) : null);
        }

        private static XElement GetTableInfoForWorkbook(SpreadsheetDocument spreadsheet, MetricsGetterSettings settings)
        {
            var workbookPart = spreadsheet.WorkbookPart;
            var xd = workbookPart.GetXDocument();
            var partInformation =
                new XElement(H.Sheets,
                    xd.Root
                    .Element(S.sheets)
                    .Elements(S.sheet)
                    .Select(sh =>
                    {
                        var rid = (string)sh.Attribute(R.id);
                        var sheetName = (string)sh.Attribute("name");
                        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(rid);
                        return GetTableInfoForSheet(spreadsheet, worksheetPart, sheetName, settings);
                    }));
            return partInformation;
        }

        public static XElement? GetTableInfoForSheet(SpreadsheetDocument spreadsheetDocument, WorksheetPart sheetPart, string sheetName,
            MetricsGetterSettings settings)
        {
            var xd = sheetPart.GetXDocument();
            var sheetInformation = new XElement(H.Sheet,
                    new XAttribute(H.Name, sheetName),
                    xd.Root.Elements(S.tableParts).Elements(S.tablePart).Select(tp =>
                    {
                        var rId = (string)tp.Attribute(R.id);
                        var tablePart = (TableDefinitionPart)sheetPart.GetPartById(rId);
                        var txd = tablePart.GetXDocument();
                        var tableName = (string)txd.Root.Attribute("displayName");
                        XElement? tableCellData = null;
                        if (settings.IncludeXlsxTableCellData)
                        {
                            var xlsxTable = spreadsheetDocument.Table(tableName);
                            tableCellData = new XElement(H.TableData,
                                xlsxTable.TableRows()
                                    .Select(row =>
                                    {
                                        var rowElement = new XElement(H.Row,
                                            xlsxTable.TableColumns().Select(col =>
                                            {
                                                var cellElement = new XElement(H.Cell,
                                                    new XAttribute(H.Name, col.Name),
                                                    new XAttribute(H.Val, (string)row[col.Name]));
                                                return cellElement;
                                            }));
                                        return rowElement;
                                    }));
                        }
                        var table = new XElement(H.Table,
                            new XAttribute(H.Name, (string)txd.Root.Attribute("name")),
                            new XAttribute(H.DisplayName, tableName),
                            new XElement(H.Columns,
                                txd.Root.Element(S.tableColumns).Elements(S.tableColumn)
                                .Select(tc => new XElement(H.Column,
                                    new XAttribute(H.Name, (string)tc.Attribute("name"))))),
                                    tableCellData
                            );
                        return table;
                    })
                );
            if (!sheetInformation.HasElements)
            {
                return null;
            }

            return sheetInformation;
        }

        public static XElement GetPptxMetrics(PmlDocument pmlDoc, MetricsGetterSettings settings)
        {
            using var streamDoc = new OpenXmlMemoryStreamDocument(pmlDoc);
            using var pDoc = streamDoc.GetPresentationDocument();
            var metrics = new List<XElement>();

            var valid = ValidateAgainstSpecificVersion(pDoc, metrics, DocumentFormat.OpenXml.FileFormatVersions.Office2007, H.SdkValidationError2007);
            valid |= ValidateAgainstSpecificVersion(pDoc, metrics, DocumentFormat.OpenXml.FileFormatVersions.Office2010, H.SdkValidationError2010);
            valid |= ValidateAgainstSpecificVersion(pDoc, metrics, DocumentFormat.OpenXml.FileFormatVersions.Office2013, H.SdkValidationError2013);
            return new XElement(H.Metrics,
                new XAttribute(H.FileName, pmlDoc.FileName),
                new XAttribute(H.FileType, "PresentationML"),
                metrics,
                settings.RetrieveNamespaceList ? RetrieveNamespaceList(pDoc) : null,
                settings.RetrieveContentTypeList ? RetrieveContentTypeList(pDoc) : null);
        }

        private static object? GetStyleHierarchy(WordprocessingDocument document)
        {
            var stylePart = document.MainDocumentPart.StyleDefinitionsPart;
            if (stylePart == null)
            {
                return null;
            }

            var xd = stylePart.GetXDocument();
            var stylesWithPath = xd.Root
                .Elements(W.style)
                .Select(s =>
                {
                    var styleString = (string)s.Attribute(W.styleId);
                    var thisStyle = s;
                    while (true)
                    {
                        var baseStyle = (string)thisStyle.Elements(W.basedOn).Attributes(W.val).FirstOrDefault();
                        if (baseStyle == null)
                        {
                            break;
                        }

                        styleString = baseStyle + "/" + styleString;
                        thisStyle = xd.Root.Elements(W.style).FirstOrDefault(ts => ts.Attribute(W.styleId).Value == baseStyle);
                        if (thisStyle == null)
                        {
                            break;
                        }
                    }
                    return styleString;
                })
                .OrderBy(n => n)
                .ToList();
            var styleHierarchy = new XElement(H.StyleHierarchy);
            foreach (var item in stylesWithPath)
            {
                var styleChain = item.Split('/');
                var elementToAddTo = styleHierarchy;
                foreach (var inChain in styleChain.PtSkipLast(1))
                {
                    elementToAddTo = elementToAddTo.Elements(H.Style).FirstOrDefault(z => z.Attribute(H.Id).Value == inChain);
                }

                var styleToAdd = styleChain.Last();
                elementToAddTo.Add(
                    new XElement(H.Style,
                        new XAttribute(H.Id, styleChain.Last()),
                        new XAttribute(H.Type, (string)xd.Root.Elements(W.style).First(z => z.Attribute(W.styleId).Value == styleToAdd).Attribute(W.type))));
            }
            return styleHierarchy;
        }

        private static XElement? GetMetricsForWmlPart(OpenXmlPart part, MetricsGetterSettings settings)
        {
            XElement? contentControls = null;
            if (part is MainDocumentPart ||
                part is HeaderPart ||
                part is FooterPart ||
                part is FootnotesPart ||
                part is EndnotesPart)
            {
                var xd = part.GetXDocument();
                contentControls = GetContentControlsTransform(xd.Root, settings) as XElement;
                if (!contentControls.HasElements)
                {
                    contentControls = null;
                }
            }
            var partMetrics = new XElement(H.Part,
                new XAttribute(H.ContentType, part.ContentType),
                new XAttribute(H.Uri, part.Uri.ToString()),
                contentControls);
            if (partMetrics.HasElements)
            {
                return partMetrics;
            }

            return null;
        }

        private static object? GetContentControlsTransform(XNode node, MetricsGetterSettings settings)
        {
            if (node is XElement element)
            {
                if (element == element.Document.Root)
                {
                    return new XElement(H.ContentControls,
                        element.Nodes().Select(n => GetContentControlsTransform(n, settings)));
                }

                if (element.Name == W.sdt)
                {
                    var tag = (string)element.Elements(W.sdtPr).Elements(W.tag).Attributes(W.val).FirstOrDefault();
                    var tagAttr = tag != null ? new XAttribute(H.Tag, tag) : null;

                    var alias = (string)element.Elements(W.sdtPr).Elements(W.alias).Attributes(W.val).FirstOrDefault();
                    var aliasAttr = alias != null ? new XAttribute(H.Alias, alias) : null;

                    var xPathAttr = new XAttribute(H.XPath, element.GetXPath());

                    var isText = element.Elements(W.sdtPr).Elements(W.text).Any();
                    var isBibliography = element.Elements(W.sdtPr).Elements(W.bibliography).Any();
                    var isCitation = element.Elements(W.sdtPr).Elements(W.citation).Any();
                    var isComboBox = element.Elements(W.sdtPr).Elements(W.comboBox).Any();
                    var isDate = element.Elements(W.sdtPr).Elements(W.date).Any();
                    var isDocPartList = element.Elements(W.sdtPr).Elements(W.docPartList).Any();
                    var isDocPartObj = element.Elements(W.sdtPr).Elements(W.docPartObj).Any();
                    var isDropDownList = element.Elements(W.sdtPr).Elements(W.dropDownList).Any();
                    var isEquation = element.Elements(W.sdtPr).Elements(W.equation).Any();
                    var isGroup = element.Elements(W.sdtPr).Elements(W.group).Any();
                    var isPicture = element.Elements(W.sdtPr).Elements(W.picture).Any();
                    var isRichText = element.Elements(W.sdtPr).Elements(W.richText).Any() ||
                        !isText &&
                        !isBibliography &&
                        !isCitation &&
                        !isComboBox &&
                        !isDate &&
                        !isDocPartList &&
                        !isDocPartObj &&
                        !isDropDownList &&
                        !isEquation &&
                        !isGroup &&
                        !isPicture;
                    string? type = null;
                    if (isText)
                    {
                        type = "Text";
                    }

                    if (isBibliography)
                    {
                        type = "Bibliography";
                    }

                    if (isCitation)
                    {
                        type = "Citation";
                    }

                    if (isComboBox)
                    {
                        type = "ComboBox";
                    }

                    if (isDate)
                    {
                        type = "Date";
                    }

                    if (isDocPartList)
                    {
                        type = "DocPartList";
                    }

                    if (isDocPartObj)
                    {
                        type = "DocPartObj";
                    }

                    if (isDropDownList)
                    {
                        type = "DropDownList";
                    }

                    if (isEquation)
                    {
                        type = "Equation";
                    }

                    if (isGroup)
                    {
                        type = "Group";
                    }

                    if (isPicture)
                    {
                        type = "Picture";
                    }

                    if (isRichText)
                    {
                        type = "RichText";
                    }

                    var typeAttr = new XAttribute(H.Type, type);

                    return new XElement(H.ContentControl,
                        typeAttr,
                        tagAttr,
                        aliasAttr,
                        xPathAttr,
                        element.Nodes().Select(n => GetContentControlsTransform(n, settings)));
                }

                return element.Nodes().Select(n => GetContentControlsTransform(n, settings));
            }
            if (settings.IncludeTextInContentControls)
            {
                return node;
            }

            return null;
        }
    }

    public static class H
    {
        public static readonly XName ActiveX = "ActiveX";
        public static readonly XName Alias = "Alias";
        public static readonly XName AltChunk = "AltChunk";
        public static readonly XName Arguments = "Arguments";
        public static readonly XName AsciiCharCount = "AsciiCharCount";
        public static readonly XName AsciiRunCount = "AsciiRunCount";
        public static readonly XName AverageParagraphLength = "AverageParagraphLength";
        public static readonly XName BaselineReport = "BaselineReport";
        public static readonly XName Batch = "Batch";
        public static readonly XName BatchName = "BatchName";
        public static readonly XName BatchSelector = "BatchSelector";
        public static readonly XName CSCharCount = "CSCharCount";
        public static readonly XName CSRunCount = "CSRunCount";
        public static readonly XName Catalog = "Catalog";
        public static readonly XName CatalogList = "CatalogList";
        public static readonly XName CatalogListFile = "CatalogListFile";
        public static readonly XName CaughtException = "CaughtException";
        public static readonly XName Cell = "Cell";
        public static readonly XName Column = "Column";
        public static readonly XName Columns = "Columns";
        public static readonly XName ComplexField = "ComplexField";
        public static readonly XName Computer = "Computer";
        public static readonly XName Computers = "Computers";
        public static readonly XName ContentControl = "ContentControl";
        public static readonly XName ContentControls = "ContentControls";
        public static readonly XName ContentType = "ContentType";
        public static readonly XName ContentTypes = "ContentTypes";
        public static readonly XName CustomXmlMarkup = "CustomXmlMarkup";
        public static readonly XName DLL = "DLL";
        public static readonly XName DefaultDialogValuesFile = "DefaultDialogValuesFile";
        public static readonly XName DefaultValues = "DefaultValues";
        public static readonly XName Dependencies = "Dependencies";
        public static readonly XName DestinationDir = "DestinationDir";
        public static readonly XName Directory = "Directory";
        public static readonly XName DirectoryPattern = "DirectoryPattern";
        public static readonly XName DisplayName = "DisplayName";
        public static readonly XName DoJobQueueName = "DoJobQueueName";
        public static readonly XName Document = "Document";
        public static readonly XName DocumentProtection = "DocumentProtection";
        public static readonly XName DocumentSelector = "DocumentSelector";
        public static readonly XName DocumentType = "DocumentType";
        public static readonly XName Documents = "Documents";
        public static readonly XName EastAsiaCharCount = "EastAsiaCharCount";
        public static readonly XName EastAsiaRunCount = "EastAsiaRunCount";
        public static readonly XName ElementCount = "ElementCount";
        public static readonly XName EmbeddedXlsx = "EmbeddedXlsx";
        public static readonly XName Error = "Error";
        public static readonly XName Exception = "Exception";
        public static readonly XName Exe = "Exe";
        public static readonly XName ExeRoot = "ExeRoot";
        public static readonly XName Extension = "Extension";
        public static readonly XName File = "File";
        public static readonly XName FileLength = "FileLength";
        public static readonly XName FileName = "FileName";
        public static readonly XName FilePattern = "FilePattern";
        public static readonly XName FileType = "FileType";
        public static readonly XName Guid = "Guid";
        public static readonly XName HAnsiCharCount = "HAnsiCharCount";
        public static readonly XName HAnsiRunCount = "HAnsiRunCount";
        public static readonly XName RevisionTracking = "RevisionTracking";
        public static readonly XName Hyperlink = "Hyperlink";
        public static readonly XName IPAddress = "IPAddress";
        public static readonly XName Id = "Id";
        public static readonly XName Invalid = "Invalid";
        public static readonly XName InvalidHyperlink = "InvalidHyperlink";
        public static readonly XName InvalidHyperlinkException = "InvalidHyperlinkException";
        public static readonly XName InvalidSaveThroughXslt = "InvalidSaveThroughXslt";
        public static readonly XName JobComplete = "JobComplete";
        public static readonly XName JobExe = "JobExe";
        public static readonly XName JobName = "JobName";
        public static readonly XName JobSpec = "JobSpec";
        public static readonly XName Languages = "Languages";
        public static readonly XName LegacyFrame = "LegacyFrame";
        public static readonly XName LocalDoJobQueue = "LocalDoJobQueue";
        public static readonly XName MachineName = "MachineName";
        public static readonly XName MaxConcurrentJobs = "MaxConcurrentJobs";
        public static readonly XName MaxDocumentsInJob = "MaxDocumentsInJob";
        public static readonly XName MaxParagraphLength = "MaxParagraphLength";
        public static readonly XName Message = "Message";
        public static readonly XName Metrics = "Metrics";
        public static readonly XName MultiDirectory = "MultiDirectory";
        public static readonly XName MultiFontRun = "MultiFontRun";
        public static readonly XName MultiServerQueue = "MultiServerQueue";
        public static readonly XName Name = "Name";
        public static readonly XName Namespaces = "Namespaces";
        public static readonly XName Namespace = "Namespace";
        public static readonly XName NamespaceName = "NamespaceName";
        public static readonly XName NamespacePrefix = "NamespacePrefix";
        public static readonly XName Note = "Note";
        public static readonly XName NumberingFormatList = "NumberingFormatList";
        public static readonly XName ObjectDisposedException = "ObjectDisposedException";
        public static readonly XName ParagraphCount = "ParagraphCount";
        public static readonly XName Part = "Part";
        public static readonly XName Parts = "Parts";
        public static readonly XName PassedDocuments = "PassedDocuments";
        public static readonly XName Path = "Path";
        public static readonly XName ProduceCatalog = "ProduceCatalog";
        public static readonly XName ReferenceToNullImage = "ReferenceToNullImage";
        public static readonly XName Report = "Report";
        public static readonly XName Root = "Root";
        public static readonly XName RootDirectory = "RootDirectory";
        public static readonly XName Row = "Row";
        public static readonly XName RunCount = "RunCount";
        public static readonly XName RunWithoutRprCount = "RunWithoutRprCount";
        public static readonly XName SdkValidationError = "SdkValidationError";
        public static readonly XName SdkValidationError2007 = "SdkValidationError2007";
        public static readonly XName SdkValidationError2010 = "SdkValidationError2010";
        public static readonly XName SdkValidationError2013 = "SdkValidationError2013";
        public static readonly XName Sheet = "Sheet";
        public static readonly XName Sheets = "Sheets";
        public static readonly XName SimpleField = "SimpleField";
        public static readonly XName Skip = "Skip";
        public static readonly XName SmartTag = "SmartTag";
        public static readonly XName SourceRootDir = "SourceRootDir";
        public static readonly XName SpawnerJobExeLocation = "SpawnerJobExeLocation";
        public static readonly XName SpawnerReady = "SpawnerReady";
        public static readonly XName Style = "Style";
        public static readonly XName StyleHierarchy = "StyleHierarchy";
        public static readonly XName SubDocument = "SubDocument";
        public static readonly XName Table = "Table";
        public static readonly XName TableData = "TableData";
        public static readonly XName Tag = "Tag";
        public static readonly XName Take = "Take";
        public static readonly XName TextBox = "TextBox";
        public static readonly XName TrackRevisionsEnabled = "TrackRevisionsEnabled";
        public static readonly XName Type = "Type";
        public static readonly XName Uri = "Uri";
        public static readonly XName Val = "Val";
        public static readonly XName Valid = "Valid";
        public static readonly XName WindowStyle = "WindowStyle";
        public static readonly XName XPath = "XPath";
        public static readonly XName ZeroLengthText = "ZeroLengthText";
        public static readonly XName custDataLst = "custDataLst";
        public static readonly XName custShowLst = "custShowLst";
        public static readonly XName kinsoku = "kinsoku";
        public static readonly XName modifyVerifier = "modifyVerifier";
        public static readonly XName photoAlbum = "photoAlbum";
    }
}