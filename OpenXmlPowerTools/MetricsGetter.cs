// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using JetBrains.Annotations;

#pragma warning disable CS0414

namespace OpenXmlPowerTools
{
    [PublicAPI]
    public static class MetricsGetter
    {
        private static Lazy<Graphics> Graphics { get; } = new(() =>
        {
            Image image = new Bitmap(1, 1);
            return System.Drawing.Graphics.FromImage(image);
        });

        public static XElement GetMetrics(string fileName, MetricsGetterSettings settings)
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

                using WordprocessingDocument document = WordprocessingDocument.Open(ms, true);

                bool hasTrackedRevisions = RevisionAccepter.HasTrackedRevisions(document);

                if (hasTrackedRevisions)
                {
                    RevisionAccepter.AcceptRevisions(document);
                }

                XElement metrics1 = GetWmlMetrics(wmlDoc.FileName, false, document, settings);

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
#if !NET35
                        UriFixer.FixInvalidUri(ms, FixUri);
#endif
                        wmlDoc = new WmlDocument("dummy.docx", ms.ToArray());
                    }

                    using (var ms = new MemoryStream())
                    {
                        ms.Write(wmlDoc.DocumentByteArray, 0, wmlDoc.DocumentByteArray.Length);

                        using WordprocessingDocument document = WordprocessingDocument.Open(ms, true);

                        bool hasTrackedRevisions = RevisionAccepter.HasTrackedRevisions(document);

                        if (hasTrackedRevisions)
                        {
                            RevisionAccepter.AcceptRevisions(document);
                        }

                        XElement metrics2 = GetWmlMetrics(wmlDoc.FileName, true, document, settings);

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

        private static Uri FixUri(string brokenUri)
        {
            return new Uri("http://broken-link/");
        }

        private static XElement GetWmlMetrics(
            string fileName,
            bool invalidHyperlink,
            WordprocessingDocument wDoc,
            MetricsGetterSettings settings)
        {
            var parts = new XElement(H.Parts, wDoc.GetAllParts().Select(part => GetMetricsForWmlPart(part, settings)));

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
                settings.RetrieveContentTypeList ? RetrieveContentTypeList(wDoc) : null);

            return metrics;
        }

        private static XElement RetrieveNamespaceList(OpenXmlPackage oxPkg)
        {
            Package pkg = oxPkg.Package;

            IEnumerable<ZipPackagePart> nonRelationshipParts = pkg.GetParts()
                .Cast<ZipPackagePart>()
                .Where(p => p.ContentType != "application/vnd.openxmlformats-package.relationships+xml");

            IEnumerable<ZipPackagePart> xmlParts = nonRelationshipParts
                .Where(p => p.ContentType
                    .ToLower(CultureInfo.InvariantCulture)
                    .EndsWith("xml", true, CultureInfo.InvariantCulture));

            var uniqueNamespaces = new HashSet<string>();

            foreach (ZipPackagePart xp in xmlParts)
            {
                using Stream st = xp.GetStream();

                try
                {
                    XDocument xdoc = XDocument.Load(st);

                    List<string> namespaces = xdoc
                        .Descendants()
                        .Attributes()
                        .Where(a => a.IsNamespaceDeclaration)
                        .Select(a => $"{a.Name.LocalName}|{a.Value}")
                        .OrderBy(t => t)
                        .Distinct()
                        .ToList();

                    foreach (string item in namespaces)
                    {
                        uniqueNamespaces.Add(item);
                    }
                }
                catch
                {
                    // If caught exception, forget about it. Just trying to get a most complete survey
                    // possible of all namespaces in all documents.
                    // If caught exception, chances are the document is bad anyway.
                }
            }

            var xe = new XElement(H.Namespaces,
                uniqueNamespaces.OrderBy(t => t)
                    .Select(n =>
                    {
                        string[] spl = n.Split('|');

                        return new XElement(H.Namespace,
                            new XAttribute(H.NamespacePrefix, spl[0]),
                            new XAttribute(H.NamespaceName, spl[1]));
                    }));

            return xe;
        }

        private static XElement RetrieveContentTypeList(OpenXmlPackage oxPkg)
        {
            Package pkg = oxPkg.Package;

            IEnumerable<ZipPackagePart> nonRelationshipParts = pkg.GetParts()
                .Cast<ZipPackagePart>()
                .Where(p => p.ContentType != "application/vnd.openxmlformats-package.relationships+xml");

            IEnumerable<string> contentTypes = nonRelationshipParts
                .Select(p => p.ContentType)
                .OrderBy(t => t)
                .Distinct();

            var xe = new XElement(H.ContentTypes,
                contentTypes.Select(ct => new XElement(H.ContentType, new XAttribute(H.Val, ct))));

            return xe;
        }

        private static List<XElement> GetMiscWmlMetrics(WordprocessingDocument document, bool invalidHyperlink)
        {
            var metrics = new List<XElement>();
            var notes = new List<string>();
            var elementCountDictionary = new Dictionary<XName, int>();

            if (invalidHyperlink)
            {
                metrics.Add(new XElement(H.InvalidHyperlink, new XAttribute(H.Val, true)));
            }

            ValidateWordprocessingDocument(document, metrics, notes, elementCountDictionary);

            return metrics;
        }

        private static bool ValidateWordprocessingDocument(
            WordprocessingDocument wDoc,
            List<XElement> metrics,
            List<string> notes,
            Dictionary<XName, int> metricCountDictionary)
        {
            bool valid = ValidateAgainstSpecificVersion(wDoc, metrics, DocumentFormat.OpenXml.FileFormatVersions.Office2007,
                H.SdkValidationError2007);

            valid |= ValidateAgainstSpecificVersion(wDoc, metrics, DocumentFormat.OpenXml.FileFormatVersions.Office2010,
                H.SdkValidationError2010);
#if !NET35
            valid |= ValidateAgainstSpecificVersion(wDoc, metrics, DocumentFormat.OpenXml.FileFormatVersions.Office2013,
                H.SdkValidationError2013);
#endif

            var elementCount = 0;
            var paragraphCount = 0;
            var textCount = 0;

            foreach (OpenXmlPart part in wDoc.ContentParts())
            {
                XDocument xDoc = part.GetXDocument();

                foreach (XElement e in xDoc.Descendants())
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
                        var relId = (string) e.Attribute(R.embed);

                        if (relId != null)
                        {
                            ValidateImageExists(part, relId, metricCountDictionary);
                        }

                        relId = (string) e.Attribute(R.pict);

                        if (relId != null)
                        {
                            ValidateImageExists(part, relId, metricCountDictionary);
                        }

                        relId = (string) e.Attribute(R.id);

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
                            textCount += ((string) e).Length;
                        }
                    }
                }
            }

            metrics.AddRange(metricCountDictionary.Select(item => new XElement(item.Key, new XAttribute(H.Val, item.Value))));
            metrics.Add(new XElement(H.ElementCount, new XAttribute(H.Val, elementCount)));

            metrics.Add(
                new XElement(H.AverageParagraphLength, new XAttribute(H.Val, (int) (textCount / (double) paragraphCount))));

            if (wDoc.GetAllParts()
                .Any(part => part.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
            {
                metrics.Add(new XElement(H.EmbeddedXlsx, new XAttribute(H.Val, true)));
            }

            NumberingFormatListAssembly(wDoc, metrics);

            XDocument wxDoc = wDoc.MainDocumentPart.GetXDocument();

            foreach (XElement d in wxDoc.Descendants())
            {
                if (d.Name == W.saveThroughXslt)
                {
                    var rid = (string) d.Attribute(R.id);

                    ExternalRelationship tempExternalRelationship = wDoc
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

        private static bool ValidateAgainstSpecificVersion(
            WordprocessingDocument wDoc,
            List<XElement> metrics,
            DocumentFormat.OpenXml.FileFormatVersions versionToValidateAgainst,
            XName versionSpecificMetricName)
        {
            var validator = new OpenXmlValidator(versionToValidateAgainst);
            IEnumerable<ValidationErrorInfo> errors = validator.Validate(wDoc).ToList();
            bool valid = !errors.Any();

            if (valid)
            {
                return true;
            }

            if (metrics.All(e => e.Name != H.SdkValidationError))
            {
                metrics.Add(new XElement(H.SdkValidationError, new XAttribute(H.Val, true)));
            }

            AddValidationErrorInfos(metrics, versionSpecificMetricName, errors);

            return false;
        }

        private static bool ValidateAgainstSpecificVersion(
            SpreadsheetDocument sDoc,
            List<XElement> metrics,
            DocumentFormat.OpenXml.FileFormatVersions versionToValidateAgainst,
            XName versionSpecificMetricName)
        {
            var validator = new OpenXmlValidator(versionToValidateAgainst);
            IEnumerable<ValidationErrorInfo> errors = validator.Validate(sDoc).ToList();
            bool valid = !errors.Any();

            if (valid)
            {
                return true;
            }

            if (metrics.All(e => e.Name != H.SdkValidationError))
            {
                metrics.Add(new XElement(H.SdkValidationError, new XAttribute(H.Val, true)));
            }

            AddValidationErrorInfos(metrics, versionSpecificMetricName, errors);

            return false;
        }

        private static bool ValidateAgainstSpecificVersion(
            PresentationDocument pDoc,
            List<XElement> metrics,
            DocumentFormat.OpenXml.FileFormatVersions versionToValidateAgainst,
            XName versionSpecificMetricName)
        {
            var validator = new OpenXmlValidator(versionToValidateAgainst);
            IEnumerable<ValidationErrorInfo> errors = validator.Validate(pDoc).ToList();
            bool valid = !errors.Any();

            if (valid)
            {
                return true;
            }

            if (metrics.All(e => e.Name != H.SdkValidationError))
            {
                metrics.Add(new XElement(H.SdkValidationError, new XAttribute(H.Val, true)));
            }

            AddValidationErrorInfos(metrics, versionSpecificMetricName, errors);

            return false;
        }

        private static void AddValidationErrorInfos(List<XElement> metrics, XName versionSpecificMetricName, IEnumerable<ValidationErrorInfo> errors)
        {
            metrics.Add(new XElement(versionSpecificMetricName, new XAttribute(H.Val, true),
                errors.Take(3)
                    .Select(err =>
                    {
                        var sb = new StringBuilder();

                        sb.AppendLine(err.Description.Length > 300
                            ? PtUtils.MakeValidXml(err.Description.Substring(0, 300) + " ... elided ...")
                            : PtUtils.MakeValidXml(err.Description));

                        sb.AppendLine("  in part " + PtUtils.MakeValidXml(err.Part!.Uri.ToString()));
                        sb.AppendLine("  at " + PtUtils.MakeValidXml(err.Path!.XPath));

                        return sb.ToString();
                    })));
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
            IdPartPair imagePart = part.Parts.FirstOrDefault(ipp => ipp.RelationshipId == relId);

            if (imagePart == null)
            {
                IncrementMetric(metrics, H.ReferenceToNullImage);
            }
        }

        private static void NumberingFormatListAssembly(WordprocessingDocument wDoc, List<XElement> metrics)
        {
            var numFmtList = new List<string>();

            foreach (OpenXmlPart part in wDoc.ContentParts())
            {
                XDocument xDoc = part.GetXDocument();

                numFmtList = numFmtList.Concat(xDoc
                        .Descendants(W.p)
                        .Select(p =>
                        {
                            ListItemRetriever.RetrieveListItem(wDoc, p, null);
                            var lif = p.Annotation<ListItemRetriever.ListItemInfo>();

                            if (lif is { IsListItem: true } && lif.Lvl(ListItemRetriever.GetParagraphLevel(p)) != null)
                            {
                                var numFmtForLevel = (string) lif.Lvl(ListItemRetriever.GetParagraphLevel(p))
                                    .Elements(W.numFmt)
                                    .Attributes(W.val)
                                    .FirstOrDefault();

                                if (numFmtForLevel == null)
                                {
                                    XElement numFmtElement = lif.Lvl(ListItemRetriever.GetParagraphLevel(p))
                                        .Elements(MC.AlternateContent)
                                        .Elements(MC.Choice)
                                        .Elements(W.numFmt)
                                        .FirstOrDefault();

                                    if (numFmtElement != null && (string) numFmtElement.Attribute(W.val) == "custom")
                                    {
                                        numFmtForLevel = (string) numFmtElement.Attribute(W.format);
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
                string nfls = numFmtList.StringConcatenate(s => s + ",").TrimEnd(',');
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
            public int CsCharCount;
            public int EastAsiaCharCount;
            public int HAnsiCharCount;

            public int AsciiRunCount;
            public int CsRunCount;
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

            foreach (OpenXmlPart part in wDoc.ContentParts())
            {
                XDocument xDoc = part.GetXDocument();

                foreach (XElement run in xDoc.Descendants(W.r))
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

            if (formattingMetrics.CsCharCount > 0)
            {
                metrics.Add(new XElement(H.CSCharCount, new XAttribute(H.Val, formattingMetrics.CsCharCount)));
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

            if (formattingMetrics.CsRunCount > 0)
            {
                metrics.Add(new XElement(H.CSRunCount, new XAttribute(H.Val, formattingMetrics.CsRunCount)));
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
                string uls = formattingMetrics.Languages.StringConcatenate(s => s + ",").TrimEnd(',');
                metrics.Add(new XElement(H.Languages, new XAttribute(H.Val, PtUtils.MakeValidXml(uls))));
            }
        }

        private static void AnalyzeRun(
            XElement run,
            List<XElement> attList,
            List<string> notes,
            FormattingMetrics formattingMetrics,
            string uri)
        {
            string runText = run.Elements()
                .Where(e => e.Name == W.t || e.Name == W.delText)
                .Select(t => (string) t)
                .StringConcatenate();

            if (runText.Length == 0)
            {
                formattingMetrics.ZeroLengthText++;
                return;
            }

            XElement rPr = run.Element(W.rPr);

            if (rPr == null)
            {
                formattingMetrics.RunWithoutRprCount++;
                notes.Add(PtUtils.MakeValidXml($"Error in part {uri}: run without rPr at {run.GetXPath()}"));
                rPr = new XElement(W.rPr);
            }

            var csa = new FormattingAssembler.CharStyleAttributes(null, rPr);

            FormattingAssembler.FontType[] fontTypeArray = runText
                .Select(ch => FormattingAssembler.DetermineFontTypeFromCharacter(ch, csa))
                .ToArray();

            FormattingAssembler.FontType[] distinctFontTypeArray = fontTypeArray
                .Distinct()
                .ToArray();

            IEnumerable<string> distinctFonts = distinctFontTypeArray
                .Select(ft => GetFontFromFontType(csa, ft))
                .Distinct();

            List<string> languages = distinctFontTypeArray
                .Select(ft => ft switch
                {
                    FormattingAssembler.FontType.Ascii => csa.LatinLang,
                    FormattingAssembler.FontType.CS => csa.BidiLang,
                    FormattingAssembler.FontType.EastAsia => csa.EastAsiaLang,
                    _ => csa.LatinLang
                })
                .Select(lang => string.IsNullOrEmpty(lang) ? CultureInfo.CurrentCulture.Name : lang)
                .Distinct()
                .ToList();

            if (languages.Any(lang => !formattingMetrics.Languages.Contains(lang)))
            {
                formattingMetrics.Languages = formattingMetrics.Languages.Concat(languages).Distinct().ToList();
            }

            bool multiFontRun = distinctFonts.Count() > 1;

            if (multiFontRun)
            {
                formattingMetrics.MultiFontRun++;
                formattingMetrics.AsciiCharCount += fontTypeArray.Count(ft => ft == FormattingAssembler.FontType.Ascii);
                formattingMetrics.CsCharCount += fontTypeArray.Count(ft => ft == FormattingAssembler.FontType.CS);
                formattingMetrics.EastAsiaCharCount += fontTypeArray.Count(ft => ft == FormattingAssembler.FontType.EastAsia);
                formattingMetrics.HAnsiCharCount += fontTypeArray.Count(ft => ft == FormattingAssembler.FontType.HAnsi);
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
                        formattingMetrics.CsCharCount += runText.Length;
                        formattingMetrics.CsRunCount++;
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

        private static string GetFontFromFontType(FormattingAssembler.CharStyleAttributes csa, FormattingAssembler.FontType ft)
        {
            return ft switch
            {
                FormattingAssembler.FontType.Ascii => csa.AsciiFont,
                FormattingAssembler.FontType.CS => csa.CsFont,
                FormattingAssembler.FontType.EastAsia => csa.EastAsiaFont,
                FormattingAssembler.FontType.HAnsi => csa.HAnsiFont,
                _ => csa.AsciiFont
            };
        }

        public static XElement GetXlsxMetrics(SmlDocument smlDoc, MetricsGetterSettings settings)
        {
            using var streamDoc = new OpenXmlMemoryStreamDocument(smlDoc);
            using SpreadsheetDocument sDoc = streamDoc.GetSpreadsheetDocument();
            var metrics = new List<XElement>();

            ValidateAgainstSpecificVersion(sDoc, metrics,
                DocumentFormat.OpenXml.FileFormatVersions.Office2007, H.SdkValidationError2007);

            ValidateAgainstSpecificVersion(sDoc, metrics, DocumentFormat.OpenXml.FileFormatVersions.Office2010,
                H.SdkValidationError2010);
#if !NET35
            ValidateAgainstSpecificVersion(sDoc, metrics, DocumentFormat.OpenXml.FileFormatVersions.Office2013,
                H.SdkValidationError2013);
#endif
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
            WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
            XDocument xd = workbookPart.GetXDocument();

            var partInformation =
                new XElement(H.Sheets,
                    xd.Root!
                        .Elements(S.sheets)
                        .Elements(S.sheet)
                        .Select(sh =>
                        {
                            var rid = (string) sh.Attribute(R.id);
                            var worksheetPart = (WorksheetPart) workbookPart.GetPartById(rid);
                            var sheetName = (string)sh.Attribute("name");
                            return GetTableInfoForSheet(spreadsheet, worksheetPart, sheetName, settings);
                        }));

            return partInformation;
        }

        private static XElement GetTableInfoForSheet(
            SpreadsheetDocument spreadsheetDocument,
            WorksheetPart sheetPart,
            string sheetName,
            MetricsGetterSettings settings)
        {
            XDocument xd = sheetPart.GetXDocument();

            var sheetInformation = new XElement(H.Sheet,
                new XAttribute(H.Name, sheetName),
                xd.Root!.Elements(S.tableParts)
                    .Elements(S.tablePart)
                    .Select(tp =>
                    {
                        var rId = (string) tp.Attribute(R.id);
                        var tablePart = (TableDefinitionPart) sheetPart.GetPartById(rId);
                        XDocument txd = tablePart.GetXDocument();
                        var tableName = (string) txd.Root!.Attribute("displayName");
                        XElement tableCellData = null;

                        if (settings.IncludeXlsxTableCellData)
                        {
                            Table xlsxTable = spreadsheetDocument.Table(tableName);

                            tableCellData = new XElement(H.TableData,
                                xlsxTable.TableRows()
                                    .Select(row => new XElement(H.Row,
                                        xlsxTable.TableColumns()
                                            .Select(col => new XElement(H.Cell,
                                                new XAttribute(H.Name, col.Name),
                                                new XAttribute(H.Val, (string) row[col.Name]))))));
                        }

                        var table = new XElement(H.Table,
                            new XAttribute(H.Name, (string) txd.Root.Attribute("name")),
                            new XAttribute(H.DisplayName, tableName),
                            new XElement(H.Columns,
                                txd.Root
                                    .Elements(S.tableColumns)
                                    .Elements(S.tableColumn)
                                    .Select(tc => new XElement(H.Column, new XAttribute(H.Name, (string) tc.Attribute("name"))))),
                            tableCellData);

                        return table;
                    }));

            return !sheetInformation.HasElements ? null : sheetInformation;
        }

        public static XElement GetPptxMetrics(PmlDocument pmlDoc, MetricsGetterSettings settings)
        {
            using var streamDoc = new OpenXmlMemoryStreamDocument(pmlDoc);
            using PresentationDocument pDoc = streamDoc.GetPresentationDocument();

            var metrics = new List<XElement>();

            ValidateAgainstSpecificVersion(pDoc, metrics, DocumentFormat.OpenXml.FileFormatVersions.Office2007,
                H.SdkValidationError2007);

            ValidateAgainstSpecificVersion(pDoc, metrics, DocumentFormat.OpenXml.FileFormatVersions.Office2010,
                H.SdkValidationError2010);
#if !NET35
            ValidateAgainstSpecificVersion(pDoc, metrics, DocumentFormat.OpenXml.FileFormatVersions.Office2013,
                H.SdkValidationError2013);
#endif
            return new XElement(H.Metrics,
                new XAttribute(H.FileName, pmlDoc.FileName),
                new XAttribute(H.FileType, "PresentationML"),
                metrics,
                settings.RetrieveNamespaceList ? RetrieveNamespaceList(pDoc) : null,
                settings.RetrieveContentTypeList ? RetrieveContentTypeList(pDoc) : null);
        }

        private static object GetStyleHierarchy(WordprocessingDocument document)
        {
            StyleDefinitionsPart stylePart = document.MainDocumentPart?.StyleDefinitionsPart;

            if (stylePart == null)
            {
                return null;
            }

            List<XElement> styleElements = stylePart.GetXDocument().Elements(W.styles).Elements(W.style).ToList();

            List<string> stylesWithPath = styleElements
                .Select(s =>
                {
                    var styleString = (string) s.Attribute(W.styleId);
                    XElement thisStyle = s;

                    while (true)
                    {
                        var baseStyle = (string) thisStyle.Elements(W.basedOn).Attributes(W.val).FirstOrDefault();

                        if (baseStyle == null)
                        {
                            break;
                        }

                        styleString = baseStyle + "/" + styleString;
                        thisStyle = styleElements.FirstOrDefault(ts => ts.Attribute(W.styleId)?.Value == baseStyle);

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

            foreach (string item in stylesWithPath)
            {
                string[] styleChain = item.Split('/');
                XElement elementToAddTo = styleHierarchy;

                foreach (string inChain in styleChain.PtSkipLast(1))
                {
                    elementToAddTo = elementToAddTo.Elements(H.Style).First(z => z.Attribute(H.Id)?.Value == inChain);
                }

                string styleToAdd = styleChain.Last();

                elementToAddTo.Add(new XElement(H.Style,
                    new XAttribute(H.Id, styleChain.Last()),
                    new XAttribute(H.Type,
                        (string) styleElements
                            .First(z => z.Attribute(W.styleId)?.Value == styleToAdd)
                            .Attribute(W.type))));
            }

            return styleHierarchy;
        }

        private static XElement GetMetricsForWmlPart(OpenXmlPart part, MetricsGetterSettings settings)
        {
            XElement contentControls = null;

            if (part is MainDocumentPart or HeaderPart or FooterPart or FootnotesPart or EndnotesPart)
            {
                contentControls = (XElement) GetContentControlsTransform(part.GetXElement(), settings);

                if (!contentControls.HasElements)
                {
                    contentControls = null;
                }
            }

            var partMetrics = new XElement(H.Part,
                new XAttribute(H.ContentType, part.ContentType),
                new XAttribute(H.Uri, part.Uri.ToString()),
                contentControls);

            return partMetrics.HasElements ? partMetrics : null;
        }

        private static object GetContentControlsTransform(XNode node, MetricsGetterSettings settings)
        {
            if (node is XElement element)
            {
                if (element == element.Document?.Root)
                {
                    return new XElement(H.ContentControls,
                        element.Nodes().Select(n => GetContentControlsTransform(n, settings)));
                }

                if (element.Name == W.sdt)
                {
                    var tag = (string) element.Elements(W.sdtPr).Elements(W.tag).Attributes(W.val).FirstOrDefault();
                    XAttribute tagAttr = tag != null ? new XAttribute(H.Tag, tag) : null;

                    var alias = (string) element.Elements(W.sdtPr).Elements(W.alias).Attributes(W.val).FirstOrDefault();
                    XAttribute aliasAttr = alias != null ? new XAttribute(H.Alias, alias) : null;

                    var xPathAttr = new XAttribute(H.XPath, element.GetXPath());

                    bool isText = element.Elements(W.sdtPr).Elements(W.text).Any();
                    bool isBibliography = element.Elements(W.sdtPr).Elements(W.bibliography).Any();
                    bool isCitation = element.Elements(W.sdtPr).Elements(W.citation).Any();
                    bool isComboBox = element.Elements(W.sdtPr).Elements(W.comboBox).Any();
                    bool isDate = element.Elements(W.sdtPr).Elements(W.date).Any();
                    bool isDocPartList = element.Elements(W.sdtPr).Elements(W.docPartList).Any();
                    bool isDocPartObj = element.Elements(W.sdtPr).Elements(W.docPartObj).Any();
                    bool isDropDownList = element.Elements(W.sdtPr).Elements(W.dropDownList).Any();
                    bool isEquation = element.Elements(W.sdtPr).Elements(W.equation).Any();
                    bool isGroup = element.Elements(W.sdtPr).Elements(W.group).Any();
                    bool isPicture = element.Elements(W.sdtPr).Elements(W.picture).Any();

                    bool isRichText = element.Elements(W.sdtPr).Elements(W.richText).Any() ||
                        (!isText &&
                            !isBibliography &&
                            !isCitation &&
                            !isComboBox &&
                            !isDate &&
                            !isDocPartList &&
                            !isDocPartObj &&
                            !isDropDownList &&
                            !isEquation &&
                            !isGroup &&
                            !isPicture);

                    string type = null;

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

            return settings.IncludeTextInContentControls ? node : null;
        }

        public static int GetTextWidth(FontFamily ff, FontStyle fs, decimal sz, string text)
        {
            try
            {
                return GetTextWidthCore(ff, fs, sz, text);
            }
            catch (ArgumentException)
            {
                try
                {
                    const FontStyle fs2 = FontStyle.Regular;
                    return GetTextWidthCore(ff, fs2, sz, text);
                }
                catch (ArgumentException)
                {
                    const FontStyle fs2 = FontStyle.Bold;

                    try
                    {
                        return GetTextWidthCore(ff, fs2, sz, text);
                    }
                    catch (ArgumentException)
                    {
                        // if both regular and bold fail, then get metrics for Times New Roman
                        // use the original FontStyle (in fs)
                        var ff2 = new FontFamily("Times New Roman");
                        return GetTextWidthCore(ff2, fs, sz, text);
                    }
                }
            }
            catch (OverflowException)
            {
                // This happened on Azure but interestingly enough not while testing locally.
                return 0;
            }
        }

        private static int GetTextWidthCore(FontFamily ff, FontStyle fs, decimal sz, string text)
        {
            try
            {
                using var f = new Font(ff, (float)sz / 2f, fs);
                var proposedSize = new Size(int.MaxValue, int.MaxValue);
                SizeF sf = Graphics.Value.MeasureString(text, f, proposedSize);
                return (int)sf.Width;
            }
            catch
            {
                return 0;
            }
        }
    }

    // ReSharper disable InconsistentNaming

#pragma warning disable IDE1006 // Naming Styles

    internal static class H
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

#pragma warning restore IDE1006 // Naming Styles
}
