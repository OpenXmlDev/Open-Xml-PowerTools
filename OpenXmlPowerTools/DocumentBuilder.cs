// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#define TestForUnsupportedDocuments
#define MergeStylesWithSameNames

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    public partial class WmlDocument : OpenXmlPowerToolsDocument
    {
        public IEnumerable<WmlDocument> SplitOnSections()
        {
            return DocumentBuilder.SplitOnSections(this);
        }
    }

    public class Source
    {
        public WmlDocument WmlDocument { get; set; }
        public int Start { get; set; }
        public int Count { get; set; }
        public bool KeepSections { get; set; }
        public bool DiscardHeadersAndFootersInKeptSections { get; set; }
        public string InsertId { get; set; }

        public Source(string fileName)
        {
            WmlDocument = new WmlDocument(fileName);
            Start = 0;
            Count = Int32.MaxValue;
            KeepSections = false;
            InsertId = null;
        }

        public Source(WmlDocument source)
        {
            WmlDocument = source;
            Start = 0;
            Count = Int32.MaxValue;
            KeepSections = false;
            InsertId = null;
        }

        public Source(string fileName, bool keepSections)
        {
            WmlDocument = new WmlDocument(fileName);
            Start = 0;
            Count = Int32.MaxValue;
            KeepSections = keepSections;
            InsertId = null;
        }

        public Source(WmlDocument source, bool keepSections)
        {
            WmlDocument = source;
            Start = 0;
            Count = Int32.MaxValue;
            KeepSections = keepSections;
            InsertId = null;
        }

        public Source(string fileName, string insertId)
        {
            WmlDocument = new WmlDocument(fileName);
            Start = 0;
            Count = Int32.MaxValue;
            KeepSections = false;
            InsertId = insertId;
        }

        public Source(WmlDocument source, string insertId)
        {
            WmlDocument = source;
            Start = 0;
            Count = Int32.MaxValue;
            KeepSections = false;
            InsertId = insertId;
        }

        public Source(string fileName, int start, bool keepSections)
        {
            WmlDocument = new WmlDocument(fileName);
            Start = start;
            Count = Int32.MaxValue;
            KeepSections = keepSections;
            InsertId = null;
        }

        public Source(WmlDocument source, int start, bool keepSections)
        {
            WmlDocument = source;
            Start = start;
            Count = Int32.MaxValue;
            KeepSections = keepSections;
            InsertId = null;
        }

        public Source(string fileName, int start, string insertId)
        {
            WmlDocument = new WmlDocument(fileName);
            Start = start;
            Count = Int32.MaxValue;
            KeepSections = false;
            InsertId = insertId;
        }

        public Source(WmlDocument source, int start, string insertId)
        {
            WmlDocument = source;
            Start = start;
            Count = Int32.MaxValue;
            KeepSections = false;
            InsertId = insertId;
        }

        public Source(string fileName, int start, int count, bool keepSections)
        {
            WmlDocument = new WmlDocument(fileName);
            Start = start;
            Count = count;
            KeepSections = keepSections;
            InsertId = null;
        }

        public Source(WmlDocument source, int start, int count, bool keepSections)
        {
            WmlDocument = source;
            Start = start;
            Count = count;
            KeepSections = keepSections;
            InsertId = null;
        }

        public Source(string fileName, int start, int count, string insertId)
        {
            WmlDocument = new WmlDocument(fileName);
            Start = start;
            Count = count;
            KeepSections = false;
            InsertId = insertId;
        }

        public Source(WmlDocument source, int start, int count, string insertId)
        {
            WmlDocument = source;
            Start = start;
            Count = count;
            KeepSections = false;
            InsertId = insertId;
        }
    }

    public class DocumentBuilderSettings
    {
        public HashSet<string> CustomXmlGuidList = null;
        public bool NormalizeStyleIds = false;
    }

    public static class DocumentBuilder
    {
        public static void BuildDocument(List<Source> sources, string fileName)
        {
            using (OpenXmlMemoryStreamDocument streamDoc = OpenXmlMemoryStreamDocument.CreateWordprocessingDocument())
            {
                using (WordprocessingDocument output = streamDoc.GetWordprocessingDocument())
                {
                    BuildDocument(sources, output, new DocumentBuilderSettings());
                    output.Close();
                }
                streamDoc.GetModifiedDocument().SaveAs(fileName);
            }
        }

        public static void BuildDocument(List<Source> sources, string fileName, DocumentBuilderSettings settings)
        {
            using (OpenXmlMemoryStreamDocument streamDoc = OpenXmlMemoryStreamDocument.CreateWordprocessingDocument())
            {
                using (WordprocessingDocument output = streamDoc.GetWordprocessingDocument())
                {
                    BuildDocument(sources, output, settings);
                    output.Close();
                }
                streamDoc.GetModifiedDocument().SaveAs(fileName);
            }
        }

        public static WmlDocument BuildDocument(List<Source> sources)
        {
            using (OpenXmlMemoryStreamDocument streamDoc = OpenXmlMemoryStreamDocument.CreateWordprocessingDocument())
            {
                using (WordprocessingDocument output = streamDoc.GetWordprocessingDocument())
                {
                    BuildDocument(sources, output, new DocumentBuilderSettings());
                    output.Close();
                }
                return streamDoc.GetModifiedWmlDocument();
            }
        }

        public static WmlDocument BuildDocument(List<Source> sources, DocumentBuilderSettings settings)
        {
            using (OpenXmlMemoryStreamDocument streamDoc = OpenXmlMemoryStreamDocument.CreateWordprocessingDocument())
            {
                using (WordprocessingDocument output = streamDoc.GetWordprocessingDocument())
                {
                    BuildDocument(sources, output, settings);
                    output.Close();
                }
                return streamDoc.GetModifiedWmlDocument();
            }
        }

        private struct TempSource
        {
            public int Start;
            public int Count;
        };

        private class Atbi
        {
            public XElement BlockLevelContent;
            public int Index;
        }

        private class Atbid
        {
            public XElement BlockLevelContent;
            public int Index;
            public int Div;
        }

        private const string Yes = "yes";
        private const string Utf8 = "UTF-8";
        private const string OnePointZero = "1.0";

        public static IEnumerable<WmlDocument> SplitOnSections(WmlDocument doc)
        {
            List<TempSource> tempSourceList;
            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(doc))
            using (WordprocessingDocument document = streamDoc.GetWordprocessingDocument())
            {
                XDocument mainDocument = document.MainDocumentPart.GetXDocument();
                var divs = mainDocument
                    .Root
                    .Element(W.body)
                    .Elements()
                    .Select((p, i) => new Atbi
                    {
                        BlockLevelContent = p,
                        Index = i,
                    })
                    .Rollup(new Atbid
                        {
                            BlockLevelContent = (XElement)null,
                            Index = -1,
                            Div = 0,
                        },
                        (b, p) =>
                        {
                            XElement elementBefore = b.BlockLevelContent
                                .SiblingsBeforeSelfReverseDocumentOrder()
                                .FirstOrDefault();
                            if (elementBefore != null && elementBefore.Descendants(W.sectPr).Any())
                                return new Atbid
                                {
                                    BlockLevelContent = b.BlockLevelContent,
                                    Index = b.Index,
                                    Div = p.Div + 1,
                                };
                            return new Atbid
                            {
                                BlockLevelContent = b.BlockLevelContent,
                                Index = b.Index,
                                Div = p.Div,
                            };
                        });
                var groups = divs
                    .GroupAdjacent(b => b.Div);
                tempSourceList = groups
                    .Select(g => new TempSource
                    {
                        Start = g.First().Index,
                        Count = g.Count(),
                    })
                    .ToList();
                foreach (var ts in tempSourceList)
                {
                    List<Source> sources = new List<Source>()
                    {
                        new Source(doc, ts.Start, ts.Count, true)
                    };
                    WmlDocument newDoc = DocumentBuilder.BuildDocument(sources);
                    newDoc = AdjustSectionBreak(newDoc);
                    yield return newDoc;
                }
            }
        }

        private static WmlDocument AdjustSectionBreak(WmlDocument doc)
        {
            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(doc))
            {
                using (WordprocessingDocument document = streamDoc.GetWordprocessingDocument())
                {
                    XDocument mainXDoc = document.MainDocumentPart.GetXDocument();
                    XElement lastElement = mainXDoc.Root
                        .Element(W.body)
                        .Elements()
                        .LastOrDefault();
                    if (lastElement != null)
                    {
                        if (lastElement.Name != W.sectPr &&
                            lastElement.Descendants(W.sectPr).Any())
                        {
                            mainXDoc.Root.Element(W.body).Add(lastElement.Descendants(W.sectPr).First());
                            lastElement.Descendants(W.sectPr).Remove();
                            if (!lastElement.Elements()
                                .Where(e => e.Name != W.pPr)
                                .Any())
                                lastElement.Remove();
                            document.MainDocumentPart.PutXDocument();
                        }
                    }
                }
                return streamDoc.GetModifiedWmlDocument();
            }
        }

        private static void BuildDocument(List<Source> sources, WordprocessingDocument output, DocumentBuilderSettings settings)
        {
            WmlDocument wmlGlossaryDocument = CoalesceGlossaryDocumentParts(sources, settings);

            if (RelationshipMarkup == null)
                InitRelationshipMarkup();

            // This list is used to eliminate duplicate images
            List<ImageData> images = new List<ImageData>();
            XDocument mainPart = output.MainDocumentPart.GetXDocument();
            mainPart.Declaration.Standalone = Yes;
            mainPart.Declaration.Encoding = Utf8;
            mainPart.Root.ReplaceWith(
                new XElement(W.document, NamespaceAttributes,
                    new XElement(W.body)));
            if (sources.Count > 0)
            {
                // the following function makes sure that for a given style name, the same style ID is used for all documents.
                if (settings != null && settings.NormalizeStyleIds)
                    sources = NormalizeStyleNamesAndIds(sources);

                using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(sources[0].WmlDocument))
                using (WordprocessingDocument doc = streamDoc.GetWordprocessingDocument())
                {
                    CopyStartingParts(doc, output, images);
                    CopySpecifiedCustomXmlParts(doc, output, settings);
                }

                int sourceNum2 = 0;
                foreach (Source source in sources)
                {
                    if (source.InsertId != null)
                    {
                        while (true)
                        {
#if false
                            modify AppendDocument so that it can take a part.
                            for each in main document part, header parts, footer parts
                                are there any PtOpenXml.Insert elements in any of them?
                            if so, then open and process all.
#endif
                            bool foundInMainDocPart = false;
                            XDocument mainXDoc = output.MainDocumentPart.GetXDocument();
                            if (mainXDoc.Descendants(PtOpenXml.Insert).Any(d => (string)d.Attribute(PtOpenXml.Id) == source.InsertId))
                                foundInMainDocPart = true;
                            if (foundInMainDocPart)
                            {
                                using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(source.WmlDocument))
                                using (WordprocessingDocument doc = streamDoc.GetWordprocessingDocument())
                                {
#if TestForUnsupportedDocuments
                                    // throws exceptions if a document contains unsupported content
                                    TestForUnsupportedDocument(doc, sources.IndexOf(source));
#endif
                                    if (foundInMainDocPart)
                                    {
                                        if (source.KeepSections && source.DiscardHeadersAndFootersInKeptSections)
                                            RemoveHeadersAndFootersFromSections(doc);
                                        else if (source.KeepSections)
                                            ProcessSectionsForLinkToPreviousHeadersAndFooters(doc);

                                        List<XElement> contents = doc.MainDocumentPart.GetXDocument()
                                            .Root
                                            .Element(W.body)
                                            .Elements()
                                            .Skip(source.Start)
                                            .Take(source.Count)
                                            .ToList();

                                        try
                                        {
                                            AppendDocument(doc, output, contents, source.KeepSections, source.InsertId, images);
                                        }
                                        catch (DocumentBuilderInternalException dbie)
                                        {
                                            if (dbie.Message.Contains("{0}"))
                                                throw new DocumentBuilderException(string.Format(dbie.Message, sourceNum2));
                                            else
                                                throw dbie;
                                        }
                                    }
                                }
                            }
                            else
                                break;
                        }
                    }
                    else
                    {
                        using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(source.WmlDocument))
                        using (WordprocessingDocument doc = streamDoc.GetWordprocessingDocument())
                        {
#if TestForUnsupportedDocuments
                            // throws exceptions if a document contains unsupported content
                            TestForUnsupportedDocument(doc, sources.IndexOf(source));
#endif
                            if (source.KeepSections && source.DiscardHeadersAndFootersInKeptSections)
                                RemoveHeadersAndFootersFromSections(doc);
                            else if (source.KeepSections)
                                ProcessSectionsForLinkToPreviousHeadersAndFooters(doc);

                            var body = doc.MainDocumentPart.GetXDocument()
                                .Root
                                .Element(W.body);

                            if (body == null)
                                throw new DocumentBuilderException(
                                    String.Format("Source {0} is unsupported document - contains no body element in the correct namespace", sourceNum2));

                            List<XElement> contents = body
                                .Elements()
                                .Skip(source.Start)
                                .Take(source.Count)
                                .ToList();
                            try
                            {
                                AppendDocument(doc, output, contents, source.KeepSections, null, images);
                            }
                            catch (DocumentBuilderInternalException dbie)
                            {
                                if (dbie.Message.Contains("{0}"))
                                    throw new DocumentBuilderException(string.Format(dbie.Message, sourceNum2));
                                else
                                    throw dbie;
                            }
                        }
                    }
                    ++sourceNum2;
                }
                if (!sources.Any(s => s.KeepSections))
                {
                    using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(sources[0].WmlDocument))
                    using (WordprocessingDocument doc = streamDoc.GetWordprocessingDocument())
                    {
                        var body = doc.MainDocumentPart.GetXDocument().Root.Element(W.body);

                        if (body != null && body.Elements().Any())
						{
							var sectPr = doc.MainDocumentPart.GetXDocument().Root.Elements(W.body)
								.Elements().LastOrDefault();
							if (sectPr != null && sectPr.Name == W.sectPr)
							{
								AddSectionAndDependencies(doc, output, sectPr, images);
								output.MainDocumentPart.GetXDocument().Root.Element(W.body).Add(sectPr);
							}
						}
                    }
                }
                else
                {
                    FixUpSectionProperties(output);

                    // Any sectPr elements that do not have headers and footers should take their headers and footers from the *next* section,
                    // i.e. from the running section.
                    var mxd = output.MainDocumentPart.GetXDocument();
                    var sections = mxd.Descendants(W.sectPr).Reverse().ToList();

                    CachedHeaderFooter[] cachedHeaderFooter = new[]
                    {
                        new CachedHeaderFooter() { Ref = W.headerReference, Type = "first" },
                        new CachedHeaderFooter() { Ref = W.headerReference, Type = "even" },
                        new CachedHeaderFooter() { Ref = W.headerReference, Type = "default" },
                        new CachedHeaderFooter() { Ref = W.footerReference, Type = "first" },
                        new CachedHeaderFooter() { Ref = W.footerReference, Type = "even" },
                        new CachedHeaderFooter() { Ref = W.footerReference, Type = "default" },
                    };

                    bool firstSection = true;
                    foreach (var sect in sections)
                    {
                        if (firstSection)
                        {
                            foreach (var hf in cachedHeaderFooter)
                            {
                                var referenceElement = sect.Elements(hf.Ref).FirstOrDefault(z => (string)z.Attribute(W.type) == hf.Type);
                                if (referenceElement != null)
                                    hf.CachedPartRid = (string)referenceElement.Attribute(R.id);
                            }
                            firstSection = false;
                            continue;
                        }
                        else
                        {
                            CopyOrCacheHeaderOrFooter(output, cachedHeaderFooter, sect, W.headerReference, "first");
                            CopyOrCacheHeaderOrFooter(output, cachedHeaderFooter, sect, W.headerReference, "even");
                            CopyOrCacheHeaderOrFooter(output, cachedHeaderFooter, sect, W.headerReference, "default");
                            CopyOrCacheHeaderOrFooter(output, cachedHeaderFooter, sect, W.footerReference, "first");
                            CopyOrCacheHeaderOrFooter(output, cachedHeaderFooter, sect, W.footerReference, "even");
                            CopyOrCacheHeaderOrFooter(output, cachedHeaderFooter, sect, W.footerReference, "default");
                        }

                    }
                }

                // Now can process PtOpenXml:Insert elements in headers / footers
                int sourceNum = 0;
                foreach (Source source in sources)
                {
                    if (source.InsertId != null)
                    {
                        while (true)
                        {
#if false
                            this uses an overload of AppendDocument that takes a part.
                            for each in main document part, header parts, footer parts
                                are there any PtOpenXml.Insert elements in any of them?
                            if so, then open and process all.
#endif
                            bool foundInHeadersFooters = false;
                            if (output.MainDocumentPart.HeaderParts.Any(hp =>
                            {
                                var hpXDoc = hp.GetXDocument();
                                return hpXDoc.Descendants(PtOpenXml.Insert).Any(d => (string)d.Attribute(PtOpenXml.Id) == source.InsertId);
                            }))
                                foundInHeadersFooters = true;
                            if (output.MainDocumentPart.FooterParts.Any(fp =>
                            {
                                var hpXDoc = fp.GetXDocument();
                                return hpXDoc.Descendants(PtOpenXml.Insert).Any(d => (string)d.Attribute(PtOpenXml.Id) == source.InsertId);
                            }))
                                foundInHeadersFooters = true;

                            if (foundInHeadersFooters)
                            {
                                using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(source.WmlDocument))
                                using (WordprocessingDocument doc = streamDoc.GetWordprocessingDocument())
                                {
#if TestForUnsupportedDocuments
                                    // throws exceptions if a document contains unsupported content
                                    TestForUnsupportedDocument(doc, sources.IndexOf(source));
#endif
                                    var partList = output.MainDocumentPart.HeaderParts.Cast<OpenXmlPart>().Concat(output.MainDocumentPart.FooterParts.Cast<OpenXmlPart>()).ToList();
                                    foreach (var part in partList)
                                    {
                                        var partXDoc = part.GetXDocument();
                                        if (!partXDoc.Descendants(PtOpenXml.Insert).Any(d => (string)d.Attribute(PtOpenXml.Id) == source.InsertId))
                                            continue;
                                        List<XElement> contents = doc.MainDocumentPart.GetXDocument()
                                            .Root
                                            .Element(W.body)
                                            .Elements()
                                            .Skip(source.Start)
                                            .Take(source.Count)
                                            .ToList();

                                        try
                                        {
                                            AppendDocument(doc, output, part, contents, source.KeepSections, source.InsertId, images);
                                        }
                                        catch (DocumentBuilderInternalException dbie)
                                        {
                                            if (dbie.Message.Contains("{0}"))
                                                throw new DocumentBuilderException(string.Format(dbie.Message, sourceNum));
                                            else
                                                throw dbie;
                                        }
                                    }
                                }
                            }
                            else
                                break;
                        }
                    }
                    ++sourceNum;
                }
                if (sources.Any(s => s.KeepSections) && !output.MainDocumentPart.GetXDocument().Root.Descendants(W.sectPr).Any())
                {
                    using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(sources[0].WmlDocument))
                    using (WordprocessingDocument doc = streamDoc.GetWordprocessingDocument())
                    {
                        var sectPr = doc.MainDocumentPart.GetXDocument().Root.Element(W.body)
                            .Elements().LastOrDefault();
                        if (sectPr != null && sectPr.Name == W.sectPr)
                        {
                            AddSectionAndDependencies(doc, output, sectPr, images);
                            output.MainDocumentPart.GetXDocument().Root.Element(W.body).Add(sectPr);
                        }
                    }
                }
                AdjustDocPrIds(output);
            }

            if (wmlGlossaryDocument != null)
                WriteGlossaryDocumentPart(wmlGlossaryDocument, output, images);

            foreach (var part in output.GetAllParts())
                if (part.Annotation<XDocument>() != null)
                    part.PutXDocument();
        }

        // there are two scenarios that need to be handled
        // - if I find a style name that maps to a style ID different from one already mapped
        // - if a style name maps to a style ID that is already used for a different style
        // - then need to correct things
        //   - make a complete list of all things that need to be changed, for every correction
        //   - do the corrections all at one
        //   - mark the document as changed, and change it in the sources.
        private static List<Source> NormalizeStyleNamesAndIds(List<Source> sources)
        {
            Dictionary<string, string> styleNameMap = new Dictionary<string, string>();
            HashSet<string> styleIds = new HashSet<string>();
            List<Source> newSources = new List<Source>();

            foreach (var src in sources)
            {
                var newSrc = AddAndRectify(src, styleNameMap, styleIds);
                newSources.Add(newSrc);
            }
            return newSources;
        }

        private static Source AddAndRectify(Source src, Dictionary<string, string> styleNameMap, HashSet<string> styleIds)
        {
            bool modified = false;
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(src.WmlDocument.DocumentByteArray, 0, src.WmlDocument.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    Dictionary<string, string> correctionList = new Dictionary<string, string>();
                    var thisStyleNameMap = GetStyleNameMap(wDoc);
                    foreach (var pair in thisStyleNameMap)
                    {
                        var styleName = pair.Key;
                        var styleId = pair.Value;
                        // if the styleNameMap does not contain an entry for this name
                        if (!styleNameMap.ContainsKey(styleName))
                        {
                            // if the id is already used
                            if (styleIds.Contains(styleId))
                            {
                                // this style uses a styleId that is used for another style.
                                // randomly generate new styleId
                                while (true)
                                {
                                    var newStyleId = GenStyleIdFromStyleName(styleName);
                                    if (! styleIds.Contains(newStyleId))
                                    {
                                        correctionList.Add(styleId, newStyleId);
                                        styleNameMap.Add(styleName, newStyleId);
                                        styleIds.Add(newStyleId);
                                        break;
                                    }
                                }
                            }
                            // otherwise we just add to the styleNameMap
                            else
                            {
                                styleNameMap.Add(styleName, styleId);
                                styleIds.Add(styleId);
                            }
                        }
                        // but if the styleNameMap does contain an entry for this name
                        else
                        {
                            // if the id is the same as the existing ID, then nothing to do
                            if (styleNameMap[styleName] == styleId)
                                continue;
                            correctionList.Add(styleId, styleNameMap[styleName]);
                        }
                    }
                    if (correctionList.Any())
                    {
                        modified = true;
                        AdjustStyleIdsForDocument(wDoc, correctionList);
                    }
                }
                if (modified)
                {
                    var newWmlDocument = new WmlDocument(src.WmlDocument.FileName, ms.ToArray());
                    var newSrc = new Source(newWmlDocument, src.Start, src.Count, src.KeepSections);
                    newSrc.DiscardHeadersAndFootersInKeptSections = src.DiscardHeadersAndFootersInKeptSections;
                    newSrc.InsertId = src.InsertId;
                    return newSrc;
                }
            }
            return src;
        }

#if false
application/vnd.ms-word.styles.textEffects+xml                                                      styles/style/@styleId
application/vnd.ms-word.styles.textEffects+xml                                                      styles/style/basedOn/@val
application/vnd.ms-word.styles.textEffects+xml                                                      styles/style/link/@val
application/vnd.ms-word.styles.textEffects+xml                                                      styles/style/next/@val

application/vnd.ms-word.stylesWithEffects+xml                                                       styles/style/@styleId
application/vnd.ms-word.stylesWithEffects+xml                                                       styles/style/basedOn/@val
application/vnd.ms-word.stylesWithEffects+xml                                                       styles/style/link/@val
application/vnd.ms-word.stylesWithEffects+xml                                                       styles/style/next/@val

application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml                         pPr/pStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml                         rPr/rStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml                         tblPr/tblStyle/@val

application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml                    pPr/pStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml                    rPr/rStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml                    tblPr/tblStyle/@val

application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml                         pPr/pStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml                         rPr/rStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml                         tblPr/tblStyle/@val

application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml                           pPr/pStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml                           rPr/rStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml                           tblPr/tblStyle/@val

application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml                        pPr/pStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml                        rPr/rStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml                        tblPr/tblStyle/@val

application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml                           pPr/pStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml                           rPr/rStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml                           tblPr/tblStyle/@val

application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml                        abstractNum/lvl/pStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml                        abstractNum/numStyleLink/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml                        abstractNum/styleLink/@val

application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml                         settings/clickAndTypeStyle/@val

Name, not ID
===================================
application/vnd.ms-word.styles.textEffects+xml                                                      styles/style/name/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.stylesWithEffects+xml                styles/style/name/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml                           styles/style/name/@val
application/vnd.ms-word.stylesWithEffects+xml                                                       styles/style/name/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.stylesWithEffects+xml                latentStyles/lsdException/@name
application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml                           latentStyles/lsdException/@name
application/vnd.ms-word.stylesWithEffects+xml                                                       latentStyles/lsdException/@name
application/vnd.ms-word.styles.textEffects+xml                                                      latentStyles/lsdException/@name
#endif

        private static void AdjustStyleIdsForDocument(WordprocessingDocument wDoc, Dictionary<string, string> correctionList)
        {
            // update styles part
            UpdateStyleIdsForStylePart(wDoc.MainDocumentPart.StyleDefinitionsPart, correctionList);
            if (wDoc.MainDocumentPart.StylesWithEffectsPart != null)
                UpdateStyleIdsForStylePart(wDoc.MainDocumentPart.StylesWithEffectsPart, correctionList);

            // update content parts
            UpdateStyleIdsForContentPart(wDoc.MainDocumentPart, correctionList);
            foreach (var part in wDoc.MainDocumentPart.HeaderParts)
                UpdateStyleIdsForContentPart(part, correctionList);
            foreach (var part in wDoc.MainDocumentPart.FooterParts)
                UpdateStyleIdsForContentPart(part, correctionList);
            if (wDoc.MainDocumentPart.FootnotesPart != null)
                UpdateStyleIdsForContentPart(wDoc.MainDocumentPart.FootnotesPart, correctionList);
            if (wDoc.MainDocumentPart.EndnotesPart != null)
                UpdateStyleIdsForContentPart(wDoc.MainDocumentPart.EndnotesPart, correctionList);
            if (wDoc.MainDocumentPart.WordprocessingCommentsPart != null)
                UpdateStyleIdsForContentPart(wDoc.MainDocumentPart.WordprocessingCommentsPart, correctionList);
            if (wDoc.MainDocumentPart.WordprocessingCommentsExPart != null)
                UpdateStyleIdsForContentPart(wDoc.MainDocumentPart.WordprocessingCommentsExPart, correctionList);

            // update numbering part
            UpdateStyleIdsForNumberingPart(wDoc.MainDocumentPart.NumberingDefinitionsPart, correctionList);
        }

        private static void UpdateStyleIdsForNumberingPart(OpenXmlPart part, Dictionary<string, string> correctionList)
        {
#if false
application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml                        abstractNum/lvl/pStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml                        abstractNum/numStyleLink/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml                        abstractNum/styleLink/@val
#endif
            var numXDoc = part.GetXDocument();
            var numAttributeChangeList = correctionList
                .Select(cor =>
                    new
                    {
                        NewId = cor.Value,
                        PStyleAttributesToChange = numXDoc
                            .Descendants(W.pStyle)
                            .Attributes(W.val)
                            .Where(a => a.Value == cor.Key)
                            .ToList(),
                        NumStyleLinkAttributesToChange = numXDoc
                            .Descendants(W.numStyleLink)
                            .Attributes(W.val)
                            .Where(a => a.Value == cor.Key)
                            .ToList(),
                        StyleLinkAttributesToChange = numXDoc
                            .Descendants(W.styleLink)
                            .Attributes(W.val)
                            .Where(a => a.Value == cor.Key)
                            .ToList(),
                    }
                )
                .ToList();
            foreach (var item in numAttributeChangeList)
            {
                foreach (var att in item.PStyleAttributesToChange)
                    att.Value = item.NewId;
                foreach (var att in item.NumStyleLinkAttributesToChange)
                    att.Value = item.NewId;
                foreach (var att in item.StyleLinkAttributesToChange)
                    att.Value = item.NewId;
            }
            part.PutXDocument();
        }

        private static void UpdateStyleIdsForStylePart(OpenXmlPart part, Dictionary<string, string> correctionList)
        {
#if false
application/vnd.ms-word.styles.textEffects+xml                                                      styles/style/@styleId
application/vnd.ms-word.styles.textEffects+xml                                                      styles/style/basedOn/@val
application/vnd.ms-word.styles.textEffects+xml                                                      styles/style/link/@val
application/vnd.ms-word.styles.textEffects+xml                                                      styles/style/next/@val
#endif
            var styleXDoc = part.GetXDocument();
            var styleAttributeChangeList = correctionList
                .Select(cor =>
                    new
                    {
                        NewId = cor.Value,
                        StyleIdAttributesToChange = styleXDoc
                            .Root
                            .Elements(W.style)
                            .Attributes(W.styleId)
                            .Where(a => a.Value == cor.Key)
                            .ToList(),
                        BasedOnAttributesToChange = styleXDoc
                            .Root
                            .Elements(W.style)
                            .Elements(W.basedOn)
                            .Attributes(W.val)
                            .Where(a => a.Value == cor.Key)
                            .ToList(),
                        NextAttributesToChange = styleXDoc
                            .Root
                            .Elements(W.style)
                            .Elements(W.next)
                            .Attributes(W.val)
                            .Where(a => a.Value == cor.Key)
                            .ToList(),
                        LinkAttributesToChange = styleXDoc
                            .Root
                            .Elements(W.style)
                            .Elements(W.link)
                            .Attributes(W.val)
                            .Where(a => a.Value == cor.Key)
                            .ToList(),
                    }
                )
                .ToList();
            foreach (var item in styleAttributeChangeList)
            {
                foreach (var att in item.StyleIdAttributesToChange)
                    att.Value = item.NewId;
                foreach (var att in item.BasedOnAttributesToChange)
                    att.Value = item.NewId;
                foreach (var att in item.NextAttributesToChange)
                    att.Value = item.NewId;
                foreach (var att in item.LinkAttributesToChange)
                    att.Value = item.NewId;
            }
            part.PutXDocument();
        }

        private static void UpdateStyleIdsForContentPart(OpenXmlPart part, Dictionary<string, string> correctionList)
        {
#if false
application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml                    pPr/pStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml                    rPr/rStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml                    tblPr/tblStyle/@val
#endif
            var xDoc = part.GetXDocument();
            var mainAttributeChangeList = correctionList
                .Select(cor =>
                    new
                    {
                        NewId = cor.Value,
                        PStyleAttributesToChange = xDoc
                            .Descendants(W.pStyle)
                            .Attributes(W.val)
                            .Where(a => a.Value == cor.Key)
                            .ToList(),
                        RStyleAttributesToChange = xDoc
                            .Descendants(W.rStyle)
                            .Attributes(W.val)
                            .Where(a => a.Value == cor.Key)
                            .ToList(),
                        TblStyleAttributesToChange = xDoc
                            .Descendants(W.tblStyle)
                            .Attributes(W.val)
                            .Where(a => a.Value == cor.Key)
                            .ToList(),
                    }
                )
                .ToList();
            foreach (var item in mainAttributeChangeList)
            {
                foreach (var att in item.PStyleAttributesToChange)
                    att.Value = item.NewId;
                foreach (var att in item.RStyleAttributesToChange)
                    att.Value = item.NewId;
                foreach (var att in item.TblStyleAttributesToChange)
                    att.Value = item.NewId;
            }
            part.PutXDocument();
        }

        private static string GenStyleIdFromStyleName(string styleName)
        {
            var newStyleId = styleName
                .Replace("_", "")
                .Replace("#", "")
                .Replace(".", "") + ((new Random()).Next(990) + 9).ToString();
            return newStyleId;
        }

        private static Dictionary<string, string> GetStyleNameMap(WordprocessingDocument wDoc)
        {
            var sxDoc = wDoc.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
            var thisDocumentDictionary = sxDoc
                .Root
                .Elements(W.style)
                .ToDictionary(
                    z => (string)z.Elements(W.name).Attributes(W.val).FirstOrDefault(),
                    z => (string)z.Attribute(W.styleId));
            return thisDocumentDictionary;
        }

#if false
        At various locations in Open-Xml-PowerTools, you will find examples of Open XML markup that is associated with code associated with
        querying or generating that markup.  This is an example of the GlossaryDocumentPart.

<w:glossaryDocument xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid wp14">
  <w:docParts>
    <w:docPart>
      <w:docPartPr>
        <w:name w:val="CDE7B64C7BB446AE905B622B0A882EB6" />
        <w:category>
          <w:name w:val="General" />
          <w:gallery w:val="placeholder" />
        </w:category>
        <w:types>
          <w:type w:val="bbPlcHdr" />
        </w:types>
        <w:behaviors>
          <w:behavior w:val="content" />
        </w:behaviors>
        <w:guid w:val="{13882A71-B5B7-4421-ACBB-9B61C61B3034}" />
      </w:docPartPr>
      <w:docPartBody>
        <w:p w:rsidR="00004EEA" w:rsidRDefault="00AD57F5" w:rsidP="00AD57F5">
#endif

        private static void WriteGlossaryDocumentPart(WmlDocument wmlGlossaryDocument, WordprocessingDocument output, List<ImageData> images)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(wmlGlossaryDocument.DocumentByteArray, 0, wmlGlossaryDocument.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    var fromXDoc = wDoc.MainDocumentPart.GetXDocument();

                    var outputGlossaryDocumentPart = output.MainDocumentPart.AddNewPart<GlossaryDocumentPart>();
                    var newXDoc = new XDocument(
                        new XDeclaration(OnePointZero, Utf8, Yes),
                        new XElement(W.glossaryDocument,
                            NamespaceAttributes,
                            new XElement(W.docParts,
                                fromXDoc.Descendants(W.docPart))));
                    outputGlossaryDocumentPart.PutXDocument(newXDoc);

                    CopyGlossaryDocumentPartsToGD(wDoc, output, fromXDoc.Root.Descendants(W.docPart), images);
                    CopyRelatedPartsForContentParts(wDoc.MainDocumentPart, outputGlossaryDocumentPart, new[] { fromXDoc.Root }, images);
                }
            }
        }

        private static WmlDocument CoalesceGlossaryDocumentParts(IEnumerable<Source> sources, DocumentBuilderSettings settings)
        {
            List<Source> allGlossaryDocuments = sources
                .Select(s => DocumentBuilder.ExtractGlossaryDocument(s.WmlDocument))
                .Where(s => s != null)
                .Select(s => new Source(s))
                .ToList();

            if (!allGlossaryDocuments.Any())
                return null;

            WmlDocument coalescedRaw = DocumentBuilder.BuildDocument(allGlossaryDocuments);

            // now need to do some fix up
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(coalescedRaw.DocumentByteArray, 0, coalescedRaw.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    var mainXDoc = wDoc.MainDocumentPart.GetXDocument();

                    var newBody = new XElement(W.body,
                        new XElement(W.docParts,
                            mainXDoc.Root.Element(W.body).Elements(W.docParts).Elements(W.docPart)));

                    mainXDoc.Root.Element(W.body).ReplaceWith(newBody);

                    wDoc.MainDocumentPart.PutXDocument();
                }

                WmlDocument coalescedGlossaryDocument = new WmlDocument("Coalesced.docx", ms.ToArray());

                return coalescedGlossaryDocument;
            }
        }

        private static void InitRelationshipMarkup()
        {
            RelationshipMarkup = new Dictionary<XName, XName[]>()
                {
                    //{ button,           new [] { image }},
                    { A.blip,             new [] { R.embed, R.link }},
                    { A.hlinkClick,       new [] { R.id }},
                    { A.relIds,           new [] { R.cs, R.dm, R.lo, R.qs }},
                    //{ a14:imgLayer,     new [] { R.embed }},
                    //{ ax:ocx,           new [] { R.id }},
                    { C.chart,            new [] { R.id }},
                    { C.externalData,     new [] { R.id }},
                    { C.userShapes,       new [] { R.id }},
                    { DGM.relIds,         new [] { R.cs, R.dm, R.lo, R.qs }},
                    { O.OLEObject,        new [] { R.id }},
                    { VML.fill,           new [] { R.id }},
                    { VML.imagedata,      new [] { R.href, R.id, R.pict }},
                    { VML.stroke,         new [] { R.id }},
                    { W.altChunk,         new [] { R.id }},
                    { W.attachedTemplate, new [] { R.id }},
                    { W.control,          new [] { R.id }},
                    { W.dataSource,       new [] { R.id }},
                    { W.embedBold,        new [] { R.id }},
                    { W.embedBoldItalic,  new [] { R.id }},
                    { W.embedItalic,      new [] { R.id }},
                    { W.embedRegular,     new [] { R.id }},
                    { W.footerReference,  new [] { R.id }},
                    { W.headerReference,  new [] { R.id }},
                    { W.headerSource,     new [] { R.id }},
                    { W.hyperlink,        new [] { R.id }},
                    { W.printerSettings,  new [] { R.id }},
                    { W.recipientData,    new [] { R.id }},  // Mail merge, not required
                    { W.saveThroughXslt,  new [] { R.id }},
                    { W.sourceFileName,   new [] { R.id }},  // Framesets, not required
                    { W.src,              new [] { R.id }},  // Mail merge, not required
                    { W.subDoc,           new [] { R.id }},  // Sub documents, not required
                    //{ w14:contentPart,  new [] { R.id }},
                    { WNE.toolbarData,    new [] { R.id }},
                };
        }

        private static void CopySpecifiedCustomXmlParts(WordprocessingDocument sourceDocument, WordprocessingDocument output, DocumentBuilderSettings settings)
        {
            if (settings.CustomXmlGuidList == null || !settings.CustomXmlGuidList.Any())
                return;

            foreach (CustomXmlPart customXmlPart in sourceDocument.MainDocumentPart.CustomXmlParts)
            {
                OpenXmlPart propertyPart = customXmlPart
                    .Parts
                    .Select(p => p.OpenXmlPart)
                    .Where(p => p.ContentType == "application/vnd.openxmlformats-officedocument.customXmlProperties+xml")
                    .FirstOrDefault();
                if (propertyPart != null)
                {
                    XDocument propertyPartDoc = propertyPart.GetXDocument();
#if false
        At various locations in Open-Xml-PowerTools, you will find examples of Open XML markup that is associated with code associated with
        querying or generating that markup.  This is an example of the Custom XML Properties part.

<ds:datastoreItem ds:itemID="{1337A0C2-E6EE-4612-ACA5-E0E5A513381D}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">
  <ds:schemaRefs />
</ds:datastoreItem>
#endif
                    var itemID = (string)propertyPartDoc.Root.Attribute(DS.itemID);
                    if (itemID != null)
                    {
                        itemID = itemID.Trim('{', '}');
                        if (settings.CustomXmlGuidList.Contains(itemID))
                        {
                            CustomXmlPart newPart = output.MainDocumentPart.AddCustomXmlPart(customXmlPart.ContentType);
                            newPart.GetXDocument().Add(customXmlPart.GetXDocument().Root);
                            foreach (OpenXmlPart propPart in customXmlPart.Parts.Select(p => p.OpenXmlPart))
                            {
                                CustomXmlPropertiesPart newPropPart = newPart.AddNewPart<CustomXmlPropertiesPart>();
                                newPropPart.GetXDocument().Add(propPart.GetXDocument().Root);
                            }
                        }
                    }
                }
            }
        }

        private static void RemoveHeadersAndFootersFromSections(WordprocessingDocument doc)
        {
            var mdXDoc = doc.MainDocumentPart.GetXDocument();
            var sections = mdXDoc.Descendants(W.sectPr).ToList();
            foreach (var sect in sections)
            {
                sect.Elements(W.headerReference).Remove();
                sect.Elements(W.footerReference).Remove();
            }
            doc.MainDocumentPart.PutXDocument();
        }

        private class CachedHeaderFooter
        {
            public XName Ref;
            public string Type;
            public string CachedPartRid;
        };

        private static void ProcessSectionsForLinkToPreviousHeadersAndFooters(WordprocessingDocument doc)
        {
            CachedHeaderFooter[] cachedHeaderFooter = new[]
            {
                new CachedHeaderFooter() { Ref = W.headerReference, Type = "first" },
                new CachedHeaderFooter() { Ref = W.headerReference, Type = "even" },
                new CachedHeaderFooter() { Ref = W.headerReference, Type = "default" },
                new CachedHeaderFooter() { Ref = W.footerReference, Type = "first" },
                new CachedHeaderFooter() { Ref = W.footerReference, Type = "even" },
                new CachedHeaderFooter() { Ref = W.footerReference, Type = "default" },
            };

            var mdXDoc = doc.MainDocumentPart.GetXDocument();
            var sections = mdXDoc.Descendants(W.sectPr).ToList();
            var firstSection = true;
            foreach (var sect in sections)
            {
                if (firstSection)
                {
                    var headerFirst = FindReference(sect, W.headerReference, "first");
                    var headerDefault = FindReference(sect, W.headerReference, "default");
                    var headerEven = FindReference(sect, W.headerReference, "even");
                    var footerFirst = FindReference(sect, W.footerReference, "first");
                    var footerDefault = FindReference(sect, W.footerReference, "default");
                    var footerEven = FindReference(sect, W.footerReference, "even");

                    if (headerEven == null)
                    {
                        if (headerDefault != null)
                            AddReferenceToExistingHeaderOrFooter(doc.MainDocumentPart, sect, (string)headerDefault.Attribute(R.id), W.headerReference, "even");
                        else
                            InitEmptyHeaderOrFooter(doc.MainDocumentPart, sect, W.headerReference, "even");
                    }

                    if (headerFirst == null)
                    {
                        if (headerDefault != null)
                            AddReferenceToExistingHeaderOrFooter(doc.MainDocumentPart, sect, (string)headerDefault.Attribute(R.id), W.headerReference, "first");
                        else
                            InitEmptyHeaderOrFooter(doc.MainDocumentPart, sect, W.headerReference, "first");
                    }

                    if (footerEven == null)
                    {
                        if (footerDefault != null)
                            AddReferenceToExistingHeaderOrFooter(doc.MainDocumentPart, sect, (string)footerDefault.Attribute(R.id), W.footerReference, "even");
                        else
                            InitEmptyHeaderOrFooter(doc.MainDocumentPart, sect, W.footerReference, "even");
                    }

                    if (footerFirst == null)
                    {
                        if (footerDefault != null)
                            AddReferenceToExistingHeaderOrFooter(doc.MainDocumentPart, sect, (string)footerDefault.Attribute(R.id), W.footerReference, "first");
                        else
                            InitEmptyHeaderOrFooter(doc.MainDocumentPart, sect, W.footerReference, "first");
                    }

                    foreach (var hf in cachedHeaderFooter)
                    {
                        if (sect.Elements(hf.Ref).FirstOrDefault(z => (string)z.Attribute(W.type) == hf.Type) == null)
                            InitEmptyHeaderOrFooter(doc.MainDocumentPart, sect, hf.Ref, hf.Type);
                        var reference = sect.Elements(hf.Ref).FirstOrDefault(z => (string)z.Attribute(W.type) == hf.Type);
                        if (reference == null)
                            throw new OpenXmlPowerToolsException("Internal error");
                        hf.CachedPartRid = (string)reference.Attribute(R.id);
                    }
                    firstSection = false;
                    continue;
                }
                else
                {
                    CopyOrCacheHeaderOrFooter(doc, cachedHeaderFooter, sect, W.headerReference, "first");
                    CopyOrCacheHeaderOrFooter(doc, cachedHeaderFooter, sect, W.headerReference, "even");
                    CopyOrCacheHeaderOrFooter(doc, cachedHeaderFooter, sect, W.headerReference, "default");
                    CopyOrCacheHeaderOrFooter(doc, cachedHeaderFooter, sect, W.footerReference, "first");
                    CopyOrCacheHeaderOrFooter(doc, cachedHeaderFooter, sect, W.footerReference, "even");
                    CopyOrCacheHeaderOrFooter(doc, cachedHeaderFooter, sect, W.footerReference, "default");
                }
            }
            doc.MainDocumentPart.PutXDocument();
        }

        private static void CopyOrCacheHeaderOrFooter(WordprocessingDocument doc, CachedHeaderFooter[] cachedHeaderFooter, XElement sect, XName referenceXName, string type)
        {
            var referenceElement = FindReference(sect, referenceXName, type);
            if (referenceElement == null)
            {
                var cachedPartRid = cachedHeaderFooter.FirstOrDefault(z => z.Ref == referenceXName && z.Type == type).CachedPartRid;
                AddReferenceToExistingHeaderOrFooter(doc.MainDocumentPart, sect, cachedPartRid, referenceXName, type);
            }
            else
            {
                var cachedPart = cachedHeaderFooter.FirstOrDefault(z => z.Ref == referenceXName && z.Type == type);
                cachedPart.CachedPartRid = (string)referenceElement.Attribute(R.id);
            }
        }

        private static XElement FindReference(XElement sect, XName reference, string type)
        {
            return sect.Elements(reference).FirstOrDefault(z =>
                {
                    return (string)z.Attribute(W.type) == type;
                });
        }

        private static void AddReferenceToExistingHeaderOrFooter(MainDocumentPart mainDocPart, XElement sect, string rId, XName reference, string toType)
        {
            if (reference == W.headerReference)
            {
                var referenceToAdd = new XElement(W.headerReference,
                    new XAttribute(W.type, toType),
                    new XAttribute(R.id, rId));
                sect.AddFirst(referenceToAdd);
            }
            else
            {
                var referenceToAdd = new XElement(W.footerReference,
                    new XAttribute(W.type, toType),
                    new XAttribute(R.id, rId));
                sect.AddFirst(referenceToAdd);
            }
        }

        private static void InitEmptyHeaderOrFooter(MainDocumentPart mainDocPart, XElement sect, XName referenceXName, string toType)
        {
            XDocument xDoc = null;
            if (referenceXName == W.headerReference)
            {
                xDoc = XDocument.Parse(
                    @"<?xml version='1.0' encoding='utf-8' standalone='yes'?>
                    <w:hdr xmlns:wpc='http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas'
                           xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006'
                           xmlns:o='urn:schemas-microsoft-com:office:office'
                           xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
                           xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'
                           xmlns:v='urn:schemas-microsoft-com:vml'
                           xmlns:wp14='http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing'
                           xmlns:wp='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
                           xmlns:w10='urn:schemas-microsoft-com:office:word'
                           xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                           xmlns:w14='http://schemas.microsoft.com/office/word/2010/wordml'
                           xmlns:w15='http://schemas.microsoft.com/office/word/2012/wordml'
                           xmlns:wpg='http://schemas.microsoft.com/office/word/2010/wordprocessingGroup'
                           xmlns:wpi='http://schemas.microsoft.com/office/word/2010/wordprocessingInk'
                           xmlns:wne='http://schemas.microsoft.com/office/word/2006/wordml'
                           xmlns:wps='http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
                           mc:Ignorable='w14 w15 wp14'>
                      <w:p>
                        <w:pPr>
                          <w:pStyle w:val='Header' />
                        </w:pPr>
                        <w:r>
                          <w:t></w:t>
                        </w:r>
                      </w:p>
                    </w:hdr>");
                var newHeaderPart = mainDocPart.AddNewPart<HeaderPart>();
                newHeaderPart.PutXDocument(xDoc);
                var referenceToAdd = new XElement(W.headerReference,
                    new XAttribute(W.type, toType),
                    new XAttribute(R.id, mainDocPart.GetIdOfPart(newHeaderPart)));
                sect.AddFirst(referenceToAdd);
            }
            else
            {
                xDoc = XDocument.Parse(@"<?xml version='1.0' encoding='utf-8' standalone='yes'?>
                    <w:ftr xmlns:wpc='http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas'
                           xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006'
                           xmlns:o='urn:schemas-microsoft-com:office:office'
                           xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
                           xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'
                           xmlns:v='urn:schemas-microsoft-com:vml'
                           xmlns:wp14='http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing'
                           xmlns:wp='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
                           xmlns:w10='urn:schemas-microsoft-com:office:word'
                           xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                           xmlns:w14='http://schemas.microsoft.com/office/word/2010/wordml'
                           xmlns:w15='http://schemas.microsoft.com/office/word/2012/wordml'
                           xmlns:wpg='http://schemas.microsoft.com/office/word/2010/wordprocessingGroup'
                           xmlns:wpi='http://schemas.microsoft.com/office/word/2010/wordprocessingInk'
                           xmlns:wne='http://schemas.microsoft.com/office/word/2006/wordml'
                           xmlns:wps='http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
                           mc:Ignorable='w14 w15 wp14'>
                      <w:p>
                        <w:pPr>
                          <w:pStyle w:val='Footer' />
                        </w:pPr>
                        <w:r>
                          <w:t></w:t>
                        </w:r>
                      </w:p>
                    </w:ftr>");
                var newFooterPart = mainDocPart.AddNewPart<FooterPart>();
                newFooterPart.PutXDocument(xDoc);
                var referenceToAdd = new XElement(W.footerReference,
                    new XAttribute(W.type, toType),
                    new XAttribute(R.id, mainDocPart.GetIdOfPart(newFooterPart)));
                sect.AddFirst(referenceToAdd);
            }
        }

        private static void TestPartForUnsupportedContent(OpenXmlPart part, int sourceNumber)
        {
            XNamespace[] obsoleteNamespaces = new[]
                {
                    XNamespace.Get("http://schemas.microsoft.com/office/word/2007/5/30/wordml"),
                    XNamespace.Get("http://schemas.microsoft.com/office/word/2008/9/16/wordprocessingDrawing"),
                    XNamespace.Get("http://schemas.microsoft.com/office/word/2009/2/wordml"),
                };
            XDocument xDoc = part.GetXDocument();
            XElement invalidElement = xDoc.Descendants()
                .FirstOrDefault(d =>
                    {
                        bool b = d.Name == W.subDoc ||
                            d.Name == W.control ||
                            d.Name == W.altChunk ||
                            d.Name.LocalName == "contentPart" ||
                            obsoleteNamespaces.Contains(d.Name.Namespace);
                        bool b2 = b ||
                            d.Attributes().Any(a => obsoleteNamespaces.Contains(a.Name.Namespace));
                        return b2;
                    });
            if (invalidElement != null)
            {
                if (invalidElement.Name == W.subDoc)
                    throw new DocumentBuilderException(String.Format("Source {0} is unsupported document - contains sub document",
                        sourceNumber));
                if (invalidElement.Name == W.control)
                    throw new DocumentBuilderException(String.Format("Source {0} is unsupported document - contains ActiveX controls",
                        sourceNumber));
                if (invalidElement.Name == W.altChunk)
                    throw new DocumentBuilderException(String.Format("Source {0} is unsupported document - contains altChunk",
                        sourceNumber));
                if (invalidElement.Name.LocalName == "contentPart")
                    throw new DocumentBuilderException(String.Format("Source {0} is unsupported document - contains contentPart content",
                        sourceNumber));
                if (obsoleteNamespaces.Contains(invalidElement.Name.Namespace) ||
                    invalidElement.Attributes().Any(a => obsoleteNamespaces.Contains(a.Name.Namespace)))
                    throw new DocumentBuilderException(String.Format("Source {0} is unsupported document - contains obsolete namespace",
                        sourceNumber));
            }
        }

        //What does not work:
        //- sub docs
        //- bidi text appears to work but has not been tested
        //- languages other than en-us appear to work but have not been tested
        //- documents with activex controls
        //- mail merge source documents (look for dataSource in settings)
        //- documents with ink
        //- documents with frame sets and frames
        private static void TestForUnsupportedDocument(WordprocessingDocument doc, int sourceNumber)
        {
            if (doc.MainDocumentPart.GetXDocument().Root == null)
                throw new DocumentBuilderException(string.Format("Source {0} is an invalid document - MainDocumentPart contains no content.", sourceNumber));

            if ((string)doc.MainDocumentPart.GetXDocument().Root.Name.NamespaceName == "http://purl.oclc.org/ooxml/wordprocessingml/main")
                throw new DocumentBuilderException(string.Format("Source {0} is saved in strict mode, not supported", sourceNumber));

            // note: if ever want to support section changes, need to address the code that rationalizes headers and footers, propagating to sections that inherit headers/footers from prev section
            foreach (var d in doc.MainDocumentPart.GetXDocument().Descendants())
            {
                if (d.Name == W.sectPrChange)
                    throw new DocumentBuilderException(string.Format("Source {0} contains section changes (w:sectPrChange), not supported", sourceNumber));

                // note: if ever want to support Open-Xml-PowerTools attributes, need to make sure that all attributes are propagated in all cases
                //if (d.Name.Namespace == PtOpenXml.ptOpenXml ||
                //    d.Name.Namespace == PtOpenXml.pt)
                //    throw new DocumentBuilderException(string.Format("Source {0} contains Open-Xml-PowerTools markup, not supported", sourceNumber));
                //if (d.Attributes().Any(a => a.Name.Namespace == PtOpenXml.ptOpenXml || a.Name.Namespace == PtOpenXml.pt))
                //    throw new DocumentBuilderException(string.Format("Source {0} contains Open-Xml-PowerTools markup, not supported", sourceNumber));
            }

            TestPartForUnsupportedContent(doc.MainDocumentPart, sourceNumber);
            foreach (var hdr in doc.MainDocumentPart.HeaderParts)
                TestPartForUnsupportedContent(hdr, sourceNumber);
            foreach (var ftr in doc.MainDocumentPart.FooterParts)
                TestPartForUnsupportedContent(ftr, sourceNumber);
            if (doc.MainDocumentPart.FootnotesPart != null)
                TestPartForUnsupportedContent(doc.MainDocumentPart.FootnotesPart, sourceNumber);
            if (doc.MainDocumentPart.EndnotesPart != null)
                TestPartForUnsupportedContent(doc.MainDocumentPart.EndnotesPart, sourceNumber);

            if (doc.MainDocumentPart.DocumentSettingsPart != null &&
                doc.MainDocumentPart.DocumentSettingsPart.GetXDocument().Descendants().Any(d => d.Name == W.src ||
                d.Name == W.recipientData || d.Name == W.mailMerge))
                throw new DocumentBuilderException(String.Format("Source {0} is unsupported document - contains Mail Merge content",
                    sourceNumber));
            if (doc.MainDocumentPart.WebSettingsPart != null &&
                doc.MainDocumentPart.WebSettingsPart.GetXDocument().Descendants().Any(d => d.Name == W.frameset))
                throw new DocumentBuilderException(String.Format("Source {0} is unsupported document - contains a frameset", sourceNumber));
            var numberingElements = doc.MainDocumentPart
                .GetXDocument()
                .Descendants(W.numPr)
                .Where(n =>
                    {
                        bool zeroId = (int?)n.Attribute(W.id) == 0;
                        bool hasChildInsId = n.Elements(W.ins).Any();
                        if (zeroId || hasChildInsId)
                            return false;
                        return true;
                    })
                .ToList();
            if (numberingElements.Any() &&
                doc.MainDocumentPart.NumberingDefinitionsPart == null)
                throw new DocumentBuilderException(String.Format(
                    "Source {0} is invalid document - contains numbering markup but no numbering part", sourceNumber));
        }

        private static void FixUpSectionProperties(WordprocessingDocument newDocument)
        {
            XDocument mainDocumentXDoc = newDocument.MainDocumentPart.GetXDocument();
            mainDocumentXDoc.Declaration.Standalone = Yes;
            mainDocumentXDoc.Declaration.Encoding = Utf8;
            XElement body = mainDocumentXDoc.Root.Element(W.body);
            var sectionPropertiesToMove = body
                .Elements()
                .Take(body.Elements().Count() - 1)
                .Where(e => e.Name == W.sectPr)
                .ToList();
            foreach (var s in sectionPropertiesToMove)
            {
                var p = s.SiblingsBeforeSelfReverseDocumentOrder().First();
                if (p.Element(W.pPr) == null)
                    p.AddFirst(new XElement(W.pPr));
                p.Element(W.pPr).Add(s);
            }
            foreach (var s in sectionPropertiesToMove)
                s.Remove();
        }

        private static void AddSectionAndDependencies(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument,
            XElement sectionMarkup, List<ImageData> images)
        {
            var headerReferences = sectionMarkup.Elements(W.headerReference);
            foreach (var headerReference in headerReferences)
            {
                string oldRid = headerReference.Attribute(R.id).Value;
                HeaderPart oldHeaderPart = null;
                try
                {
                    oldHeaderPart = (HeaderPart)sourceDocument.MainDocumentPart.GetPartById(oldRid);
                }
                catch (ArgumentOutOfRangeException)
                {
                    var message = string.Format("ArgumentOutOfRangeException, attempting to get header rId={0}", oldRid);
                    throw new OpenXmlPowerToolsException(message);
                }
                XDocument oldHeaderXDoc = oldHeaderPart.GetXDocument();
                if (oldHeaderXDoc != null && oldHeaderXDoc.Root != null)
                    CopyNumbering(sourceDocument, newDocument, new[] { oldHeaderXDoc.Root }, images);
                HeaderPart newHeaderPart = newDocument.MainDocumentPart.AddNewPart<HeaderPart>();
                XDocument newHeaderXDoc = newHeaderPart.GetXDocument();
                newHeaderXDoc.Declaration.Standalone = Yes;
                newHeaderXDoc.Declaration.Encoding = Utf8;
                newHeaderXDoc.Add(oldHeaderXDoc.Root);
                headerReference.Attribute(R.id).Value = newDocument.MainDocumentPart.GetIdOfPart(newHeaderPart);
                AddRelationships(oldHeaderPart, newHeaderPart, new[] { newHeaderXDoc.Root });
                CopyRelatedPartsForContentParts(oldHeaderPart, newHeaderPart, new[] { newHeaderXDoc.Root }, images);
            }

            var footerReferences = sectionMarkup.Elements(W.footerReference);
            foreach (var footerReference in footerReferences)
            {
                string oldRid = footerReference.Attribute(R.id).Value;
                var oldFooterPart2 = sourceDocument.MainDocumentPart.GetPartById(oldRid);
                if (!(oldFooterPart2 is FooterPart))
                    throw new DocumentBuilderException("Invalid document - invalid footer part.");

                FooterPart oldFooterPart = (FooterPart)oldFooterPart2;
                XDocument oldFooterXDoc = oldFooterPart.GetXDocument();
                if (oldFooterXDoc != null && oldFooterXDoc.Root != null)
                    CopyNumbering(sourceDocument, newDocument, new[] { oldFooterXDoc.Root }, images);
                FooterPart newFooterPart = newDocument.MainDocumentPart.AddNewPart<FooterPart>();
                XDocument newFooterXDoc = newFooterPart.GetXDocument();
                newFooterXDoc.Declaration.Standalone = Yes;
                newFooterXDoc.Declaration.Encoding = Utf8;
                newFooterXDoc.Add(oldFooterXDoc.Root);
                footerReference.Attribute(R.id).Value = newDocument.MainDocumentPart.GetIdOfPart(newFooterPart);
                AddRelationships(oldFooterPart, newFooterPart, new[] { newFooterXDoc.Root });
                CopyRelatedPartsForContentParts(oldFooterPart, newFooterPart, new[] { newFooterXDoc.Root }, images);
            }
        }

        private static void MergeStyles(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument, XDocument fromStyles, XDocument toStyles, IEnumerable<XElement> newContent)
        {
#if MergeStylesWithSameNames
            var newIds = new Dictionary<string, string>();
#endif
            if (fromStyles.Root == null)
                return;

            foreach (XElement style in fromStyles.Root.Elements(W.style))
            {
                var fromId = (string)style.Attribute(W.styleId);
                var fromName = (string)style.Elements(W.name).Attributes(W.val).FirstOrDefault();

                var toStyle = toStyles
                    .Root
                    .Elements(W.style)
                    .FirstOrDefault(st => (string)st.Elements(W.name).Attributes(W.val).FirstOrDefault() == fromName);

                if (toStyle == null)
                {
#if MergeStylesWithSameNames
                    var linkElement = style.Element(W.link);
                    string linkedId;
                    if (linkElement != null && newIds.TryGetValue(linkElement.Attribute(W.val).Value, out linkedId))
                    {
                        var linkedStyle = toStyles.Root.Elements(W.style)
                            .First(o => o.Attribute(W.styleId).Value == linkedId);
                        if (linkedStyle.Element(W.link) != null)
                            newIds.Add(fromId, linkedStyle.Element(W.link).Attribute(W.val).Value);
                        continue;
                    }

                    //string name = (string)style.Elements(W.name).Attributes(W.val).FirstOrDefault();
                    //var namedStyle = toStyles
                    //    .Root
                    //    .Elements(W.style)
                    //    .Where(st => st.Element(W.name) != null)
                    //    .FirstOrDefault(o => (string)o.Element(W.name).Attribute(W.val) == name);
                    //if (namedStyle != null)
                    //{
                    //    if (! newIds.ContainsKey(fromId))
                    //        newIds.Add(fromId, namedStyle.Attribute(W.styleId).Value);
                    //    continue;
                    //}
#endif

                    int number = 1;
                    int abstractNumber = 0;
                    XDocument oldNumbering = null;
                    XDocument newNumbering = null;
                    foreach (XElement numReference in style.Descendants(W.numPr))
                    {
                        XElement idElement = numReference.Descendants(W.numId).FirstOrDefault();
                        if (idElement != null)
                        {
                            if (oldNumbering == null)
                            {
                                if (sourceDocument.MainDocumentPart.NumberingDefinitionsPart != null)
                                    oldNumbering = sourceDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                                else
                                {
                                    oldNumbering = new XDocument();
                                    oldNumbering.Declaration = new XDeclaration(OnePointZero, Utf8, Yes);
                                    oldNumbering.Add(new XElement(W.numbering, NamespaceAttributes));
                                }
                            }
                            if (newNumbering == null)
                            {
                                if (newDocument.MainDocumentPart.NumberingDefinitionsPart != null)
                                {
                                    newNumbering = newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                                    newNumbering.Declaration.Standalone = Yes;
                                    newNumbering.Declaration.Encoding = Utf8;
                                    var numIds = newNumbering
                                        .Root
                                        .Elements(W.num)
                                        .Select(f => (int)f.Attribute(W.numId));
                                    if (numIds.Any())
                                        number = numIds.Max() + 1;
                                    numIds = newNumbering
                                        .Root
                                        .Elements(W.abstractNum)
                                        .Select(f => (int)f.Attribute(W.abstractNumId));
                                    if (numIds.Any())
                                        abstractNumber = numIds.Max() + 1;
                                }
                                else
                                {
                                    newDocument.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                                    newNumbering = newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                                    newNumbering.Declaration.Standalone = Yes;
                                    newNumbering.Declaration.Encoding = Utf8;
                                    newNumbering.Add(new XElement(W.numbering, NamespaceAttributes));
                                }
                            }
                            string numId = idElement.Attribute(W.val).Value;
                            if (numId != "0")
                            {
                                XElement element = oldNumbering
                                    .Descendants()
                                    .Elements(W.num)
                                    .Where(p => ((string)p.Attribute(W.numId)) == numId)
                                    .FirstOrDefault();

                                // Copy abstract numbering element, if necessary (use matching NSID)
                                string abstractNumId = string.Empty;
                                if (element != null)
                                {
                                    abstractNumId = element
                                       .Elements(W.abstractNumId)
                                       .First()
                                       .Attribute(W.val)
                                       .Value;

                                    XElement abstractElement = oldNumbering
                                        .Descendants()
                                        .Elements(W.abstractNum)
                                        .Where(p => ((string)p.Attribute(W.abstractNumId)) == abstractNumId)
                                        .FirstOrDefault();
                                    string abstractNSID = string.Empty;
                                    if (abstractElement != null)
                                    {
                                        XElement nsidElement = abstractElement
                                            .Element(W.nsid);
                                        abstractNSID = null;
                                        if (nsidElement != null)
                                            abstractNSID = (string)nsidElement
                                                .Attribute(W.val);

                                        XElement newAbstractElement = newNumbering
                                            .Descendants()
                                            .Elements(W.abstractNum)
                                            .Where(e => e.Annotation<FromPreviousSourceSemaphore>() == null)
                                            .Where(p =>
                                            {
                                                var thisNsidElement = p.Element(W.nsid);
                                                if (thisNsidElement == null)
                                                    return false;
                                                return (string)thisNsidElement.Attribute(W.val) == abstractNSID;
                                            })
                                            .FirstOrDefault();
                                        if (newAbstractElement == null)
                                        {
                                            newAbstractElement = new XElement(abstractElement);
                                            newAbstractElement.Attribute(W.abstractNumId).Value = abstractNumber.ToString();
                                            abstractNumber++;
                                            if (newNumbering.Root.Elements(W.abstractNum).Any())
                                                newNumbering.Root.Elements(W.abstractNum).Last().AddAfterSelf(newAbstractElement);
                                            else
                                                newNumbering.Root.Add(newAbstractElement);

                                            foreach (XElement pictId in newAbstractElement.Descendants(W.lvlPicBulletId))
                                            {
                                                string bulletId = (string)pictId.Attribute(W.val);
                                                XElement numPicBullet = oldNumbering
                                                    .Descendants(W.numPicBullet)
                                                    .FirstOrDefault(d => (string)d.Attribute(W.numPicBulletId) == bulletId);
                                                int maxNumPicBulletId = new int[] { -1 }.Concat(
                                                    newNumbering.Descendants(W.numPicBullet)
                                                    .Attributes(W.numPicBulletId)
                                                    .Select(a => (int)a))
                                                    .Max() + 1;
                                                XElement newNumPicBullet = new XElement(numPicBullet);
                                                newNumPicBullet.Attribute(W.numPicBulletId).Value = maxNumPicBulletId.ToString();
                                                pictId.Attribute(W.val).Value = maxNumPicBulletId.ToString();
                                                newNumbering.Root.AddFirst(newNumPicBullet);
                                            }
                                        }
                                        string newAbstractId = newAbstractElement.Attribute(W.abstractNumId).Value;

                                        // Copy numbering element, if necessary (use matching element with no overrides)
                                        XElement newElement = null;
                                        if (!element.Elements(W.lvlOverride).Any())
                                            newElement = newNumbering
                                                .Descendants()
                                                .Elements(W.num)
                                                .Where(p => !p.Elements(W.lvlOverride).Any() &&
                                                    ((string)p.Elements(W.abstractNumId).First().Attribute(W.val)) == newAbstractId)
                                                .FirstOrDefault();
                                        if (newElement == null)
                                        {
                                            newElement = new XElement(element);
                                            newElement
                                                .Elements(W.abstractNumId)
                                                .First()
                                                .Attribute(W.val).Value = newAbstractId;
                                            newElement.Attribute(W.numId).Value = number.ToString();
                                            number++;
                                            newNumbering.Root.Add(newElement);
                                        }
                                        idElement.Attribute(W.val).Value = newElement.Attribute(W.numId).Value;
                                    }
                                }
                            }
                        }
                    }

                    var newStyle = new XElement(style);
                    // get rid of anything not in the w: namespace
                    newStyle.Descendants().Where(d => d.Name.NamespaceName != W.w).Remove();
                    newStyle.Descendants().Attributes().Where(d => d.Name.NamespaceName != W.w).Remove();
                    toStyles.Root.Add(newStyle);
                }
                else
                {
                    var toId = (string)toStyle.Attribute(W.styleId);
                    if (fromId != toId)
                    {
                        if (! newIds.ContainsKey(fromId))
                            newIds.Add(fromId, toId);
                    }
                }
            }

#if MergeStylesWithSameNames
            if (newIds.Count > 0)
            {
                foreach (var style in toStyles
                    .Root
                    .Elements(W.style))
                {
                    ConvertToNewId(style.Element(W.basedOn), newIds);
                    ConvertToNewId(style.Element(W.next), newIds);
                }

                foreach (var item in newContent.DescendantsAndSelf()
                    .Where(d => d.Name == W.pStyle ||
                                d.Name == W.rStyle ||
                                d.Name == W.tblStyle))
                {
                    ConvertToNewId(item, newIds);
                }

                if (newDocument.MainDocumentPart.NumberingDefinitionsPart != null)
                {
                    var newNumbering = newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                    ConvertNumberingPartToNewIds(newNumbering, newIds);
                }

                // Convert source document, since numberings will be copied over after styles.
                if (sourceDocument.MainDocumentPart.NumberingDefinitionsPart != null)
                {
                    var sourceNumbering = sourceDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                    ConvertNumberingPartToNewIds(sourceNumbering, newIds);
                }
            }
#endif
        }

        private static void MergeLatentStyles(XDocument fromStyles, XDocument toStyles)
        {
            var fromLatentStyles = fromStyles.Descendants(W.latentStyles).FirstOrDefault();
            if (fromLatentStyles == null)
                return;

            var toLatentStyles = toStyles.Descendants(W.latentStyles).FirstOrDefault();
            if (toLatentStyles == null)
            {
                var newLatentStylesElement = new XElement(W.latentStyles,
                    fromLatentStyles.Attributes());
                var globalDefaults = toStyles
                    .Descendants(W.docDefaults)
                    .FirstOrDefault();
                if (globalDefaults == null)
                {
                    var firstStyle = toStyles
                        .Root
                        .Elements(W.style)
                        .FirstOrDefault();
                    if (firstStyle == null)
                        toStyles.Root.Add(newLatentStylesElement);
                    else
                        firstStyle.AddBeforeSelf(newLatentStylesElement);
                }
                else
                    globalDefaults.AddAfterSelf(newLatentStylesElement);
            }
            toLatentStyles = toStyles.Descendants(W.latentStyles).FirstOrDefault();
            if (toLatentStyles == null)
                throw new OpenXmlPowerToolsException("Internal error");

            var toStylesHash = new HashSet<string>();
            foreach (var lse in toLatentStyles.Elements(W.lsdException))
                toStylesHash.Add((string)lse.Attribute(W.name));

            foreach (var fls in fromLatentStyles.Elements(W.lsdException))
            {
                var name = (string)fls.Attribute(W.name);
                if (toStylesHash.Contains(name))
                    continue;
                toLatentStyles.Add(fls);
                toStylesHash.Add(name);
            }

            var count = toLatentStyles
                .Elements(W.lsdException)
                .Count();

            toLatentStyles.SetAttributeValue(W.count, count);
        }

        private static void MergeDocDefaultStyles(XDocument xDocument, XDocument newXDoc)
        {
            var docDefaultStyles = xDocument.Descendants(W.docDefaults);
            foreach (var docDefaultStyle in docDefaultStyles)
            {
                newXDoc.Root.Add(docDefaultStyle);
            }
        }

#if MergeStylesWithSameNames
        private static void ConvertToNewId(XElement element, Dictionary<string, string> newIds)
        {
            if (element == null)
                return;

            var valueAttribute = element.Attribute(W.val);
            string newId;
            if (newIds.TryGetValue(valueAttribute.Value, out newId))
            {
                valueAttribute.Value = newId;
            }
        }

        private static void ConvertNumberingPartToNewIds(XDocument newNumbering, Dictionary<string, string> newIds)
        {
            foreach (var abstractNum in newNumbering
                .Root
                .Elements(W.abstractNum))
            {
                ConvertToNewId(abstractNum.Element(W.styleLink), newIds);
                ConvertToNewId(abstractNum.Element(W.numStyleLink), newIds);
            }

            foreach (var item in newNumbering
                .Descendants()
                .Where(d => d.Name == W.pStyle ||
                            d.Name == W.rStyle ||
                            d.Name == W.tblStyle))
            {
                ConvertToNewId(item, newIds);
            }
        }
#endif

        private static void MergeFontTables(XDocument fromFontTable, XDocument toFontTable)
        {
            foreach (XElement font in fromFontTable.Root.Elements(W.font))
            {
                string name = font.Attribute(W.name).Value;
                if (toFontTable
                    .Root
                    .Elements(W.font)
                    .Where(o => o.Attribute(W.name).Value == name)
                    .Count() == 0)
                    toFontTable.Root.Add(new XElement(font));
            }
        }

        private static void CopyStylesAndFonts(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument,
            IEnumerable<XElement> newContent)
        {
            // Copy all styles to the new document
            if (sourceDocument.MainDocumentPart.StyleDefinitionsPart != null)
            {
                XDocument oldStyles = sourceDocument.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
                if (newDocument.MainDocumentPart.StyleDefinitionsPart == null)
                {
                    newDocument.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                    XDocument newStyles = newDocument.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
                    newStyles.Declaration.Standalone = Yes;
                    newStyles.Declaration.Encoding = Utf8;
                    newStyles.Add(oldStyles.Root);
                }
                else
                {
                    XDocument newStyles = newDocument.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
                    MergeStyles(sourceDocument, newDocument, oldStyles, newStyles, newContent);
                    MergeLatentStyles(oldStyles, newStyles);
                }
            }

            // Copy all styles with effects to the new document
            if (sourceDocument.MainDocumentPart.StylesWithEffectsPart != null)
            {
                XDocument oldStyles = sourceDocument.MainDocumentPart.StylesWithEffectsPart.GetXDocument();
                if (newDocument.MainDocumentPart.StylesWithEffectsPart == null)
                {
                    newDocument.MainDocumentPart.AddNewPart<StylesWithEffectsPart>();
                    XDocument newStyles = newDocument.MainDocumentPart.StylesWithEffectsPart.GetXDocument();
                    newStyles.Declaration.Standalone = Yes;
                    newStyles.Declaration.Encoding = Utf8;
                    newStyles.Add(oldStyles.Root);
                }
                else
                {
                    XDocument newStyles = newDocument.MainDocumentPart.StylesWithEffectsPart.GetXDocument();
                    MergeStyles(sourceDocument, newDocument, oldStyles, newStyles, newContent);
                    MergeLatentStyles(oldStyles, newStyles);
                }
            }

            // Copy fontTable to the new document
            if (sourceDocument.MainDocumentPart.FontTablePart != null)
            {
                XDocument oldFontTable = sourceDocument.MainDocumentPart.FontTablePart.GetXDocument();
                if (newDocument.MainDocumentPart.FontTablePart == null)
                {
                    newDocument.MainDocumentPart.AddNewPart<FontTablePart>();
                    XDocument newFontTable = newDocument.MainDocumentPart.FontTablePart.GetXDocument();
                    newFontTable.Declaration.Standalone = Yes;
                    newFontTable.Declaration.Encoding = Utf8;
                    newFontTable.Add(oldFontTable.Root);
                }
                else
                {
                    XDocument newFontTable = newDocument.MainDocumentPart.FontTablePart.GetXDocument();
                    MergeFontTables(oldFontTable, newFontTable);
                }
            }
        }

        private static void CopyComments(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument,
            IEnumerable<XElement> newContent, List<ImageData> images)
        {
            Dictionary<int, int> commentIdMap = new Dictionary<int, int>();
            int number = 0;
            XDocument oldComments = null;
            XDocument newComments = null;
            foreach (XElement comment in newContent.DescendantsAndSelf(W.commentReference))
            {
                if (oldComments == null)
                    oldComments = sourceDocument.MainDocumentPart.WordprocessingCommentsPart.GetXDocument();
                if (newComments == null)
                {
                    if (newDocument.MainDocumentPart.WordprocessingCommentsPart != null)
                    {
                        newComments = newDocument.MainDocumentPart.WordprocessingCommentsPart.GetXDocument();
                        newComments.Declaration.Standalone = Yes;
                        newComments.Declaration.Encoding = Utf8;
                        var ids = newComments.Root.Elements(W.comment).Select(f => (int)f.Attribute(W.id));
                        if (ids.Any())
                            number = ids.Max() + 1;
                    }
                    else
                    {
                        newDocument.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
                        newComments = newDocument.MainDocumentPart.WordprocessingCommentsPart.GetXDocument();
                        newComments.Declaration.Standalone = Yes;
                        newComments.Declaration.Encoding = Utf8;
                        newComments.Add(new XElement(W.comments, NamespaceAttributes));
                    }
                }
                int id;
                if (!int.TryParse((string)comment.Attribute(W.id), out id))
                    throw new DocumentBuilderException("Invalid document - invalid comment id");
                XElement element = oldComments
                    .Descendants()
                    .Elements(W.comment)
                    .Where(p => {
                        int thisId;
                        if (! int.TryParse((string)p.Attribute(W.id), out thisId))
                            throw new DocumentBuilderException("Invalid document - invalid comment id");
                        return thisId == id;
                    })
                    .FirstOrDefault();
                if (element == null)
                    throw new DocumentBuilderException("Invalid document - comment reference without associated comment in comments part");
                XElement newElement = new XElement(element);
                newElement.Attribute(W.id).Value = number.ToString();
                newComments.Root.Add(newElement);
                if (! commentIdMap.ContainsKey(id))
                    commentIdMap.Add(id, number);
                number++;
            }
            foreach (var item in newContent.DescendantsAndSelf()
                .Where(d => d.Name == W.commentReference ||
                            d.Name == W.commentRangeStart ||
                            d.Name == W.commentRangeEnd)
                .ToList())
            {
                if (commentIdMap.ContainsKey((int)item.Attribute(W.id)))
                    item.Attribute(W.id).Value = commentIdMap[(int)item.Attribute(W.id)].ToString();
            }
            if (sourceDocument.MainDocumentPart.WordprocessingCommentsPart != null &&
                newDocument.MainDocumentPart.WordprocessingCommentsPart != null)
            {
                AddRelationships(sourceDocument.MainDocumentPart.WordprocessingCommentsPart,
                    newDocument.MainDocumentPart.WordprocessingCommentsPart,
                    new[] { newDocument.MainDocumentPart.WordprocessingCommentsPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(sourceDocument.MainDocumentPart.WordprocessingCommentsPart,
                    newDocument.MainDocumentPart.WordprocessingCommentsPart,
                    new[] { newDocument.MainDocumentPart.WordprocessingCommentsPart.GetXDocument().Root },
                    images);
            }
        }

        private static void AdjustUniqueIds(WordprocessingDocument sourceDocument,
            WordprocessingDocument newDocument, IEnumerable<XElement> newContent)
        {
            // adjust bookmark unique ids
            int maxId = 0;
            if (newDocument.MainDocumentPart.GetXDocument().Descendants(W.bookmarkStart).Any())
                maxId = newDocument.MainDocumentPart.GetXDocument().Descendants(W.bookmarkStart)
                    .Select(d => (int)d.Attribute(W.id)).Max();
            Dictionary<int, int> bookmarkIdMap = new Dictionary<int, int>();
            foreach (var item in newContent.DescendantsAndSelf().Where(bm => bm.Name == W.bookmarkStart ||
                bm.Name == W.bookmarkEnd))
            {
                int id;
                if (!int.TryParse((string)item.Attribute(W.id), out id))
                    throw new DocumentBuilderException("Invalid document - invalid value for bookmark ID");
                if (!bookmarkIdMap.ContainsKey(id))
                    bookmarkIdMap.Add(id, ++maxId);
            }
            foreach (var bookmarkElement in newContent.DescendantsAndSelf().Where(e => e.Name == W.bookmarkStart ||
                e.Name == W.bookmarkEnd))
                bookmarkElement.Attribute(W.id).Value = bookmarkIdMap[(int)bookmarkElement.Attribute(W.id)].ToString();

            // adjust shape unique ids
            // This doesn't work because OLEObjects refer to shapes by ID.
            // Punting on this, because sooner or later, this will be a non-issue.
            //foreach (var item in newContent.DescendantsAndSelf(VML.shape))
            //{
            //    Guid g = Guid.NewGuid();
            //    string s = "R" + g.ToString().Replace("-", "");
            //    item.Attribute(NoNamespace.id).Value = s;
            //}
        }

        private static void AdjustDocPrIds(WordprocessingDocument newDocument)
        {
            int docPrId = 0;
            foreach (var item in newDocument.MainDocumentPart.GetXDocument().Descendants(WP.docPr))
                item.Attribute(NoNamespace.id).Value = (++docPrId).ToString();
            foreach (var header in newDocument.MainDocumentPart.HeaderParts)
                foreach (var item in header.GetXDocument().Descendants(WP.docPr))
                    item.Attribute(NoNamespace.id).Value = (++docPrId).ToString();
            foreach (var footer in newDocument.MainDocumentPart.FooterParts)
                foreach (var item in footer.GetXDocument().Descendants(WP.docPr))
                    item.Attribute(NoNamespace.id).Value = (++docPrId).ToString();
            if (newDocument.MainDocumentPart.FootnotesPart != null)
                foreach (var item in newDocument.MainDocumentPart.FootnotesPart.GetXDocument().Descendants(WP.docPr))
                    item.Attribute(NoNamespace.id).Value = (++docPrId).ToString();
            if (newDocument.MainDocumentPart.EndnotesPart != null)
                foreach (var item in newDocument.MainDocumentPart.EndnotesPart.GetXDocument().Descendants(WP.docPr))
                    item.Attribute(NoNamespace.id).Value = (++docPrId).ToString();
        }

        // This probably doesn't need to be done, except that the Open XML SDK will not validate
        // documents that contain the o:gfxdata attribute.
        private static void RemoveGfxdata(IEnumerable<XElement> newContent)
        {
            newContent.DescendantsAndSelf().Attributes(O.gfxdata).Remove();
        }

        private static object InsertTransform(XNode node, List<XElement> newContent)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Annotation<ReplaceSemaphore>() != null)
                    return newContent;
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => InsertTransform(n, newContent)));
            }
            return node;
        }

        private class ReplaceSemaphore { }

        // Rules for sections
        // - if KeepSections for all documents in the source collection are false, then it takes the section
        //   from the first document.
        // - if you specify true for any document, and if the last section is part of the specified content,
        //   then that section is copied.  If any paragraph in the content has a section, then that section
        //   is copied.
        // - if you specify true for any document, and there are no sections for any paragraphs, then no
        //   sections are copied.
        private static void AppendDocument(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument,
            List<XElement> newContent, bool keepSection, string insertId, List<ImageData> images)
        {
            FixRanges(sourceDocument.MainDocumentPart.GetXDocument(), newContent);
            AddRelationships(sourceDocument.MainDocumentPart, newDocument.MainDocumentPart, newContent);
            CopyRelatedPartsForContentParts(sourceDocument.MainDocumentPart, newDocument.MainDocumentPart,
                newContent, images);

            // Append contents
            XDocument newMainXDoc = newDocument.MainDocumentPart.GetXDocument();
            newMainXDoc.Declaration.Standalone = Yes;
            newMainXDoc.Declaration.Encoding = Utf8;
            if (keepSection == false)
            {
                List<XElement> adjustedContents = newContent.Where(e => e.Name != W.sectPr).ToList();
                adjustedContents.DescendantsAndSelf(W.sectPr).Remove();
                newContent = adjustedContents;
            }
            var listOfSectionProps = newContent.DescendantsAndSelf(W.sectPr).ToList();
            foreach (var sectPr in listOfSectionProps)
                AddSectionAndDependencies(sourceDocument, newDocument, sectPr, images);
            CopyStylesAndFonts(sourceDocument, newDocument, newContent);
            CopyNumbering(sourceDocument, newDocument, newContent, images);
            CopyComments(sourceDocument, newDocument, newContent, images);
            CopyFootnotes(sourceDocument, newDocument, newContent, images);
            CopyEndnotes(sourceDocument, newDocument, newContent, images);
            AdjustUniqueIds(sourceDocument, newDocument, newContent);
            RemoveGfxdata(newContent);
            CopyCustomXmlPartsForDataBoundContentControls(sourceDocument, newDocument, newContent);
            CopyWebExtensions(sourceDocument, newDocument);
            if (insertId != null)
            {
                XElement insertElementToReplace = newMainXDoc
                    .Descendants(PtOpenXml.Insert)
                    .FirstOrDefault(i => (string)i.Attribute(PtOpenXml.Id) == insertId);
                if (insertElementToReplace != null)
                    insertElementToReplace.AddAnnotation(new ReplaceSemaphore());
                newMainXDoc.Element(W.document).ReplaceWith((XElement)InsertTransform(newMainXDoc.Root, newContent));
            }
            else
                newMainXDoc.Root.Element(W.body).Add(newContent);

            if (newMainXDoc.Descendants().Any(d =>
            {
                if (d.Name.Namespace == PtOpenXml.pt || d.Name.Namespace == PtOpenXml.ptOpenXml)
                    return true;
                if (d.Attributes().Any(att => att.Name.Namespace == PtOpenXml.pt || att.Name.Namespace == PtOpenXml.ptOpenXml))
                    return true;
                return false;
            }))
            {
                var root = newMainXDoc.Root;
                if (!root.Attributes().Any(na => na.Value == PtOpenXml.pt.NamespaceName))
                {
                    root.Add(new XAttribute(XNamespace.Xmlns + "pt", PtOpenXml.pt.NamespaceName));
                    AddToIgnorable(root, "pt");
                }
                if (!root.Attributes().Any(na => na.Value == PtOpenXml.ptOpenXml.NamespaceName))
                {
                    root.Add(new XAttribute(XNamespace.Xmlns + "pt14", PtOpenXml.ptOpenXml.NamespaceName));
                    AddToIgnorable(root, "pt14");
                }
            }
        }

        private static void CopyWebExtensions(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument)
        {
            if (sourceDocument.WebExTaskpanesPart != null && newDocument.WebExTaskpanesPart == null)
            {
                newDocument.AddWebExTaskpanesPart();
                newDocument.WebExTaskpanesPart.GetXDocument().Add(sourceDocument.WebExTaskpanesPart.GetXDocument().Root);

                foreach (var sourceWebExtensionPart in sourceDocument.WebExTaskpanesPart.WebExtensionParts)
                {
                    var newWebExtensionpart = newDocument.WebExTaskpanesPart.AddNewPart<WebExtensionPart>(
                        sourceDocument.WebExTaskpanesPart.GetIdOfPart(sourceWebExtensionPart));
                    newWebExtensionpart.GetXDocument().Add(sourceWebExtensionPart.GetXDocument().Root);
                }
            }
        }

        private static void AddToIgnorable(XElement root, string v)
        {
            var ignorable = root.Attribute(MC.Ignorable);
            if (ignorable != null)
            {
                var val = (string)ignorable;
                val = val + " " + v;
                ignorable.Remove();
                root.SetAttributeValue(MC.Ignorable, val);
            }
        }

        /// ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        /// New method to support new functionality
        private static void AppendDocument(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument, OpenXmlPart part,
            List<XElement> newContent, bool keepSection, string insertId, List<ImageData> images)
        {
            // Append contents
            XDocument partXDoc = part.GetXDocument();
            partXDoc.Declaration.Standalone = Yes;
            partXDoc.Declaration.Encoding = Utf8;

            FixRanges(part.GetXDocument(), newContent);
            AddRelationships(sourceDocument.MainDocumentPart, part, newContent);
            CopyRelatedPartsForContentParts(sourceDocument.MainDocumentPart, part,
                newContent, images);

            // never keep sections for content to be inserted into a header/footer
            List<XElement> adjustedContents = newContent.Where(e => e.Name != W.sectPr).ToList();
            adjustedContents.DescendantsAndSelf(W.sectPr).Remove();
            newContent = adjustedContents;

            CopyNumbering(sourceDocument, newDocument, newContent, images);
            CopyComments(sourceDocument, newDocument, newContent, images);
            AdjustUniqueIds(sourceDocument, newDocument, newContent);
            RemoveGfxdata(newContent);

            if (insertId == null)
                throw new OpenXmlPowerToolsException("Internal error");

            XElement insertElementToReplace = partXDoc
                .Descendants(PtOpenXml.Insert)
                .FirstOrDefault(i => (string)i.Attribute(PtOpenXml.Id) == insertId);
            if (insertElementToReplace != null)
                insertElementToReplace.AddAnnotation(new ReplaceSemaphore());
            partXDoc.Elements().First().ReplaceWith((XElement)InsertTransform(partXDoc.Root, newContent));
        }
        /// ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        public static WmlDocument ExtractGlossaryDocument(WmlDocument wmlGlossaryDocument)
        {
            if (RelationshipMarkup == null)
                InitRelationshipMarkup();

            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(wmlGlossaryDocument.DocumentByteArray, 0, wmlGlossaryDocument.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, false))
                {
                    if (wDoc.MainDocumentPart.GlossaryDocumentPart == null)
                        return null;

                    var fromXd = wDoc.MainDocumentPart.GlossaryDocumentPart.GetXDocument();
                    if (fromXd.Root == null)
                        return null;

                    using (MemoryStream outMs = new MemoryStream())
                    {
                        using (WordprocessingDocument outWDoc = WordprocessingDocument.Create(outMs, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                        {
                            List<ImageData> images = new List<ImageData>();

                            MainDocumentPart mdp = outWDoc.AddMainDocumentPart();
                            var mdpXd = mdp.GetXDocument();
                            XElement root = new XElement(W.document);
                            if (mdpXd.Root == null)
                                mdpXd.Add(root);
                            else
                                mdpXd.Root.ReplaceWith(root);
                            root.Add(new XElement(W.body,
                                fromXd.Root.Elements(W.docParts)));
                            mdp.PutXDocument();

                            var newContent = fromXd.Root.Elements(W.docParts);
                            CopyGlossaryDocumentPartsFromGD(wDoc, outWDoc, newContent, images);
                            CopyRelatedPartsForContentParts(wDoc.MainDocumentPart.GlossaryDocumentPart, mdp, newContent, images);
                        }
                        return new WmlDocument("Glossary.docx", outMs.ToArray());
                    }
                }
            }
        }

        private static void CopyGlossaryDocumentPartsFromGD(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument,
            IEnumerable<XElement> newContent, List<ImageData> images)
        {
            // Copy all styles to the new document
            if (sourceDocument.MainDocumentPart.GlossaryDocumentPart.StyleDefinitionsPart != null)
            {
                XDocument oldStyles = sourceDocument.MainDocumentPart.GlossaryDocumentPart.StyleDefinitionsPart.GetXDocument();
                if (newDocument.MainDocumentPart.StyleDefinitionsPart == null)
                {
                    newDocument.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                    XDocument newStyles = newDocument.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
                    newStyles.Declaration.Standalone = Yes;
                    newStyles.Declaration.Encoding = Utf8;
                    newStyles.Add(oldStyles.Root);
                    newDocument.MainDocumentPart.StyleDefinitionsPart.PutXDocument();
                }
                else
                {
                    XDocument newStyles = newDocument.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
                    MergeStyles(sourceDocument, newDocument, oldStyles, newStyles, newContent);
                    newDocument.MainDocumentPart.StyleDefinitionsPart.PutXDocument();
                }
            }

            // Copy fontTable to the new document
            if (sourceDocument.MainDocumentPart.GlossaryDocumentPart.FontTablePart != null)
            {
                XDocument oldFontTable = sourceDocument.MainDocumentPart.GlossaryDocumentPart.FontTablePart.GetXDocument();
                if (newDocument.MainDocumentPart.FontTablePart == null)
                {
                    newDocument.MainDocumentPart.AddNewPart<FontTablePart>();
                    XDocument newFontTable = newDocument.MainDocumentPart.FontTablePart.GetXDocument();
                    newFontTable.Declaration.Standalone = Yes;
                    newFontTable.Declaration.Encoding = Utf8;
                    newFontTable.Add(oldFontTable.Root);
                    newDocument.MainDocumentPart.FontTablePart.PutXDocument();
                }
                else
                {
                    XDocument newFontTable = newDocument.MainDocumentPart.FontTablePart.GetXDocument();
                    MergeFontTables(oldFontTable, newFontTable);
                    newDocument.MainDocumentPart.FontTablePart.PutXDocument();
                }
            }

            DocumentSettingsPart oldSettingsPart = sourceDocument.MainDocumentPart.GlossaryDocumentPart.DocumentSettingsPart;
            if (oldSettingsPart != null)
            {
                DocumentSettingsPart newSettingsPart = newDocument.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
                XDocument settingsXDoc = oldSettingsPart.GetXDocument();
                AddRelationships(oldSettingsPart, newSettingsPart, new[] { settingsXDoc.Root });
                //CopyFootnotesPart(sourceDocument, newDocument, settingsXDoc, images);
                //CopyEndnotesPart(sourceDocument, newDocument, settingsXDoc, images);
                XDocument newXDoc = newDocument.MainDocumentPart.DocumentSettingsPart.GetXDocument();
                newXDoc.Declaration.Standalone = Yes;
                newXDoc.Declaration.Encoding = Utf8;
                newXDoc.Add(settingsXDoc.Root);
                CopyRelatedPartsForContentParts(oldSettingsPart, newSettingsPart, new[] { newXDoc.Root }, images);
                newSettingsPart.PutXDocument(newXDoc);
            }

            WebSettingsPart oldWebSettingsPart = sourceDocument.MainDocumentPart.GlossaryDocumentPart.WebSettingsPart;
            if (oldWebSettingsPart != null)
            {
                WebSettingsPart newWebSettingsPart = newDocument.MainDocumentPart.AddNewPart<WebSettingsPart>();
                XDocument settingsXDoc = oldWebSettingsPart.GetXDocument();
                AddRelationships(oldWebSettingsPart, newWebSettingsPart, new[] { settingsXDoc.Root });
                XDocument newXDoc = newDocument.MainDocumentPart.WebSettingsPart.GetXDocument();
                newXDoc.Declaration.Standalone = Yes;
                newXDoc.Declaration.Encoding = Utf8;
                newXDoc.Add(settingsXDoc.Root);
                newWebSettingsPart.PutXDocument(newXDoc);
            }

            NumberingDefinitionsPart oldNumberingDefinitionsPart = sourceDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart;
            if (oldNumberingDefinitionsPart != null)
            {
                CopyNumberingForGlossaryDocumentPartFromGD(oldNumberingDefinitionsPart, newDocument, newContent, images);
            }
        }

        private static void CopyGlossaryDocumentPartsToGD(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument,
            IEnumerable<XElement> newContent, List<ImageData> images)
        {
            // Copy all styles to the new document
            if (sourceDocument.MainDocumentPart.StyleDefinitionsPart != null)
            {
                XDocument oldStyles = sourceDocument.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
                newDocument.MainDocumentPart.GlossaryDocumentPart.AddNewPart<StyleDefinitionsPart>();
                XDocument newStyles = newDocument.MainDocumentPart.GlossaryDocumentPart.StyleDefinitionsPart.GetXDocument();
                newStyles.Declaration.Standalone = Yes;
                newStyles.Declaration.Encoding = Utf8;
                newStyles.Add(oldStyles.Root);
                newDocument.MainDocumentPart.GlossaryDocumentPart.StyleDefinitionsPart.PutXDocument();
            }

            // Copy fontTable to the new document
            if (sourceDocument.MainDocumentPart.FontTablePart != null)
            {
                XDocument oldFontTable = sourceDocument.MainDocumentPart.FontTablePart.GetXDocument();
                newDocument.MainDocumentPart.GlossaryDocumentPart.AddNewPart<FontTablePart>();
                XDocument newFontTable = newDocument.MainDocumentPart.GlossaryDocumentPart.FontTablePart.GetXDocument();
                newFontTable.Declaration.Standalone = Yes;
                newFontTable.Declaration.Encoding = Utf8;
                newFontTable.Add(oldFontTable.Root);
                newDocument.MainDocumentPart.FontTablePart.PutXDocument();
            }

            DocumentSettingsPart oldSettingsPart = sourceDocument.MainDocumentPart.DocumentSettingsPart;
            if (oldSettingsPart != null)
            {
                DocumentSettingsPart newSettingsPart = newDocument.MainDocumentPart.GlossaryDocumentPart.AddNewPart<DocumentSettingsPart>();
                XDocument settingsXDoc = oldSettingsPart.GetXDocument();
                AddRelationships(oldSettingsPart, newSettingsPart, new[] { settingsXDoc.Root });
                //CopyFootnotesPart(sourceDocument, newDocument, settingsXDoc, images);
                //CopyEndnotesPart(sourceDocument, newDocument, settingsXDoc, images);
                XDocument newXDoc = newDocument.MainDocumentPart.GlossaryDocumentPart.DocumentSettingsPart.GetXDocument();
                newXDoc.Declaration.Standalone = Yes;
                newXDoc.Declaration.Encoding = Utf8;
                newXDoc.Add(settingsXDoc.Root);
                CopyRelatedPartsForContentParts(oldSettingsPart, newSettingsPart, new[] { newXDoc.Root }, images);
                newSettingsPart.PutXDocument(newXDoc);
            }

            WebSettingsPart oldWebSettingsPart = sourceDocument.MainDocumentPart.WebSettingsPart;
            if (oldWebSettingsPart != null)
            {
                WebSettingsPart newWebSettingsPart = newDocument.MainDocumentPart.GlossaryDocumentPart.AddNewPart<WebSettingsPart>();
                XDocument settingsXDoc = oldWebSettingsPart.GetXDocument();
                AddRelationships(oldWebSettingsPart, newWebSettingsPart, new[] { settingsXDoc.Root });
                XDocument newXDoc = newDocument.MainDocumentPart.GlossaryDocumentPart.WebSettingsPart.GetXDocument();
                newXDoc.Declaration.Standalone = Yes;
                newXDoc.Declaration.Encoding = Utf8;
                newXDoc.Add(settingsXDoc.Root);
                newWebSettingsPart.PutXDocument(newXDoc);
            }

            NumberingDefinitionsPart oldNumberingDefinitionsPart = sourceDocument.MainDocumentPart.NumberingDefinitionsPart;
            if (oldNumberingDefinitionsPart != null)
            {
                CopyNumberingForGlossaryDocumentPartToGD(oldNumberingDefinitionsPart, newDocument, newContent, images);
            }
        }


#if false
        At various locations in Open-Xml-PowerTools, you will find examples of Open XML markup that is associated with code associated with
        querying or generating that markup.  This is an example of the GlossaryDocument part.

<w:glossaryDocument xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid wp14">
  <w:docParts>
    <w:docPart>
      <w:docPartPr>
        <w:name w:val="CDE7B64C7BB446AE905B622B0A882EB6" />
        <w:category>
          <w:name w:val="General" />
          <w:gallery w:val="placeholder" />
        </w:category>
        <w:types>
          <w:type w:val="bbPlcHdr" />
        </w:types>
        <w:behaviors>
          <w:behavior w:val="content" />
        </w:behaviors>
        <w:guid w:val="{13882A71-B5B7-4421-ACBB-9B61C61B3034}" />
      </w:docPartPr>
      <w:docPartBody>
        <w:p w:rsidR="00004EEA" w:rsidRDefault="00AD57F5" w:rsidP="00AD57F5">
          <w:pPr>
            <w:pStyle w:val="CDE7B64C7BB446AE905B622B0A882EB6" />
          </w:pPr>
          <w:r w:rsidRPr="00FB619D">
            <w:rPr>
              <w:rStyle w:val="PlaceholderText" />
              <w:lang w:val="da-DK" />
            </w:rPr>
            <w:t>Produktnavn</w:t>
          </w:r>
          <w:r w:rsidRPr="007379EE">
            <w:rPr>
              <w:rStyle w:val="PlaceholderText" />
            </w:rPr>
            <w:t>.</w:t>
          </w:r>
        </w:p>
      </w:docPartBody>
    </w:docPart>
  </w:docParts>
</w:glossaryDocument>
#endif

        private static void CopyCustomXmlPartsForDataBoundContentControls(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument, IEnumerable<XElement> newContent)
        {
            List<string> itemList = new List<string>();
            foreach (string itemId in newContent
                .Descendants(W.dataBinding)
                .Select(e => (string)e.Attribute(W.storeItemID)))
                if (!itemList.Contains(itemId))
                    itemList.Add(itemId);
            foreach (CustomXmlPart customXmlPart in sourceDocument.MainDocumentPart.CustomXmlParts)
            {
                OpenXmlPart propertyPart = customXmlPart
                    .Parts
                    .Select(p => p.OpenXmlPart)
                    .Where(p => p.ContentType == "application/vnd.openxmlformats-officedocument.customXmlProperties+xml")
                    .FirstOrDefault();
                if (propertyPart != null)
                {
                    XDocument propertyPartDoc = propertyPart.GetXDocument();
                    if (itemList.Contains(propertyPartDoc.Root.Attribute(DS.itemID).Value))
                    {
                        CustomXmlPart newPart = newDocument.MainDocumentPart.AddCustomXmlPart(customXmlPart.ContentType);
                        newPart.GetXDocument().Add(customXmlPart.GetXDocument().Root);
                        foreach (OpenXmlPart propPart in customXmlPart.Parts.Select(p => p.OpenXmlPart))
                        {
                            CustomXmlPropertiesPart newPropPart = newPart.AddNewPart<CustomXmlPropertiesPart>();
                            newPropPart.GetXDocument().Add(propPart.GetXDocument().Root);
                        }
                    }
                }
            }
        }

        private static Dictionary<XName, XName[]> RelationshipMarkup = null;

        private static void UpdateContent(IEnumerable<XElement> newContent, XName elementToModify, string oldRid, string newRid)
        {
            foreach (var attributeName in RelationshipMarkup[elementToModify])
            {
                var elementsToUpdate = newContent
                    .Descendants(elementToModify)
                    .Where(e => (string)e.Attribute(attributeName) == oldRid);
                foreach (var element in elementsToUpdate)
                    element.Attribute(attributeName).Value = newRid;
            }
        }

        private static void AddRelationships(OpenXmlPart oldPart, OpenXmlPart newPart, IEnumerable<XElement> newContent)
        {
            var relevantElements = newContent.DescendantsAndSelf()
                .Where(d => RelationshipMarkup.ContainsKey(d.Name) &&
                    d.Attributes().Any(a => RelationshipMarkup[d.Name].Contains(a.Name)));
            foreach (var e in relevantElements)
            {
                if (e.Name == W.hyperlink)
                {
                    string relId = (string)e.Attribute(R.id);
                    if (string.IsNullOrEmpty(relId))
                        continue;
                    var tempHyperlink = newPart.HyperlinkRelationships.FirstOrDefault(h => h.Id == relId);
                    if (tempHyperlink != null)
                        continue;
                    Guid g = Guid.NewGuid();
                    string newRid = $"R{g:N}";
                    var oldHyperlink = oldPart.HyperlinkRelationships.FirstOrDefault(h => h.Id == relId);
                    if (oldHyperlink == null)
                        continue;
                    //throw new DocumentBuilderInternalException("Internal Error 0002");
                    newPart.AddHyperlinkRelationship(oldHyperlink.Uri, oldHyperlink.IsExternal, newRid);
                    UpdateContent(newContent, e.Name, relId, newRid);
                }
                if (e.Name == W.attachedTemplate || e.Name == W.saveThroughXslt)
                {
                    string relId = (string)e.Attribute(R.id);
                    if (string.IsNullOrEmpty(relId))
                        continue;
                    var tempExternalRelationship = newPart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (tempExternalRelationship != null)
                        continue;
                    Guid g = Guid.NewGuid();
                    string newRid = $"R{g:N}";
                    var oldRel = oldPart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (oldRel == null)
                        throw new DocumentBuilderInternalException("Source {0} is invalid document - hyperlink contains invalid references");
                    newPart.AddExternalRelationship(oldRel.RelationshipType, oldRel.Uri, newRid);
                    UpdateContent(newContent, e.Name, relId, newRid);
                }
                if (e.Name == A.hlinkClick || e.Name == A.hlinkHover || e.Name == A.hlinkMouseOver)
                {
                    string relId = (string)e.Attribute(R.id);
                    if (string.IsNullOrEmpty(relId))
                        continue;
                    var tempHyperlink = newPart.HyperlinkRelationships.FirstOrDefault(h => h.Id == relId);
                    if (tempHyperlink != null)
                        continue;
                    Guid g = Guid.NewGuid();
                    string newRid = $"R{g:N}";
                    var oldHyperlink = oldPart.HyperlinkRelationships.FirstOrDefault(h => h.Id == relId);
                    if (oldHyperlink == null)
                        continue;
                    newPart.AddHyperlinkRelationship(oldHyperlink.Uri, oldHyperlink.IsExternal, newRid);
                    UpdateContent(newContent, e.Name, relId, newRid);
                }
                if (e.Name == VML.imagedata)
                {
                    string relId = (string)e.Attribute(R.href);
                    if (string.IsNullOrEmpty(relId))
                        continue;
                    var tempExternalRelationship = newPart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (tempExternalRelationship != null)
                        continue;
                    Guid g = Guid.NewGuid();
                    string newRid = $"R{g:N}";
                    var oldRel = oldPart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (oldRel == null)
                        throw new DocumentBuilderInternalException("Internal Error 0006");
                    newPart.AddExternalRelationship(oldRel.RelationshipType, oldRel.Uri, newRid);
                    UpdateContent(newContent, e.Name, relId, newRid);
                }
                if (e.Name == A.blip)
                {
                    // <a:blip r:embed="rId6" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" />
                    string relId = (string)e.Attribute(R.link);
                    //if (relId == null)
                    //    relId = (string)e.Attribute(R.embed);
                    if (string.IsNullOrEmpty(relId))
                        continue;
                    var tempExternalRelationship = newPart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (tempExternalRelationship != null)
                        continue;
                    Guid g = Guid.NewGuid();
                    string newRid = $"R{g:N}";
                    var oldRel = oldPart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (oldRel == null)
                        continue;
                    newPart.AddExternalRelationship(oldRel.RelationshipType, oldRel.Uri, newRid);
                    UpdateContent(newContent, e.Name, relId, newRid);
                }
            }
        }

        private class FromPreviousSourceSemaphore { };

        private static void CopyNumbering(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument,
            IEnumerable<XElement> newContent, List<ImageData> images)
        {
            Dictionary<int, int> numIdMap = new Dictionary<int, int>();
            int number = 1;
            int abstractNumber = 0;
            XDocument oldNumbering = null;
            XDocument newNumbering = null;

            foreach (XElement numReference in newContent.DescendantsAndSelf(W.numPr))
            {
                XElement idElement = numReference.Descendants(W.numId).FirstOrDefault();
                if (idElement != null)
                {
                    if (oldNumbering == null)
                        oldNumbering = sourceDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                    if (newNumbering == null)
                    {
                        if (newDocument.MainDocumentPart.NumberingDefinitionsPart != null)
                        {
                            newNumbering = newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                            var numIds = newNumbering
                                .Root
                                .Elements(W.num)
                                .Select(f => (int)f.Attribute(W.numId));
                            if (numIds.Any())
                                number = numIds.Max() + 1;
                            numIds = newNumbering
                                .Root
                                .Elements(W.abstractNum)
                                .Select(f => (int)f.Attribute(W.abstractNumId));
                            if (numIds.Any())
                                abstractNumber = numIds.Max() + 1;
                        }
                        else
                        {
                            newDocument.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                            newNumbering = newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                            newNumbering.Declaration.Standalone = Yes;
                            newNumbering.Declaration.Encoding = Utf8;
                            newNumbering.Add(new XElement(W.numbering, NamespaceAttributes));
                        }
                    }
                    int numId = (int)idElement.Attribute(W.val);
                    if (numId != 0)
                    {
                        XElement element = oldNumbering
                            .Descendants(W.num)
                            .Where(p => ((int)p.Attribute(W.numId)) == numId)
                            .FirstOrDefault();
                        if (element == null)
                            continue;

                        // Copy abstract numbering element, if necessary (use matching NSID)
                        string abstractNumIdStr = (string)element
                            .Elements(W.abstractNumId)
                            .First()
                            .Attribute(W.val);
                        int abstractNumId;
                        if (!int.TryParse(abstractNumIdStr, out abstractNumId))
                            throw new DocumentBuilderException("Invalid document - invalid value for abstractNumId");

                        XElement abstractElement = oldNumbering
                            .Descendants()
                            .Elements(W.abstractNum)
                            .Where(p => ((int)p.Attribute(W.abstractNumId)) == abstractNumId)
                            .First();
                        XElement nsidElement = abstractElement
                            .Element(W.nsid);
                        string abstractNSID = null;
                        if (nsidElement != null)
                            abstractNSID = (string)nsidElement
                                .Attribute(W.val);
                        XElement newAbstractElement = newNumbering
                            .Descendants()
                            .Elements(W.abstractNum)
                            .Where(e => e.Annotation<FromPreviousSourceSemaphore>() == null)
                            .Where(p =>
                            {
                                var thisNsidElement = p.Element(W.nsid);
                                if (thisNsidElement == null)
                                    return false;
                                return (string)thisNsidElement.Attribute(W.val) == abstractNSID;
                            })
                            .FirstOrDefault();
                        if (newAbstractElement == null)
                        {
                            newAbstractElement = new XElement(abstractElement);
                            newAbstractElement.Attribute(W.abstractNumId).Value = abstractNumber.ToString();
                            abstractNumber++;
                            if (newNumbering.Root.Elements(W.abstractNum).Any())
                                newNumbering.Root.Elements(W.abstractNum).Last().AddAfterSelf(newAbstractElement);
                            else
                                newNumbering.Root.Add(newAbstractElement);

                            foreach (XElement pictId in newAbstractElement.Descendants(W.lvlPicBulletId))
                            {
                                string bulletId = (string)pictId.Attribute(W.val);
                                XElement numPicBullet = oldNumbering
                                    .Descendants(W.numPicBullet)
                                    .FirstOrDefault(d => (string)d.Attribute(W.numPicBulletId) == bulletId);
                                int maxNumPicBulletId = new int[] { -1 }.Concat(
                                    newNumbering.Descendants(W.numPicBullet)
                                    .Attributes(W.numPicBulletId)
                                    .Select(a => (int)a))
                                    .Max() + 1;
                                XElement newNumPicBullet = new XElement(numPicBullet);
                                newNumPicBullet.Attribute(W.numPicBulletId).Value = maxNumPicBulletId.ToString();
                                pictId.Attribute(W.val).Value = maxNumPicBulletId.ToString();
                                newNumbering.Root.AddFirst(newNumPicBullet);
                            }
                        }
                        string newAbstractId = newAbstractElement.Attribute(W.abstractNumId).Value;

                        // Copy numbering element, if necessary (use matching element with no overrides)
                        XElement newElement;
                        if (numIdMap.ContainsKey(numId))
                        {
                            newElement = newNumbering
                                .Descendants()
                                .Elements(W.num)
                                .Where(e => e.Annotation<FromPreviousSourceSemaphore>() == null)
                                .Where(p => ((int)p.Attribute(W.numId)) == numIdMap[numId])
                                .First();
                        }
                        else
                        {
                            newElement = new XElement(element);
                            newElement
                                .Elements(W.abstractNumId)
                                .First()
                                .Attribute(W.val).Value = newAbstractId;
                            newElement.Attribute(W.numId).Value = number.ToString();
                            numIdMap.Add(numId, number);
                            number++;
                            newNumbering.Root.Add(newElement);
                        }
                        idElement.Attribute(W.val).Value = newElement.Attribute(W.numId).Value;
                    }
                }
            }
            if (newNumbering != null)
            {
                foreach (var abstractNum in newNumbering.Descendants(W.abstractNum))
                    abstractNum.AddAnnotation(new FromPreviousSourceSemaphore());
                foreach (var num in newNumbering.Descendants(W.num))
                    num.AddAnnotation(new FromPreviousSourceSemaphore());
            }

            if (newDocument.MainDocumentPart.NumberingDefinitionsPart != null &&
                sourceDocument.MainDocumentPart.NumberingDefinitionsPart != null)
            {
                AddRelationships(sourceDocument.MainDocumentPart.NumberingDefinitionsPart,
                    newDocument.MainDocumentPart.NumberingDefinitionsPart,
                    new[] { newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(sourceDocument.MainDocumentPart.NumberingDefinitionsPart,
                    newDocument.MainDocumentPart.NumberingDefinitionsPart,
                    new[] { newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument().Root }, images);
            }
        }

        // Note: the following two methods were added with almost exact duplicate code to the method above, because I do not want to touch that code.
        private static void CopyNumberingForGlossaryDocumentPartFromGD(NumberingDefinitionsPart sourceNumberingPart, WordprocessingDocument newDocument,
            IEnumerable<XElement> newContent, List<ImageData> images)
        {
            Dictionary<int, int> numIdMap = new Dictionary<int, int>();
            int number = 1;
            int abstractNumber = 0;
            XDocument oldNumbering = null;
            XDocument newNumbering = null;

            foreach (XElement numReference in newContent.DescendantsAndSelf(W.numPr))
            {
                XElement idElement = numReference.Descendants(W.numId).FirstOrDefault();
                if (idElement != null)
                {
                    if (oldNumbering == null)
                        oldNumbering = sourceNumberingPart.GetXDocument();
                    if (newNumbering == null)
                    {
                        if (newDocument.MainDocumentPart.NumberingDefinitionsPart != null)
                        {
                            newNumbering = newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                            var numIds = newNumbering
                                .Root
                                .Elements(W.num)
                                .Select(f => (int)f.Attribute(W.numId));
                            if (numIds.Any())
                                number = numIds.Max() + 1;
                            numIds = newNumbering
                                .Root
                                .Elements(W.abstractNum)
                                .Select(f => (int)f.Attribute(W.abstractNumId));
                            if (numIds.Any())
                                abstractNumber = numIds.Max() + 1;
                        }
                        else
                        {
                            newDocument.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                            newNumbering = newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                            newNumbering.Declaration.Standalone = Yes;
                            newNumbering.Declaration.Encoding = Utf8;
                            newNumbering.Add(new XElement(W.numbering, NamespaceAttributes));
                        }
                    }
                    int numId = (int)idElement.Attribute(W.val);
                    if (numId != 0)
                    {
                        XElement element = oldNumbering
                            .Descendants(W.num)
                            .Where(p => ((int)p.Attribute(W.numId)) == numId)
                            .FirstOrDefault();
                        if (element == null)
                            continue;

                        // Copy abstract numbering element, if necessary (use matching NSID)
                        string abstractNumIdStr = (string)element
                            .Elements(W.abstractNumId)
                            .First()
                            .Attribute(W.val);
                        int abstractNumId;
                        if (!int.TryParse(abstractNumIdStr, out abstractNumId))
                            throw new DocumentBuilderException("Invalid document - invalid value for abstractNumId");
                        XElement abstractElement = oldNumbering
                            .Descendants()
                            .Elements(W.abstractNum)
                            .Where(p => ((int)p.Attribute(W.abstractNumId)) == abstractNumId)
                            .First();
                        XElement nsidElement = abstractElement
                            .Element(W.nsid);
                        string abstractNSID = null;
                        if (nsidElement != null)
                            abstractNSID = (string)nsidElement
                                .Attribute(W.val);
                        XElement newAbstractElement = newNumbering
                            .Descendants()
                            .Elements(W.abstractNum)
                            .Where(e => e.Annotation<FromPreviousSourceSemaphore>() == null)
                            .Where(p =>
                            {
                                var thisNsidElement = p.Element(W.nsid);
                                if (thisNsidElement == null)
                                    return false;
                                return (string)thisNsidElement.Attribute(W.val) == abstractNSID;
                            })
                            .FirstOrDefault();
                        if (newAbstractElement == null)
                        {
                            newAbstractElement = new XElement(abstractElement);
                            newAbstractElement.Attribute(W.abstractNumId).Value = abstractNumber.ToString();
                            abstractNumber++;
                            if (newNumbering.Root.Elements(W.abstractNum).Any())
                                newNumbering.Root.Elements(W.abstractNum).Last().AddAfterSelf(newAbstractElement);
                            else
                                newNumbering.Root.Add(newAbstractElement);

                            foreach (XElement pictId in newAbstractElement.Descendants(W.lvlPicBulletId))
                            {
                                string bulletId = (string)pictId.Attribute(W.val);
                                XElement numPicBullet = oldNumbering
                                    .Descendants(W.numPicBullet)
                                    .FirstOrDefault(d => (string)d.Attribute(W.numPicBulletId) == bulletId);
                                int maxNumPicBulletId = new int[] { -1 }.Concat(
                                    newNumbering.Descendants(W.numPicBullet)
                                    .Attributes(W.numPicBulletId)
                                    .Select(a => (int)a))
                                    .Max() + 1;
                                XElement newNumPicBullet = new XElement(numPicBullet);
                                newNumPicBullet.Attribute(W.numPicBulletId).Value = maxNumPicBulletId.ToString();
                                pictId.Attribute(W.val).Value = maxNumPicBulletId.ToString();
                                newNumbering.Root.AddFirst(newNumPicBullet);
                            }
                        }
                        string newAbstractId = newAbstractElement.Attribute(W.abstractNumId).Value;

                        // Copy numbering element, if necessary (use matching element with no overrides)
                        XElement newElement;
                        if (numIdMap.ContainsKey(numId))
                        {
                            newElement = newNumbering
                                .Descendants()
                                .Elements(W.num)
                                .Where(e => e.Annotation<FromPreviousSourceSemaphore>() == null)
                                .Where(p => ((int)p.Attribute(W.numId)) == numIdMap[numId])
                                .First();
                        }
                        else
                        {
                            newElement = new XElement(element);
                            newElement
                                .Elements(W.abstractNumId)
                                .First()
                                .Attribute(W.val).Value = newAbstractId;
                            newElement.Attribute(W.numId).Value = number.ToString();
                            numIdMap.Add(numId, number);
                            number++;
                            newNumbering.Root.Add(newElement);
                        }
                        idElement.Attribute(W.val).Value = newElement.Attribute(W.numId).Value;
                    }
                }
            }
            if (newNumbering != null)
            {
                foreach (var abstractNum in newNumbering.Descendants(W.abstractNum))
                    abstractNum.AddAnnotation(new FromPreviousSourceSemaphore());
                foreach (var num in newNumbering.Descendants(W.num))
                    num.AddAnnotation(new FromPreviousSourceSemaphore());
            }

            if (newDocument.MainDocumentPart.NumberingDefinitionsPart != null &&
                sourceNumberingPart != null)
            {
                AddRelationships(sourceNumberingPart,
                    newDocument.MainDocumentPart.NumberingDefinitionsPart,
                    new[] { newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(sourceNumberingPart,
                    newDocument.MainDocumentPart.NumberingDefinitionsPart,
                    new[] { newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument().Root }, images);
            }
            if (newDocument.MainDocumentPart.NumberingDefinitionsPart != null)
                newDocument.MainDocumentPart.NumberingDefinitionsPart.PutXDocument();
        }

        private static void CopyNumberingForGlossaryDocumentPartToGD(NumberingDefinitionsPart sourceNumberingPart, WordprocessingDocument newDocument,
            IEnumerable<XElement> newContent, List<ImageData> images)
        {
            Dictionary<int, int> numIdMap = new Dictionary<int, int>();
            int number = 1;
            int abstractNumber = 0;
            XDocument oldNumbering = null;
            XDocument newNumbering = null;

            foreach (XElement numReference in newContent.DescendantsAndSelf(W.numPr))
            {
                XElement idElement = numReference.Descendants(W.numId).FirstOrDefault();
                if (idElement != null)
                {
                    if (oldNumbering == null)
                        oldNumbering = sourceNumberingPart.GetXDocument();
                    if (newNumbering == null)
                    {
                        if (newDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart != null)
                        {
                            newNumbering = newDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart.GetXDocument();
                            var numIds = newNumbering
                                .Root
                                .Elements(W.num)
                                .Select(f => (int)f.Attribute(W.numId));
                            if (numIds.Any())
                                number = numIds.Max() + 1;
                            numIds = newNumbering
                                .Root
                                .Elements(W.abstractNum)
                                .Select(f => (int)f.Attribute(W.abstractNumId));
                            if (numIds.Any())
                                abstractNumber = numIds.Max() + 1;
                        }
                        else
                        {
                            newDocument.MainDocumentPart.GlossaryDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                            newNumbering = newDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart.GetXDocument();
                            newNumbering.Declaration.Standalone = Yes;
                            newNumbering.Declaration.Encoding = Utf8;
                            newNumbering.Add(new XElement(W.numbering, NamespaceAttributes));
                        }
                    }
                    int numId = (int)idElement.Attribute(W.val);
                    if (numId != 0)
                    {
                        XElement element = oldNumbering
                            .Descendants(W.num)
                            .Where(p => ((int)p.Attribute(W.numId)) == numId)
                            .FirstOrDefault();
                        if (element == null)
                            continue;

                        // Copy abstract numbering element, if necessary (use matching NSID)
                        string abstractNumIdStr = (string)element
                            .Elements(W.abstractNumId)
                            .First()
                            .Attribute(W.val);
                        int abstractNumId;
                        if (!int.TryParse(abstractNumIdStr, out abstractNumId))
                            throw new DocumentBuilderException("Invalid document - invalid value for abstractNumId");
                        XElement abstractElement = oldNumbering
                            .Descendants()
                            .Elements(W.abstractNum)
                            .Where(p => ((int)p.Attribute(W.abstractNumId)) == abstractNumId)
                            .First();
                        XElement nsidElement = abstractElement
                            .Element(W.nsid);
                        string abstractNSID = null;
                        if (nsidElement != null)
                            abstractNSID = (string)nsidElement
                                .Attribute(W.val);
                        XElement newAbstractElement = newNumbering
                            .Descendants()
                            .Elements(W.abstractNum)
                            .Where(e => e.Annotation<FromPreviousSourceSemaphore>() == null)
                            .Where(p =>
                            {
                                var thisNsidElement = p.Element(W.nsid);
                                if (thisNsidElement == null)
                                    return false;
                                return (string)thisNsidElement.Attribute(W.val) == abstractNSID;
                            })
                            .FirstOrDefault();
                        if (newAbstractElement == null)
                        {
                            newAbstractElement = new XElement(abstractElement);
                            newAbstractElement.Attribute(W.abstractNumId).Value = abstractNumber.ToString();
                            abstractNumber++;
                            if (newNumbering.Root.Elements(W.abstractNum).Any())
                                newNumbering.Root.Elements(W.abstractNum).Last().AddAfterSelf(newAbstractElement);
                            else
                                newNumbering.Root.Add(newAbstractElement);

                            foreach (XElement pictId in newAbstractElement.Descendants(W.lvlPicBulletId))
                            {
                                string bulletId = (string)pictId.Attribute(W.val);
                                XElement numPicBullet = oldNumbering
                                    .Descendants(W.numPicBullet)
                                    .FirstOrDefault(d => (string)d.Attribute(W.numPicBulletId) == bulletId);
                                int maxNumPicBulletId = new int[] { -1 }.Concat(
                                    newNumbering.Descendants(W.numPicBullet)
                                    .Attributes(W.numPicBulletId)
                                    .Select(a => (int)a))
                                    .Max() + 1;
                                XElement newNumPicBullet = new XElement(numPicBullet);
                                newNumPicBullet.Attribute(W.numPicBulletId).Value = maxNumPicBulletId.ToString();
                                pictId.Attribute(W.val).Value = maxNumPicBulletId.ToString();
                                newNumbering.Root.AddFirst(newNumPicBullet);
                            }
                        }
                        string newAbstractId = newAbstractElement.Attribute(W.abstractNumId).Value;

                        // Copy numbering element, if necessary (use matching element with no overrides)
                        XElement newElement;
                        if (numIdMap.ContainsKey(numId))
                        {
                            newElement = newNumbering
                                .Descendants()
                                .Elements(W.num)
                                .Where(e => e.Annotation<FromPreviousSourceSemaphore>() == null)
                                .Where(p => ((int)p.Attribute(W.numId)) == numIdMap[numId])
                                .First();
                        }
                        else
                        {
                            newElement = new XElement(element);
                            newElement
                                .Elements(W.abstractNumId)
                                .First()
                                .Attribute(W.val).Value = newAbstractId;
                            newElement.Attribute(W.numId).Value = number.ToString();
                            numIdMap.Add(numId, number);
                            number++;
                            newNumbering.Root.Add(newElement);
                        }
                        idElement.Attribute(W.val).Value = newElement.Attribute(W.numId).Value;
                    }
                }
            }
            if (newNumbering != null)
            {
                foreach (var abstractNum in newNumbering.Descendants(W.abstractNum))
                    abstractNum.AddAnnotation(new FromPreviousSourceSemaphore());
                foreach (var num in newNumbering.Descendants(W.num))
                    num.AddAnnotation(new FromPreviousSourceSemaphore());
            }

            if (newDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart != null &&
                sourceNumberingPart != null)
            {
                AddRelationships(sourceNumberingPart,
                    newDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart,
                    new[] { newDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(sourceNumberingPart,
                    newDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart,
                    new[] { newDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart.GetXDocument().Root }, images);
            }
            if (newDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart != null)
                newDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart.PutXDocument();
        }

        private static void CopyRelatedImage(OpenXmlPart oldContentPart, OpenXmlPart newContentPart, XElement imageReference, XName attributeName,
            List<ImageData> images)
        {
            string relId = (string)imageReference.Attribute(attributeName);
            if (string.IsNullOrEmpty(relId))
                return;

            // First look to see if this relId has already been added to the new document.
            // This is necessary for those parts that get processed with both old and new ids, such as the comments
            // part.  This is not necessary for parts such as the main document part, but this code won't malfunction
            // in that case.
            var tempPartIdPair5 = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
            if (tempPartIdPair5 != null)
                return;

            ExternalRelationship tempEr5 = newContentPart.ExternalRelationships.FirstOrDefault(er => er.Id == relId);
            if (tempEr5 != null)
                return;

            var ipp2 = oldContentPart.Parts.FirstOrDefault(ipp => ipp.RelationshipId == relId);
            if (ipp2 != null)
            {
                var oldPart2 = ipp2.OpenXmlPart;
                if (!(oldPart2 is ImagePart))
                    throw new DocumentBuilderException("Invalid document - target part is not ImagePart");

                ImagePart oldPart = (ImagePart)ipp2.OpenXmlPart;
                ImageData temp = ManageImageCopy(oldPart, newContentPart, images);
                if (temp.ImagePart == null)
                {
                    ImagePart newPart = null;
                    if (newContentPart is MainDocumentPart)
                        newPart = ((MainDocumentPart)newContentPart).AddImagePart(oldPart.ContentType);
                    if (newContentPart is HeaderPart)
                        newPart = ((HeaderPart)newContentPart).AddImagePart(oldPart.ContentType);
                    if (newContentPart is FooterPart)
                        newPart = ((FooterPart)newContentPart).AddImagePart(oldPart.ContentType);
                    if (newContentPart is EndnotesPart)
                        newPart = ((EndnotesPart)newContentPart).AddImagePart(oldPart.ContentType);
                    if (newContentPart is FootnotesPart)
                        newPart = ((FootnotesPart)newContentPart).AddImagePart(oldPart.ContentType);
                    if (newContentPart is ThemePart)
                        newPart = ((ThemePart)newContentPart).AddImagePart(oldPart.ContentType);
                    if (newContentPart is WordprocessingCommentsPart)
                        newPart = ((WordprocessingCommentsPart)newContentPart).AddImagePart(oldPart.ContentType);
                    if (newContentPart is DocumentSettingsPart)
                        newPart = ((DocumentSettingsPart)newContentPart).AddImagePart(oldPart.ContentType);
                    if (newContentPart is ChartPart)
                        newPart = ((ChartPart)newContentPart).AddImagePart(oldPart.ContentType);
                    if (newContentPart is NumberingDefinitionsPart)
                        newPart = ((NumberingDefinitionsPart)newContentPart).AddImagePart(oldPart.ContentType);
                    if (newContentPart is DiagramDataPart)
                        newPart = ((DiagramDataPart)newContentPart).AddImagePart(oldPart.ContentType);
                    if (newContentPart is ChartDrawingPart)
                        newPart = ((ChartDrawingPart)newContentPart).AddImagePart(oldPart.ContentType);
                    temp.ImagePart = newPart;
                    var id = newContentPart.GetIdOfPart(newPart);
                    temp.AddContentPartRelTypeResourceIdTupple(newContentPart, newPart.RelationshipType, id);
                    imageReference.Attribute(attributeName).Value = id;
                    temp.WriteImage(newPart);
                }
                else
                {
                    var refRel = newContentPart.Parts.FirstOrDefault(pip =>
                    {
                        var rel = temp.ContentPartRelTypeIdList.FirstOrDefault(cpr =>
                        {
                            var found = cpr.ContentPart == newContentPart;
                            return found;
                        });
                        return rel != null;
                    });
                    if (refRel != null)
                    {
                        imageReference.Attribute(attributeName).Value = temp.ContentPartRelTypeIdList.First(cpr =>
                        {
                            var found = cpr.ContentPart == newContentPart;
                            return found;
                        }).RelationshipId;
                        return;
                    }
                    var g = new Guid();
                    var newId = $"R{g:N}".Substring(0, 16);
                    newContentPart.CreateRelationshipToPart(temp.ImagePart, newId);
                    imageReference.Attribute(R.id).Value = newId;
                }
            }
            else
            {
                ExternalRelationship er = oldContentPart.ExternalRelationships.FirstOrDefault(er1 => er1.Id == relId);
                if (er != null)
                {
                    ExternalRelationship newEr = newContentPart.AddExternalRelationship(er.RelationshipType, er.Uri);
                    imageReference.Attribute(R.id).Value = newEr.Id;
                }
                throw new DocumentBuilderInternalException("Source {0} is unsupported document - contains reference to NULL image");
            }
        }

        private static void CopyRelatedPartsForContentParts(OpenXmlPart oldContentPart, OpenXmlPart newContentPart,
            IEnumerable<XElement> newContent, List<ImageData> images)
        {
            var relevantElements = newContent.DescendantsAndSelf()
                .Where(d => d.Name == VML.imagedata || d.Name == VML.fill || d.Name == VML.stroke || d.Name == A.blip)
                .ToList();
            foreach (XElement imageReference in relevantElements)
            {
                CopyRelatedImage(oldContentPart, newContentPart, imageReference, R.embed, images);
                CopyRelatedImage(oldContentPart, newContentPart, imageReference, R.pict, images);
                CopyRelatedImage(oldContentPart, newContentPart, imageReference, R.id, images);
            }

            foreach (XElement diagramReference in newContent.DescendantsAndSelf().Where(d => d.Name == DGM.relIds || d.Name == A.relIds))
            {
                // dm attribute
                string relId = diagramReference.Attribute(R.dm).Value;
                var ipp = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (ipp != null)
                {
                    OpenXmlPart tempPart = ipp.OpenXmlPart;
                    continue;
                }

                ExternalRelationship tempEr = newContentPart.ExternalRelationships.FirstOrDefault(er2 => er2.Id == relId);
                if (tempEr != null)
                    continue;

                OpenXmlPart oldPart = oldContentPart.GetPartById(relId);
                OpenXmlPart newPart = newContentPart.AddNewPart<DiagramDataPart>();
                newPart.GetXDocument().Add(oldPart.GetXDocument().Root);
                diagramReference.Attribute(R.dm).Value = newContentPart.GetIdOfPart(newPart);
                AddRelationships(oldPart, newPart, new[] { newPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(oldPart, newPart, new[] { newPart.GetXDocument().Root }, images);

                // lo attribute
                relId = diagramReference.Attribute(R.lo).Value;
                var ipp2 = newContentPart.Parts.FirstOrDefault(z => z.RelationshipId == relId);
                if (ipp2 != null)
                {
                    OpenXmlPart tempPart = ipp2.OpenXmlPart;
                    continue;
                }


                ExternalRelationship tempEr4 = newContentPart.ExternalRelationships.FirstOrDefault(er3 => er3.Id == relId);
                if (tempEr4 != null)
                    continue;

                oldPart = oldContentPart.GetPartById(relId);
                newPart = newContentPart.AddNewPart<DiagramLayoutDefinitionPart>();
                newPart.GetXDocument().Add(oldPart.GetXDocument().Root);
                diagramReference.Attribute(R.lo).Value = newContentPart.GetIdOfPart(newPart);
                AddRelationships(oldPart, newPart, new[] { newPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(oldPart, newPart, new[] { newPart.GetXDocument().Root }, images);

                // qs attribute
                relId = diagramReference.Attribute(R.qs).Value;
                var ipp5 = newContentPart.Parts.FirstOrDefault(z => z.RelationshipId == relId);
                if (ipp5 != null)
                {
                    OpenXmlPart tempPart = ipp5.OpenXmlPart;
                    continue;
                }

                ExternalRelationship tempEr5 = newContentPart.ExternalRelationships.FirstOrDefault(z => z.Id == relId);
                if (tempEr5 != null)
                    continue;

                oldPart = oldContentPart.GetPartById(relId);
                newPart = newContentPart.AddNewPart<DiagramStylePart>();
                newPart.GetXDocument().Add(oldPart.GetXDocument().Root);
                diagramReference.Attribute(R.qs).Value = newContentPart.GetIdOfPart(newPart);
                AddRelationships(oldPart, newPart, new[] { newPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(oldPart, newPart, new[] { newPart.GetXDocument().Root }, images);

                // cs attribute
                relId = diagramReference.Attribute(R.cs).Value;
                var ipp6 = newContentPart.Parts.FirstOrDefault(z => z.RelationshipId == relId);
                if (ipp6 != null)
                {
                    OpenXmlPart tempPart = ipp6.OpenXmlPart;
                    continue;
                }

                ExternalRelationship tempEr6 = newContentPart.ExternalRelationships.FirstOrDefault(z => z.Id == relId);
                if (tempEr6 != null)
                    continue;

                oldPart = oldContentPart.GetPartById(relId);
                newPart = newContentPart.AddNewPart<DiagramColorsPart>();
                newPart.GetXDocument().Add(oldPart.GetXDocument().Root);
                diagramReference.Attribute(R.cs).Value = newContentPart.GetIdOfPart(newPart);
                AddRelationships(oldPart, newPart, new[] { newPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(oldPart, newPart, new[] { newPart.GetXDocument().Root }, images);
            }

            foreach (XElement oleReference in newContent.DescendantsAndSelf(O.OLEObject))
            {
                string relId = (string)oleReference.Attribute(R.id);

                // First look to see if this relId has already been added to the new document.
                // This is necessary for those parts that get processed with both old and new ids, such as the comments
                // part.  This is not necessary for parts such as the main document part, but this code won't malfunction
                // in that case.
                var ipp1 = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (ipp1 != null)
                {
                    OpenXmlPart tempPart = ipp1.OpenXmlPart;
                    continue;
                }

                ExternalRelationship tempEr1 = newContentPart.ExternalRelationships.FirstOrDefault(z => z.Id == relId);
                if (tempEr1 != null)
                    continue;

                var ipp4 = oldContentPart.Parts.FirstOrDefault(z => z.RelationshipId == relId);
                if (ipp4 != null)
                {
                    OpenXmlPart oldPart = oldContentPart.GetPartById(relId);
                    OpenXmlPart newPart = null;
                    if (oldPart is EmbeddedObjectPart)
                    {
                        if (newContentPart is HeaderPart)
                            newPart = ((HeaderPart)newContentPart).AddEmbeddedObjectPart(oldPart.ContentType);
                        if (newContentPart is FooterPart)
                            newPart = ((FooterPart)newContentPart).AddEmbeddedObjectPart(oldPart.ContentType);
                        if (newContentPart is MainDocumentPart)
                            newPart = ((MainDocumentPart)newContentPart).AddEmbeddedObjectPart(oldPart.ContentType);
                        if (newContentPart is FootnotesPart)
                            newPart = ((FootnotesPart)newContentPart).AddEmbeddedObjectPart(oldPart.ContentType);
                        if (newContentPart is EndnotesPart)
                            newPart = ((EndnotesPart)newContentPart).AddEmbeddedObjectPart(oldPart.ContentType);
                        if (newContentPart is WordprocessingCommentsPart)
                            newPart = ((WordprocessingCommentsPart)newContentPart).AddEmbeddedObjectPart(oldPart.ContentType);
                    }
                    else if (oldPart is EmbeddedPackagePart)
                    {
                        if (newContentPart is HeaderPart)
                            newPart = ((HeaderPart)newContentPart).AddEmbeddedPackagePart(oldPart.ContentType);
                        if (newContentPart is FooterPart)
                            newPart = ((FooterPart)newContentPart).AddEmbeddedPackagePart(oldPart.ContentType);
                        if (newContentPart is MainDocumentPart)
                            newPart = ((MainDocumentPart)newContentPart).AddEmbeddedPackagePart(oldPart.ContentType);
                        if (newContentPart is FootnotesPart)
                            newPart = ((FootnotesPart)newContentPart).AddEmbeddedPackagePart(oldPart.ContentType);
                        if (newContentPart is EndnotesPart)
                            newPart = ((EndnotesPart)newContentPart).AddEmbeddedPackagePart(oldPart.ContentType);
                        if (newContentPart is WordprocessingCommentsPart)
                            newPart = ((WordprocessingCommentsPart)newContentPart).AddEmbeddedPackagePart(oldPart.ContentType);
                        if (newContentPart is ChartPart)
                            newPart = ((ChartPart)newContentPart).AddEmbeddedPackagePart(oldPart.ContentType);
                    }
                    using (Stream oldObject = oldPart.GetStream(FileMode.Open, FileAccess.Read))
                    using (Stream newObject = newPart.GetStream(FileMode.Create, FileAccess.ReadWrite))
                    {
                        int byteCount;
                        byte[] buffer = new byte[65536];
                        while ((byteCount = oldObject.Read(buffer, 0, 65536)) != 0)
                            newObject.Write(buffer, 0, byteCount);
                    }
                    oleReference.Attribute(R.id).Value = newContentPart.GetIdOfPart(newPart);
                }
                else
                {
                    if (relId != null)
                    {
                        ExternalRelationship er = oldContentPart.GetExternalRelationship(relId);
                        ExternalRelationship newEr = newContentPart.AddExternalRelationship(er.RelationshipType, er.Uri);
                        oleReference.Attribute(R.id).Value = newEr.Id;
                    }
                }
            }

            foreach (XElement chartReference in newContent.DescendantsAndSelf(C.chart))
            {
                string relId = (string)chartReference.Attribute(R.id);
                if (string.IsNullOrEmpty(relId))
                    continue;
                var ipp2 = newContentPart.Parts.FirstOrDefault(z => z.RelationshipId == relId);
                if (ipp2 != null)
                {
                    OpenXmlPart tempPart = ipp2.OpenXmlPart;
                    continue;
                }

                ExternalRelationship tempEr2 = newContentPart.ExternalRelationships.FirstOrDefault(z => z.Id == relId);
                if (tempEr2 != null)
                    continue;

                var ipp3 = oldContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (ipp3 == null)
                    continue;
                ChartPart oldPart = (ChartPart)ipp3.OpenXmlPart;
                XDocument oldChart = oldPart.GetXDocument();
                ChartPart newPart = newContentPart.AddNewPart<ChartPart>();
                XDocument newChart = newPart.GetXDocument();
                newChart.Add(oldChart.Root);
                chartReference.Attribute(R.id).Value = newContentPart.GetIdOfPart(newPart);
                CopyChartObjects(oldPart, newPart);
                CopyRelatedPartsForContentParts(oldPart, newPart, new[] { newChart.Root }, images);
            }

            foreach (XElement userShape in newContent.DescendantsAndSelf(C.userShapes))
            {
                string relId = (string)userShape.Attribute(R.id);
                if (string.IsNullOrEmpty(relId))
                    continue;

                var ipp4 = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (ipp4 != null)
                {
                    OpenXmlPart tempPart = ipp4.OpenXmlPart;
                    continue;
                }

                ExternalRelationship tempEr4 = newContentPart.ExternalRelationships.FirstOrDefault(z => z.Id == relId);
                if (tempEr4 != null)
                    continue;

                var ipp5 = oldContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (ipp5 != null)
                {
                    ChartDrawingPart oldPart = (ChartDrawingPart)ipp5.OpenXmlPart;
                    XDocument oldXDoc = oldPart.GetXDocument();
                    ChartDrawingPart newPart = newContentPart.AddNewPart<ChartDrawingPart>();
                    XDocument newXDoc = newPart.GetXDocument();
                    newXDoc.Add(oldXDoc.Root);
                    userShape.Attribute(R.id).Value = newContentPart.GetIdOfPart(newPart);
                    AddRelationships(oldPart, newPart, newContent);
                    CopyRelatedPartsForContentParts(oldPart, newPart, new[] { newXDoc.Root }, images);
                }
            }
        }

        private static void CopyFontTable(FontTablePart oldFontTablePart, FontTablePart newFontTablePart)
        {
            var relevantElements = oldFontTablePart.GetXDocument().Descendants().Where(d => d.Name == W.embedRegular ||
                d.Name == W.embedBold || d.Name == W.embedItalic || d.Name == W.embedBoldItalic).ToList();
            foreach (XElement fontReference in relevantElements)
            {
                string relId = (string)fontReference.Attribute(R.id);
                if (string.IsNullOrEmpty(relId))
                    continue;

                var ipp1 = newFontTablePart.Parts.FirstOrDefault(z => z.RelationshipId == relId);
                if (ipp1 != null)
                {
                    OpenXmlPart tempPart = ipp1.OpenXmlPart;
                    continue;
                }

                ExternalRelationship tempEr1 = newFontTablePart.ExternalRelationships.FirstOrDefault(z => z.Id == relId);
                if (tempEr1 != null)
                    continue;

                var oldPart2 = oldFontTablePart.GetPartById(relId);
                if (oldPart2 == null || (!(oldPart2 is FontPart)))
                    throw new DocumentBuilderException("Invalid document - FontTablePart contains invalid relationship");

                FontPart oldPart = (FontPart)oldPart2;
                FontPart newPart = newFontTablePart.AddFontPart(oldPart.ContentType);
                var ResourceID = newFontTablePart.GetIdOfPart(newPart);
                using (Stream oldFont = oldPart.GetStream(FileMode.Open, FileAccess.Read))
                using (Stream newFont = newPart.GetStream(FileMode.Create, FileAccess.ReadWrite))
                {
                    int byteCount;
                    byte[] buffer = new byte[65536];
                    while ((byteCount = oldFont.Read(buffer, 0, 65536)) != 0)
                        newFont.Write(buffer, 0, byteCount);
                }
                fontReference.Attribute(R.id).Value = ResourceID;
            }
        }

        private static void CopyChartObjects(ChartPart oldChart, ChartPart newChart)
        {
            foreach (XElement dataReference in newChart.GetXDocument().Descendants(C.externalData))
            {
                string relId = dataReference.Attribute(R.id).Value;

                var ipp1 = oldChart.Parts.FirstOrDefault(z => z.RelationshipId == relId);
                if (ipp1 != null)
                {
                    var oldRelatedPart = ipp1.OpenXmlPart;
                    if (oldRelatedPart is EmbeddedPackagePart)
                    {
                        EmbeddedPackagePart oldPart = (EmbeddedPackagePart)ipp1.OpenXmlPart;
                        EmbeddedPackagePart newPart = newChart.AddEmbeddedPackagePart(oldPart.ContentType);
                        using (Stream oldObject = oldPart.GetStream(FileMode.Open, FileAccess.Read))
                        using (Stream newObject = newPart.GetStream(FileMode.Create, FileAccess.ReadWrite))
                        {
                            int byteCount;
                            byte[] buffer = new byte[65536];
                            while ((byteCount = oldObject.Read(buffer, 0, 65536)) != 0)
                                newObject.Write(buffer, 0, byteCount);
                        }
                        dataReference.Attribute(R.id).Value = newChart.GetIdOfPart(newPart);
                    }
                    else if (oldRelatedPart is EmbeddedObjectPart)
                    {
                        EmbeddedObjectPart oldPart = (EmbeddedObjectPart)ipp1.OpenXmlPart;
                        var relType = oldRelatedPart.RelationshipType;
                        var conType = oldRelatedPart.ContentType;
                        var g = new Guid();
                        string id = $"R{g:N}".Substring(0, 8);
                        var newPart = newChart.AddExtendedPart(relType, conType, ".bin", id);
                        using (Stream oldObject = oldPart.GetStream(FileMode.Open, FileAccess.Read))
                        using (Stream newObject = newPart.GetStream(FileMode.Create, FileAccess.ReadWrite))
                        {
                            int byteCount;
                            byte[] buffer = new byte[65536];
                            while ((byteCount = oldObject.Read(buffer, 0, 65536)) != 0)
                                newObject.Write(buffer, 0, byteCount);
                        }
                        dataReference.Attribute(R.id).Value = newChart.GetIdOfPart(newPart);
                    }
                }
                else
                {
                    ExternalRelationship oldRelationship = oldChart.GetExternalRelationship(relId);
                    Guid g = Guid.NewGuid();
                    string newRid = $"R{g:N}";
                    var oldRel = oldChart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (oldRel == null)
                        throw new DocumentBuilderInternalException("Internal Error 0007");
                    newChart.AddExternalRelationship(oldRel.RelationshipType, oldRel.Uri, newRid);
                    dataReference.Attribute(R.id).Value = newRid;
                }
            }
        }

        private static void CopyStartingParts(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument,
            List<ImageData> images)
        {
            // A Core File Properties part does not have implicit or explicit relationships to other parts.
            CoreFilePropertiesPart corePart = sourceDocument.CoreFilePropertiesPart;
            if (corePart != null && corePart.GetXDocument().Root != null)
            {
                newDocument.AddCoreFilePropertiesPart();
                XDocument newXDoc = newDocument.CoreFilePropertiesPart.GetXDocument();
                newXDoc.Declaration.Standalone = Yes;
                newXDoc.Declaration.Encoding = Utf8;
                XDocument sourceXDoc = corePart.GetXDocument();
                newXDoc.Add(sourceXDoc.Root);
            }

            // An application attributes part does not have implicit or explicit relationships to other parts.
            ExtendedFilePropertiesPart extPart = sourceDocument.ExtendedFilePropertiesPart;
            if (extPart != null)
            {
                OpenXmlPart newPart = newDocument.AddExtendedFilePropertiesPart();
                XDocument newXDoc = newDocument.ExtendedFilePropertiesPart.GetXDocument();
                newXDoc.Declaration.Standalone = Yes;
                newXDoc.Declaration.Encoding = Utf8;
                newXDoc.Add(extPart.GetXDocument().Root);
            }

            // An custom file properties part does not have implicit or explicit relationships to other parts.
            CustomFilePropertiesPart customPart = sourceDocument.CustomFilePropertiesPart;
            if (customPart != null)
            {
                newDocument.AddCustomFilePropertiesPart();
                XDocument newXDoc = newDocument.CustomFilePropertiesPart.GetXDocument();
                newXDoc.Declaration.Standalone = Yes;
                newXDoc.Declaration.Encoding = Utf8;
                newXDoc.Add(customPart.GetXDocument().Root);
            }

            DocumentSettingsPart oldSettingsPart = sourceDocument.MainDocumentPart.DocumentSettingsPart;
            if (oldSettingsPart != null)
            {
                DocumentSettingsPart newSettingsPart = newDocument.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
                XDocument settingsXDoc = oldSettingsPart.GetXDocument();
                AddRelationships(oldSettingsPart, newSettingsPart, new[] { settingsXDoc.Root });
                CopyFootnotesPart(sourceDocument, newDocument, settingsXDoc, images);
                CopyEndnotesPart(sourceDocument, newDocument, settingsXDoc, images);
                XDocument newXDoc = newDocument.MainDocumentPart.DocumentSettingsPart.GetXDocument();
                newXDoc.Declaration.Standalone = Yes;
                newXDoc.Declaration.Encoding = Utf8;
                newXDoc.Add(settingsXDoc.Root);
                CopyRelatedPartsForContentParts(oldSettingsPart, newSettingsPart, new[] { newXDoc.Root }, images);
            }

            WebSettingsPart oldWebSettingsPart = sourceDocument.MainDocumentPart.WebSettingsPart;
            if (oldWebSettingsPart != null)
            {
                WebSettingsPart newWebSettingsPart = newDocument.MainDocumentPart.AddNewPart<WebSettingsPart>();
                XDocument settingsXDoc = oldWebSettingsPart.GetXDocument();
                AddRelationships(oldWebSettingsPart, newWebSettingsPart, new[] { settingsXDoc.Root });
                XDocument newXDoc = newDocument.MainDocumentPart.WebSettingsPart.GetXDocument();
                newXDoc.Declaration.Standalone = Yes;
                newXDoc.Declaration.Encoding = Utf8;
                newXDoc.Add(settingsXDoc.Root);
            }

            ThemePart themePart = sourceDocument.MainDocumentPart.ThemePart;
            if (themePart != null)
            {
                ThemePart newThemePart = newDocument.MainDocumentPart.AddNewPart<ThemePart>();
                XDocument newXDoc = newDocument.MainDocumentPart.ThemePart.GetXDocument();
                newXDoc.Declaration.Standalone = Yes;
                newXDoc.Declaration.Encoding = Utf8;
                newXDoc.Add(themePart.GetXDocument().Root);
                CopyRelatedPartsForContentParts(themePart, newThemePart, new[] { newThemePart.GetXDocument().Root }, images);
            }

            // If needed to handle GlossaryDocumentPart in the future, then
            // this code should handle the following parts:
            //   MainDocumentPart.GlossaryDocumentPart.StyleDefinitionsPart
            //   MainDocumentPart.GlossaryDocumentPart.StylesWithEffectsPart

            // A Style Definitions part shall not have implicit or explicit relationships to any other part.
            StyleDefinitionsPart stylesPart = sourceDocument.MainDocumentPart.StyleDefinitionsPart;
            if (stylesPart != null)
            {
                newDocument.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                XDocument newXDoc = newDocument.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
                newXDoc.Declaration.Standalone = Yes;
                newXDoc.Declaration.Encoding = Utf8;
                newXDoc.Add(new XElement(W.styles,
                    new XAttribute(XNamespace.Xmlns + "w", W.w)
                    
                    //,
                    //stylesPart.GetXDocument().Descendants(W.docDefaults)
                    
                    //,
                    //new XElement(W.latentStyles, stylesPart.GetXDocument().Descendants(W.latentStyles).Attributes())
                    
                    ));
                MergeDocDefaultStyles(stylesPart.GetXDocument(), newXDoc);
                MergeStyles(sourceDocument, newDocument, stylesPart.GetXDocument(), newXDoc, Enumerable.Empty<XElement>());
                MergeLatentStyles(stylesPart.GetXDocument(), newXDoc);
            }

            // A Font Table part shall not have any implicit or explicit relationships to any other part.
            FontTablePart fontTablePart = sourceDocument.MainDocumentPart.FontTablePart;
            if (fontTablePart != null)
            {
                newDocument.MainDocumentPart.AddNewPart<FontTablePart>();
                XDocument newXDoc = newDocument.MainDocumentPart.FontTablePart.GetXDocument();
                newXDoc.Declaration.Standalone = Yes;
                newXDoc.Declaration.Encoding = Utf8;
                CopyFontTable(sourceDocument.MainDocumentPart.FontTablePart, newDocument.MainDocumentPart.FontTablePart);
                newXDoc.Add(fontTablePart.GetXDocument().Root);
            }
        }

        private static void CopyFootnotesPart(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument,
            XDocument settingsXDoc, List<ImageData> images)
        {
            int number = 0;
            XDocument oldFootnotes = null;
            XDocument newFootnotes = null;
            XElement footnotePr = settingsXDoc.Root.Element(W.footnotePr);
            if (footnotePr == null)
                return;
            if (sourceDocument.MainDocumentPart.FootnotesPart == null)
                return;
            foreach (XElement footnote in footnotePr.Elements(W.footnote))
            {
                if (oldFootnotes == null)
                    oldFootnotes = sourceDocument.MainDocumentPart.FootnotesPart.GetXDocument();
                if (newFootnotes == null)
                {
                    if (newDocument.MainDocumentPart.FootnotesPart != null)
                    {
                        newFootnotes = newDocument.MainDocumentPart.FootnotesPart.GetXDocument();
                        newFootnotes.Declaration.Standalone = Yes;
                        newFootnotes.Declaration.Encoding = Utf8;
                        var ids = newFootnotes.Root.Elements(W.footnote).Select(f => (int)f.Attribute(W.id));
                        if (ids.Any())
                            number = ids.Max() + 1;
                    }
                    else
                    {
                        newDocument.MainDocumentPart.AddNewPart<FootnotesPart>();
                        newFootnotes = newDocument.MainDocumentPart.FootnotesPart.GetXDocument();
                        newFootnotes.Declaration.Standalone = Yes;
                        newFootnotes.Declaration.Encoding = Utf8;
                        newFootnotes.Add(new XElement(W.footnotes, NamespaceAttributes));
                    }
                }
                string id = (string)footnote.Attribute(W.id);
                XElement element = oldFootnotes.Descendants()
                    .Elements(W.footnote)
                    .Where(p => ((string)p.Attribute(W.id)) == id)
                    .FirstOrDefault();
                if (element != null)
                {
                    XElement newElement = new XElement(element);
                    // the following adds the footnote into the new settings part
                    newElement.Attribute(W.id).Value = number.ToString();
                    newFootnotes.Root.Add(newElement);
                    footnote.Attribute(W.id).Value = number.ToString();
                    number++;
                }
            }
        }

        private static void CopyEndnotesPart(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument,
            XDocument settingsXDoc, List<ImageData> images)
        {
            int number = 0;
            XDocument oldEndnotes = null;
            XDocument newEndnotes = null;
            XElement endnotePr = settingsXDoc.Root.Element(W.endnotePr);
            if (endnotePr == null)
                return;
            if (sourceDocument.MainDocumentPart.EndnotesPart == null)
                return;
            foreach (XElement endnote in endnotePr.Elements(W.endnote))
            {
                if (oldEndnotes == null)
                    oldEndnotes = sourceDocument.MainDocumentPart.EndnotesPart.GetXDocument();
                if (newEndnotes == null)
                {
                    if (newDocument.MainDocumentPart.EndnotesPart != null)
                    {
                        newEndnotes = newDocument.MainDocumentPart.EndnotesPart.GetXDocument();
                        newEndnotes.Declaration.Standalone = Yes;
                        newEndnotes.Declaration.Encoding = Utf8;
                        var ids = newEndnotes.Root
                            .Elements(W.endnote)
                            .Select(f => (int)f.Attribute(W.id));
                        if (ids.Any())
                            number = ids.Max() + 1;
                    }
                    else
                    {
                        newDocument.MainDocumentPart.AddNewPart<EndnotesPart>();
                        newEndnotes = newDocument.MainDocumentPart.EndnotesPart.GetXDocument();
                        newEndnotes.Declaration.Standalone = Yes;
                        newEndnotes.Declaration.Encoding = Utf8;
                        newEndnotes.Add(new XElement(W.endnotes, NamespaceAttributes));
                    }
                }
                string id = (string)endnote.Attribute(W.id);
                XElement element = oldEndnotes.Descendants()
                    .Elements(W.endnote)
                    .Where(p => ((string)p.Attribute(W.id)) == id)
                    .FirstOrDefault();
                if (element != null)
                {
                    XElement newElement = new XElement(element);
                    newElement.Attribute(W.id).Value = number.ToString();
                    newEndnotes.Root.Add(newElement);
                    endnote.Attribute(W.id).Value = number.ToString();
                    number++;
                }
            }
        }

        public static void FixRanges(XDocument sourceDocument, IEnumerable<XElement> newContent)
        {
            FixRange(sourceDocument,
                newContent,
                W.commentRangeStart,
                W.commentRangeEnd,
                W.id,
                W.commentReference);
            FixRange(sourceDocument,
                newContent,
                W.bookmarkStart,
                W.bookmarkEnd,
                W.id,
                null);
            FixRange(sourceDocument,
                newContent,
                W.permStart,
                W.permEnd,
                W.id,
                null);
            FixRange(sourceDocument,
                newContent,
                W.moveFromRangeStart,
                W.moveFromRangeEnd,
                W.id,
                null);
            FixRange(sourceDocument,
                newContent,
                W.moveToRangeStart,
                W.moveToRangeEnd,
                W.id,
                null);
            DeleteUnmatchedRange(sourceDocument,
                newContent,
                W.moveFromRangeStart,
                W.moveFromRangeEnd,
                W.moveToRangeStart,
                W.name,
                W.id);
            DeleteUnmatchedRange(sourceDocument,
                newContent,
                W.moveToRangeStart,
                W.moveToRangeEnd,
                W.moveFromRangeStart,
                W.name,
                W.id);
        }

        private static void AddAtBeginning(IEnumerable<XElement> newContent, XElement contentToAdd)
        {
            if (newContent.First().Element(W.pPr) != null)
                newContent.First().Element(W.pPr).AddAfterSelf(contentToAdd);
            else
                newContent.First().AddFirst(new XElement(contentToAdd));
        }

        private static void AddAtEnd(IEnumerable<XElement> newContent, XElement contentToAdd)
        {
            if (newContent.Last().Element(W.pPr) != null)
                newContent.Last().Element(W.pPr).AddAfterSelf(new XElement(contentToAdd));
            else
                newContent.Last().Add(new XElement(contentToAdd));
        }

        // If the set of paragraphs from sourceDocument don't have a complete start/end for bookmarks,
        // comments, etc., then this adds them to the paragraph.  Note that this adds them to
        // sourceDocument, and is impure.
        private static void FixRange(XDocument sourceDocument, IEnumerable<XElement> newContent,
            XName startElement, XName endElement, XName idAttribute, XName refElement)
        {
            foreach (XElement start in newContent.DescendantsAndSelf(startElement))
            {
                string rangeId = start.Attribute(idAttribute).Value;
                if (newContent
                    .DescendantsAndSelf(endElement)
                    .Where(e => e.Attribute(idAttribute).Value == rangeId)
                    .Count() == 0)
                {
                    XElement end = sourceDocument
                        .Descendants(endElement)
                        .Where(o => o.Attribute(idAttribute).Value == rangeId)
                        .FirstOrDefault();
                    if (end != null)
                    {
                        AddAtEnd(newContent, new XElement(end));
                        if (refElement != null)
                        {
                            XElement newRef = new XElement(refElement, new XAttribute(idAttribute, rangeId));
                            AddAtEnd(newContent, new XElement(newRef));
                        }
                    }
                }
            }
            foreach (XElement end in newContent.Elements(endElement))
            {
                string rangeId = end.Attribute(idAttribute).Value;
                if (newContent
                    .DescendantsAndSelf(startElement)
                    .Where(s => s.Attribute(idAttribute).Value == rangeId)
                    .Count() == 0)
                {
                    XElement start = sourceDocument
                        .Descendants(startElement)
                        .Where(o => o.Attribute(idAttribute).Value == rangeId)
                        .FirstOrDefault();
                    if (start != null)
                        AddAtBeginning(newContent, new XElement(start));
                }
            }
        }

        private static void DeleteUnmatchedRange(XDocument sourceDocument, IEnumerable<XElement> newContent,
            XName startElement, XName endElement, XName matchTo, XName matchAttr, XName idAttr)
        {
            List<string> deleteList = new List<string>();
            foreach (XElement start in newContent.Elements(startElement))
            {
                string id = start.Attribute(matchAttr).Value;
                if (!newContent.Elements(matchTo).Where(n => n.Attribute(matchAttr).Value == id).Any())
                    deleteList.Add(start.Attribute(idAttr).Value);
            }
            foreach (string item in deleteList)
            {
                newContent.Elements(startElement).Where(n => n.Attribute(idAttr).Value == item).Remove();
                newContent.Elements(endElement).Where(n => n.Attribute(idAttr).Value == item).Remove();
                newContent.Where(p => p.Name == startElement && p.Attribute(idAttr).Value == item).Remove();
                newContent.Where(p => p.Name == endElement && p.Attribute(idAttr).Value == item).Remove();
            }
        }

        private static void CopyFootnotes(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument,
            IEnumerable<XElement> newContent, List<ImageData> images)
        {
            int number = 0;
            XDocument oldFootnotes = null;
            XDocument newFootnotes = null;
            foreach (XElement footnote in newContent.DescendantsAndSelf(W.footnoteReference))
            {
                if (oldFootnotes == null)
                    oldFootnotes = sourceDocument.MainDocumentPart.FootnotesPart.GetXDocument();
                if (newFootnotes == null)
                {
                    if (newDocument.MainDocumentPart.FootnotesPart != null)
                    {
                        newFootnotes = newDocument.MainDocumentPart.FootnotesPart.GetXDocument();
                        var ids = newFootnotes
                            .Root
                            .Elements(W.footnote)
                            .Select(f => (int)f.Attribute(W.id));
                        if (ids.Any())
                            number = ids.Max() + 1;
                    }
                    else
                    {
                        newDocument.MainDocumentPart.AddNewPart<FootnotesPart>();
                        newFootnotes = newDocument.MainDocumentPart.FootnotesPart.GetXDocument();
                        newFootnotes.Declaration.Standalone = Yes;
                        newFootnotes.Declaration.Encoding = Utf8;
                        newFootnotes.Add(new XElement(W.footnotes, NamespaceAttributes));
                    }
                }
                string id = (string)footnote.Attribute(W.id);
                XElement element = oldFootnotes
                    .Descendants()
                    .Elements(W.footnote)
                    .Where(p => ((string)p.Attribute(W.id)) == id)
                    .FirstOrDefault();
                if (element != null)
                {
                    XElement newElement = new XElement(element);
                    newElement.Attribute(W.id).Value = number.ToString();
                    newFootnotes.Root.Add(newElement);
                    footnote.Attribute(W.id).Value = number.ToString();
                    number++;
                }
            }
            if (sourceDocument.MainDocumentPart.FootnotesPart != null &&
                newDocument.MainDocumentPart.FootnotesPart != null)
            {
                AddRelationships(sourceDocument.MainDocumentPart.FootnotesPart,
                    newDocument.MainDocumentPart.FootnotesPart,
                    new[] { newDocument.MainDocumentPart.FootnotesPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(sourceDocument.MainDocumentPart.FootnotesPart,
                    newDocument.MainDocumentPart.FootnotesPart,
                    new[] { newDocument.MainDocumentPart.FootnotesPart.GetXDocument().Root }, images);
            }
        }

        private static void CopyEndnotes(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument,
            IEnumerable<XElement> newContent, List<ImageData> images)
        {
            int number = 0;
            XDocument oldEndnotes = null;
            XDocument newEndnotes = null;
            foreach (XElement endnote in newContent.DescendantsAndSelf(W.endnoteReference))
            {
                if (oldEndnotes == null)
                    oldEndnotes = sourceDocument.MainDocumentPart.EndnotesPart.GetXDocument();
                if (newEndnotes == null)
                {
                    if (newDocument.MainDocumentPart.EndnotesPart != null)
                    {
                        newEndnotes = newDocument
                            .MainDocumentPart
                            .EndnotesPart
                            .GetXDocument();
                        var ids = newEndnotes
                            .Root
                            .Elements(W.endnote)
                            .Select(f => (int)f.Attribute(W.id));
                        if (ids.Any())
                            number = ids.Max() + 1;
                    }
                    else
                    {
                        newDocument.MainDocumentPart.AddNewPart<EndnotesPart>();
                        newEndnotes = newDocument.MainDocumentPart.EndnotesPart.GetXDocument();
                        newEndnotes.Declaration.Standalone = Yes;
                        newEndnotes.Declaration.Encoding = Utf8;
                        newEndnotes.Add(new XElement(W.endnotes, NamespaceAttributes));
                    }
                }
                string id = (string)endnote.Attribute(W.id);
                XElement element = oldEndnotes
                    .Descendants()
                    .Elements(W.endnote)
                    .Where(p => ((string)p.Attribute(W.id)) == id)
                    .First();
                XElement newElement = new XElement(element);
                newElement.Attribute(W.id).Value = number.ToString();
                newEndnotes.Root.Add(newElement);
                endnote.Attribute(W.id).Value = number.ToString();
                number++;
            }
            if (sourceDocument.MainDocumentPart.EndnotesPart != null &&
                newDocument.MainDocumentPart.EndnotesPart != null)
            {
                AddRelationships(sourceDocument.MainDocumentPart.EndnotesPart,
                    newDocument.MainDocumentPart.EndnotesPart,
                    new[] { newDocument.MainDocumentPart.EndnotesPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(sourceDocument.MainDocumentPart.EndnotesPart,
                    newDocument.MainDocumentPart.EndnotesPart,
                    new[] { newDocument.MainDocumentPart.EndnotesPart.GetXDocument().Root }, images);
            }
        }

        // General function for handling images that tries to use an existing image if they are the same
        private static ImageData ManageImageCopy(ImagePart oldImage, OpenXmlPart newContentPart, List<ImageData> images)
        {
            ImageData oldImageData = new ImageData(oldImage);
            foreach (ImageData item in images)
            {
                if (newContentPart != item.ImagePart)
                    continue;
                if (item.Compare(oldImageData))
                    return item;
            }
            images.Add(oldImageData);
            return oldImageData;
        }

        private static XAttribute[] NamespaceAttributes =
        {
            new XAttribute(XNamespace.Xmlns + "wpc", WPC.wpc),
            new XAttribute(XNamespace.Xmlns + "mc", MC.mc),
            new XAttribute(XNamespace.Xmlns + "o", O.o),
            new XAttribute(XNamespace.Xmlns + "r", R.r),
            new XAttribute(XNamespace.Xmlns + "m", M.m),
            new XAttribute(XNamespace.Xmlns + "v", VML.vml),
            new XAttribute(XNamespace.Xmlns + "wp14", WP14.wp14),
            new XAttribute(XNamespace.Xmlns + "wp", WP.wp),
            new XAttribute(XNamespace.Xmlns + "w10", W10.w10),
            new XAttribute(XNamespace.Xmlns + "w", W.w),
            new XAttribute(XNamespace.Xmlns + "w14", W14.w14),
            new XAttribute(XNamespace.Xmlns + "wpg", WPG.wpg),
            new XAttribute(XNamespace.Xmlns + "wpi", WPI.wpi),
            new XAttribute(XNamespace.Xmlns + "wne", WNE.wne),
            new XAttribute(XNamespace.Xmlns + "wps", WPS.wps),
            new XAttribute(MC.Ignorable, "w14 wp14"),
        };
    }

    public class DocumentBuilderException : Exception
    {
        public DocumentBuilderException(string message) : base(message) { }
    }

    public class DocumentBuilderInternalException : Exception
    {
        public DocumentBuilderInternalException(string message) : base(message) { }
    }
}
