// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.IO.Packaging;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Linq;
using System.Xml.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using System.Drawing;
using Font = System.Drawing.Font;
using FontFamily = System.Drawing.FontFamily;

// ReSharper disable InconsistentNaming

namespace OpenXmlPowerTools
{
    public static class PtOpenXmlExtensions
    {
        public static XDocument GetXDocument(this OpenXmlPart part)
        {
            if (part == null) throw new ArgumentNullException("part");

            XDocument partXDocument = part.Annotation<XDocument>();
            if (partXDocument != null) return partXDocument;

            using (Stream partStream = part.GetStream())
            {
                if (partStream.Length == 0)
                {
                    partXDocument = new XDocument();
                    partXDocument.Declaration = new XDeclaration("1.0", "UTF-8", "yes");
                }
                else
                {
                    using (XmlReader partXmlReader = XmlReader.Create(partStream))
                        partXDocument = XDocument.Load(partXmlReader);
                }
            }

            part.AddAnnotation(partXDocument);
            return partXDocument;
        }

        public static XDocument GetXDocument(this OpenXmlPart part, out XmlNamespaceManager namespaceManager)
        {
            if (part == null) throw new ArgumentNullException("part");

            namespaceManager = part.Annotation<XmlNamespaceManager>();
            XDocument partXDocument = part.Annotation<XDocument>();
            if (partXDocument != null)
            {
                if (namespaceManager != null) return partXDocument;

                namespaceManager = GetManagerFromXDocument(partXDocument);
                part.AddAnnotation(namespaceManager);

                return partXDocument;
            }

            using (Stream partStream = part.GetStream())
            {
                if (partStream.Length == 0)
                {
                    partXDocument = new XDocument();
                    partXDocument.Declaration = new XDeclaration("1.0", "UTF-8", "yes");

                    part.AddAnnotation(partXDocument);

                    return partXDocument;
                }
                else
                {
                    using (XmlReader partXmlReader = XmlReader.Create(partStream))
                    {
                        partXDocument = XDocument.Load(partXmlReader);
                        namespaceManager = new XmlNamespaceManager(partXmlReader.NameTable);

                        part.AddAnnotation(partXDocument);
                        part.AddAnnotation(namespaceManager);

                        return partXDocument;
                    }
                }
            }
        }

        public static void PutXDocument(this OpenXmlPart part)
        {
            if (part == null) throw new ArgumentNullException("part");

            XDocument partXDocument = part.GetXDocument();
            if (partXDocument != null)
            {
#if true
                using (Stream partStream = part.GetStream(FileMode.Create, FileAccess.Write))
                using (XmlWriter partXmlWriter = XmlWriter.Create(partStream))
                    partXDocument.Save(partXmlWriter);
#else
                byte[] array = Encoding.UTF8.GetBytes(partXDocument.ToString(SaveOptions.DisableFormatting));
                using (MemoryStream ms = new MemoryStream(array))
                    part.FeedData(ms);
#endif
            }
        }

        public static void PutXDocumentWithFormatting(this OpenXmlPart part)
        {
            if (part == null) throw new ArgumentNullException("part");

            XDocument partXDocument = part.GetXDocument();
            if (partXDocument != null)
            {
                using (Stream partStream = part.GetStream(FileMode.Create, FileAccess.Write))
                {
                    XmlWriterSettings settings = new XmlWriterSettings();
                    settings.Indent = true;
                    settings.OmitXmlDeclaration = true;
                    settings.NewLineOnAttributes = true;
                    using (XmlWriter partXmlWriter = XmlWriter.Create(partStream, settings))
                        partXDocument.Save(partXmlWriter);
                }
            }
        }

        public static void PutXDocument(this OpenXmlPart part, XDocument document)
        {
            if (part == null) throw new ArgumentNullException("part");
            if (document == null) throw new ArgumentNullException("document");

            using (Stream partStream = part.GetStream(FileMode.Create, FileAccess.Write))
            using (XmlWriter partXmlWriter = XmlWriter.Create(partStream))
                document.Save(partXmlWriter);

            part.RemoveAnnotations<XDocument>();
            part.AddAnnotation(document);
        }

        private static XmlNamespaceManager GetManagerFromXDocument(XDocument xDocument)
        {
            XmlReader reader = xDocument.CreateReader();
            XDocument newXDoc = XDocument.Load(reader);

            XElement rootElement = xDocument.Elements().FirstOrDefault();
            rootElement.ReplaceWith(newXDoc.Root);

            XmlNameTable nameTable = reader.NameTable;
            XmlNamespaceManager namespaceManager = new XmlNamespaceManager(nameTable);
            return namespaceManager;
        }

        public static IEnumerable<XElement> LogicalChildrenContent(this XElement element)
        {
            if (element.Name == W.document)
                return element.Descendants(W.body).Take(1);

            if (element.Name == W.body ||
                element.Name == W.tc ||
                element.Name == W.txbxContent)
                return element
                    .DescendantsTrimmed(e =>
                        e.Name == W.tbl ||
                        e.Name == W.p)
                    .Where(e =>
                        e.Name == W.p ||
                        e.Name == W.tbl);

            if (element.Name == W.tbl)
                return element
                    .DescendantsTrimmed(W.tr)
                    .Where(e => e.Name == W.tr);

            if (element.Name == W.tr)
                return element
                    .DescendantsTrimmed(W.tc)
                    .Where(e => e.Name == W.tc);

            if (element.Name == W.p)
                return element
                    .DescendantsTrimmed(e => e.Name == W.r ||
                        e.Name == W.pict ||
                        e.Name == W.drawing)
                    .Where(e => e.Name == W.r ||
                        e.Name == W.pict ||
                        e.Name == W.drawing);

            if (element.Name == W.r)
                return element
                    .DescendantsTrimmed(e => W.SubRunLevelContent.Contains(e.Name))
                    .Where(e => W.SubRunLevelContent.Contains(e.Name));

            if (element.Name == MC.AlternateContent)
                return element
                    .DescendantsTrimmed(e =>
                        e.Name == W.pict ||
                        e.Name == W.drawing ||
                        e.Name == MC.Fallback)
                    .Where(e =>
                        e.Name == W.pict ||
                        e.Name == W.drawing);

            if (element.Name == W.pict || element.Name == W.drawing)
                return element
                    .DescendantsTrimmed(W.txbxContent)
                    .Where(e => e.Name == W.txbxContent);

            return XElement.EmptySequence;
        }

        public static IEnumerable<XElement> LogicalChildrenContent(this IEnumerable<XElement> source)
        {
            foreach (XElement e1 in source)
                foreach (XElement e2 in e1.LogicalChildrenContent())
                    yield return e2;
        }

        public static IEnumerable<XElement> LogicalChildrenContent(this XElement element, XName name)
        {
            return element.LogicalChildrenContent().Where(e => e.Name == name);
        }

        public static IEnumerable<XElement> LogicalChildrenContent(this IEnumerable<XElement> source, XName name)
        {
            foreach (XElement e1 in source)
                foreach (XElement e2 in e1.LogicalChildrenContent(name))
                    yield return e2;
        }

        public static IEnumerable<OpenXmlPart> ContentParts(this WordprocessingDocument doc)
        {
            yield return doc.MainDocumentPart;

            foreach (var hdr in doc.MainDocumentPart.HeaderParts)
                yield return hdr;

            foreach (var ftr in doc.MainDocumentPart.FooterParts)
                yield return ftr;

            if (doc.MainDocumentPart.FootnotesPart != null)
                yield return doc.MainDocumentPart.FootnotesPart;

            if (doc.MainDocumentPart.EndnotesPart != null)
                yield return doc.MainDocumentPart.EndnotesPart;
        }

        /// <summary>
        /// Creates a complete list of all parts contained in the <see cref="OpenXmlPartContainer"/>.
        /// </summary>
        /// <param name="container">
        /// A <see cref="WordprocessingDocument"/>, <see cref="SpreadsheetDocument"/>, or
        /// <see cref="PresentationDocument"/>.
        /// </param>
        /// <returns>list of <see cref="OpenXmlPart"/>s contained in the <see cref="OpenXmlPartContainer"/>.</returns>
        public static List<OpenXmlPart> GetAllParts(this OpenXmlPartContainer container)
        {
            // Use a HashSet so that parts are processed only once.
            HashSet<OpenXmlPart> partList = new HashSet<OpenXmlPart>();

            foreach (IdPartPair p in container.Parts)
                AddPart(partList, p.OpenXmlPart);

            return partList.OrderBy(p => p.ContentType).ThenBy(p => p.Uri.ToString()).ToList();
        }

        private static void AddPart(HashSet<OpenXmlPart> partList, OpenXmlPart part)
        {
            if (partList.Contains(part)) return;

            partList.Add(part);
            foreach (IdPartPair p in part.Parts)
                AddPart(partList, p.OpenXmlPart);
        }
    }

    public static class FlatOpc
    {
        private class FlatOpcTupple
        {
            public char FoCharacter;
            public int FoChunk;
        }

        private static XElement GetContentsAsXml(PackagePart part)
        {
            XNamespace pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";

            if (part.ContentType.EndsWith("xml"))
            {
                using (Stream str = part.GetStream())
                using (StreamReader streamReader = new StreamReader(str))
                using (XmlReader xr = XmlReader.Create(streamReader))
                    return new XElement(pkg + "part",
                        new XAttribute(pkg + "name", part.Uri),
                        new XAttribute(pkg + "contentType", part.ContentType),
                        new XElement(pkg + "xmlData",
                            XElement.Load(xr)
                        )
                    );
            }
            else
            {
                using (Stream str = part.GetStream())
                using (BinaryReader binaryReader = new BinaryReader(str))
                {
                    int len = (int)binaryReader.BaseStream.Length;
                    byte[] byteArray = binaryReader.ReadBytes(len);
                    // the following expression creates the base64String, then chunks
                    // it to lines of 76 characters long
                    string base64String = (System.Convert.ToBase64String(byteArray))
                        .Select
                        (
                            (c, i) => new FlatOpcTupple()
                            {
                                FoCharacter = c,
                                FoChunk = i / 76
                            }
                        )
                        .GroupBy(c => c.FoChunk)
                        .Aggregate(
                            new StringBuilder(),
                            (s, i) =>
                                s.Append(
                                    i.Aggregate(
                                        new StringBuilder(),
                                        (seed, it) => seed.Append(it.FoCharacter),
                                        sb => sb.ToString()
                                    )
                                )
                                .Append(Environment.NewLine),
                            s => s.ToString()
                        );
                    return new XElement(pkg + "part",
                        new XAttribute(pkg + "name", part.Uri),
                        new XAttribute(pkg + "contentType", part.ContentType),
                        new XAttribute(pkg + "compression", "store"),
                        new XElement(pkg + "binaryData", base64String)
                    );
                }
            }
        }

        private static XProcessingInstruction GetProcessingInstruction(string path)
        {
            var fi = new FileInfo(path);
            if (Util.IsWordprocessingML(fi.Extension))
                return new XProcessingInstruction("mso-application",
                            "progid=\"Word.Document\"");
            if (Util.IsPresentationML(fi.Extension))
                return new XProcessingInstruction("mso-application",
                            "progid=\"PowerPoint.Show\"");
            if (Util.IsSpreadsheetML(fi.Extension))
                return new XProcessingInstruction("mso-application",
                            "progid=\"Excel.Sheet\"");
            return null;
        }

        public static XmlDocument OpcToXmlDocument(string fileName)
        {
            using (Package package = Package.Open(fileName))
            {
                XNamespace pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";

                XDeclaration declaration = new XDeclaration("1.0", "UTF-8", "yes");
                XDocument doc = new XDocument(
                    declaration,
                    GetProcessingInstruction(fileName),
                    new XElement(pkg + "package",
                        new XAttribute(XNamespace.Xmlns + "pkg", pkg.ToString()),
                        package.GetParts().Select(part => GetContentsAsXml(part))
                    )
                );
                return GetXmlDocument(doc);
            }
        }

        public static XDocument OpcToXDocument(string fileName)
        {
            using (Package package = Package.Open(fileName))
            {
                XNamespace pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";

                XDeclaration declaration = new XDeclaration("1.0", "UTF-8", "yes");
                XDocument doc = new XDocument(
                    declaration,
                    GetProcessingInstruction(fileName),
                    new XElement(pkg + "package",
                        new XAttribute(XNamespace.Xmlns + "pkg", pkg.ToString()),
                        package.GetParts().Select(part => GetContentsAsXml(part))
                    )
                );
                return doc;
            }
        }

        public static string[] OpcToText(string fileName)
        {
            using (Package package = Package.Open(fileName))
            {
                XNamespace pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";

                XDeclaration declaration = new XDeclaration("1.0", "UTF-8", "yes");
                XDocument doc = new XDocument(
                    declaration,
                    GetProcessingInstruction(fileName),
                    new XElement(pkg + "package",
                        new XAttribute(XNamespace.Xmlns + "pkg", pkg.ToString()),
                        package.GetParts().Select(part => GetContentsAsXml(part))
                    )
                );
                return doc.ToString().Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            }
        }

        private static XmlDocument GetXmlDocument(XDocument document)
        {
            using (XmlReader xmlReader = document.CreateReader())
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(xmlReader);
                if (document.Declaration != null)
                {
                    XmlDeclaration dec = xmlDoc.CreateXmlDeclaration(document.Declaration.Version,
                        document.Declaration.Encoding, document.Declaration.Standalone);
                    xmlDoc.InsertBefore(dec, xmlDoc.FirstChild);
                }
                return xmlDoc;
            }
        }

        private static XDocument GetXDocument(this XmlDocument document)
        {
            XDocument xDoc = new XDocument();
            using (XmlWriter xmlWriter = xDoc.CreateWriter())
                document.WriteTo(xmlWriter);
            XmlDeclaration decl =
                document.ChildNodes.OfType<XmlDeclaration>().FirstOrDefault();
            if (decl != null)
                xDoc.Declaration = new XDeclaration(decl.Version, decl.Encoding,
                    decl.Standalone);
            return xDoc;
        }

        public static void FlatToOpc(XmlDocument doc, string outputPath)
        {
            XDocument xd = GetXDocument(doc);
            FlatToOpc(xd, outputPath);
        }

        public static void FlatToOpc(string xmlText, string outputPath)
        {
            XDocument xd = XDocument.Parse(xmlText);
            FlatToOpc(xd, outputPath);
        }

        public static void FlatToOpc(XDocument doc, string outputPath)
        {
            XNamespace pkg =
                "http://schemas.microsoft.com/office/2006/xmlPackage";
            XNamespace rel =
                "http://schemas.openxmlformats.org/package/2006/relationships";

            using (Package package = Package.Open(outputPath, FileMode.Create))
            {
                // add all parts (but not relationships)
                foreach (var xmlPart in doc.Root
                    .Elements()
                    .Where(p =>
                        (string)p.Attribute(pkg + "contentType") !=
                        "application/vnd.openxmlformats-package.relationships+xml"))
                {
                    string name = (string)xmlPart.Attribute(pkg + "name");
                    string contentType = (string)xmlPart.Attribute(pkg + "contentType");
                    if (contentType.EndsWith("xml"))
                    {
                        Uri u = new Uri(name, UriKind.Relative);
                        PackagePart part = package.CreatePart(u, contentType,
                            CompressionOption.SuperFast);
                        using (Stream str = part.GetStream(FileMode.Create))
                        using (XmlWriter xmlWriter = XmlWriter.Create(str))
                            xmlPart.Element(pkg + "xmlData")
                                .Elements()
                                .First()
                                .WriteTo(xmlWriter);
                    }
                    else
                    {
                        Uri u = new Uri(name, UriKind.Relative);
                        PackagePart part = package.CreatePart(u, contentType,
                            CompressionOption.SuperFast);
                        using (Stream str = part.GetStream(FileMode.Create))
                        using (BinaryWriter binaryWriter = new BinaryWriter(str))
                        {
                            string base64StringInChunks =
                                (string)xmlPart.Element(pkg + "binaryData");
                            char[] base64CharArray = base64StringInChunks
                                .Where(c => c != '\r' && c != '\n').ToArray();
                            byte[] byteArray =
                                System.Convert.FromBase64CharArray(base64CharArray,
                                0, base64CharArray.Length);
                            binaryWriter.Write(byteArray);
                        }
                    }
                }

                foreach (var xmlPart in doc.Root.Elements())
                {
                    string name = (string)xmlPart.Attribute(pkg + "name");
                    string contentType = (string)xmlPart.Attribute(pkg + "contentType");
                    if (contentType ==
                        "application/vnd.openxmlformats-package.relationships+xml")
                    {
                        // add the package level relationships
                        if (name == "/_rels/.rels")
                        {
                            foreach (XElement xmlRel in
                                xmlPart.Descendants(rel + "Relationship"))
                            {
                                string id = (string)xmlRel.Attribute("Id");
                                string type = (string)xmlRel.Attribute("Type");
                                string target = (string)xmlRel.Attribute("Target");
                                string targetMode =
                                    (string)xmlRel.Attribute("TargetMode");
                                if (targetMode == "External")
                                    package.CreateRelationship(
                                        new Uri(target, UriKind.Absolute),
                                        TargetMode.External, type, id);
                                else
                                    package.CreateRelationship(
                                        new Uri(target, UriKind.Relative),
                                        TargetMode.Internal, type, id);
                            }
                        }
                        else
                        // add part level relationships
                        {
                            string directory = name.Substring(0, name.IndexOf("/_rels"));
                            string relsFilename = name.Substring(name.LastIndexOf('/'));
                            string filename =
                                relsFilename.Substring(0, relsFilename.IndexOf(".rels"));
                            PackagePart fromPart = package.GetPart(
                                new Uri(directory + filename, UriKind.Relative));
                            foreach (XElement xmlRel in
                                xmlPart.Descendants(rel + "Relationship"))
                            {
                                string id = (string)xmlRel.Attribute("Id");
                                string type = (string)xmlRel.Attribute("Type");
                                string target = (string)xmlRel.Attribute("Target");
                                string targetMode =
                                    (string)xmlRel.Attribute("TargetMode");
                                if (targetMode == "External")
                                    fromPart.CreateRelationship(
                                        new Uri(target, UriKind.Absolute),
                                        TargetMode.External, type, id);
                                else
                                    fromPart.CreateRelationship(
                                        new Uri(target, UriKind.Relative),
                                        TargetMode.Internal, type, id);
                            }
                        }
                    }
                }
            }
        }
    }

    public class Base64
    {
        public static string ConvertToBase64(string fileName)
        {
            byte[] ba = System.IO.File.ReadAllBytes(fileName);
            string base64String = (System.Convert.ToBase64String(ba))
                .Select
                (
                    (c, i) => new
                    {
                        Chunk = i / 76,
                        Character = c
                    }
                )
                .GroupBy(c => c.Chunk)
                .Aggregate(
                    new StringBuilder(),
                    (s, i) =>
                        s.Append(
                            i.Aggregate(
                                new StringBuilder(),
                                (seed, it) => seed.Append(it.Character),
                                sb => sb.ToString()
                            )
                        )
                        .Append(Environment.NewLine),
                    s =>
                    {
                        s.Length -= Environment.NewLine.Length;
                        return s.ToString();
                    }
                );

            return base64String;
        }

        public static byte[] ConvertFromBase64(string fileName, string b64)
        {
            string b64b = b64.Replace("\r\n", "");
            byte[] ba = System.Convert.FromBase64String(b64b);
            return ba;
        }
    }

    public static class XmlUtil
    {
        public static XAttribute GetXmlSpaceAttribute(string value)
        {
            return (value.Length > 0) && ((value[0] == ' ') || (value[value.Length - 1] == ' '))
                ? new XAttribute(XNamespace.Xml + "space", "preserve")
                : null;
        }

        public static XAttribute GetXmlSpaceAttribute(char value)
        {
            return value == ' ' ? new XAttribute(XNamespace.Xml + "space", "preserve") : null;
        }
    }

    public static class WordprocessingMLUtil
    {
        private static HashSet<string> UnknownFonts = new HashSet<string>();
        private static HashSet<string> KnownFamilies = null;

        public static int CalcWidthOfRunInTwips(XElement r)
        {
            if (KnownFamilies == null)
            {
                KnownFamilies = new HashSet<string>();
                var families = FontFamily.Families;
                foreach (var fam in families)
                    KnownFamilies.Add(fam.Name);
            }

            var fontName = (string)r.Attribute(PtOpenXml.pt + "FontName");
            if (fontName == null)
                fontName = (string)r.Ancestors(W.p).First().Attribute(PtOpenXml.pt + "FontName");
            if (fontName == null)
                throw new OpenXmlPowerToolsException("Internal Error, should have FontName attribute");
            if (UnknownFonts.Contains(fontName))
                return 0;

            var rPr = r.Element(W.rPr);
            if (rPr == null)
                throw new OpenXmlPowerToolsException("Internal Error, should have run properties");
            var languageType = (string)r.Attribute(PtOpenXml.LanguageType);
            decimal? szn = null;
            if (languageType == "bidi")
                szn = (decimal?)rPr.Elements(W.szCs).Attributes(W.val).FirstOrDefault();
            else
                szn = (decimal?)rPr.Elements(W.sz).Attributes(W.val).FirstOrDefault();
            if (szn == null)
                szn = 22m;

            var sz = szn.GetValueOrDefault();

            // unknown font families will throw ArgumentException, in which case just return 0
            if (!KnownFamilies.Contains(fontName))
                return 0;
            // in theory, all unknown fonts are found by the above test, but if not...
            FontFamily ff;
            try
            {
                ff = new FontFamily(fontName);
            }
            catch (ArgumentException)
            {
                UnknownFonts.Add(fontName);

                return 0;
            }
            FontStyle fs = FontStyle.Regular;
            var bold = GetBoolProp(rPr, W.b) || GetBoolProp(rPr, W.bCs);
            var italic = GetBoolProp(rPr, W.i) || GetBoolProp(rPr, W.iCs);
            if (bold && !italic)
                fs = FontStyle.Bold;
            if (italic && !bold)
                fs = FontStyle.Italic;
            if (bold && italic)
                fs = FontStyle.Bold | FontStyle.Italic;

            var runText = r.DescendantsTrimmed(W.txbxContent)
                .Where(e => e.Name == W.t)
                .Select(t => (string)t)
                .StringConcatenate();

            var tabLength = r.DescendantsTrimmed(W.txbxContent)
                .Where(e => e.Name == W.tab)
                .Select(t => (decimal)t.Attribute(PtOpenXml.TabWidth))
                .Sum();

            if (runText.Length == 0 && tabLength == 0)
                return 0;

            int multiplier = 1;
            if (runText.Length <= 2)
                multiplier = 100;
            else if (runText.Length <= 4)
                multiplier = 50;
            else if (runText.Length <= 8)
                multiplier = 25;
            else if (runText.Length <= 16)
                multiplier = 12;
            else if (runText.Length <= 32)
                multiplier = 6;
            if (multiplier != 1)
            {
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < multiplier; i++)
                    sb.Append(runText);
                runText = sb.ToString();
            }

            var w = MetricsGetter.GetTextWidth(ff, fs, sz, runText);

            return (int) (w / 96m * 1440m / multiplier + tabLength * 1440m);
        }

        public static bool GetBoolProp(XElement runProps, XName xName)
        {
            var p = runProps.Element(xName);
            if (p == null)
                return false;
            var v = p.Attribute(W.val);
            if (v == null)
                return true;
            var s = v.Value.ToLower();
            if (s == "0" || s == "false")
                return false;
            if (s == "1" || s == "true")
                return true;
            return false;
        }

        private static readonly List<XName> AdditionalRunContainerNames = new List<XName>
        {
            W.w + "bdo",
            W.customXml,
            W.dir,
            W.fldSimple,
            W.hyperlink,
            W.moveFrom,
            W.moveTo,
            W.sdtContent
        };

        public static XElement CoalesceAdjacentRunsWithIdenticalFormatting(XElement runContainer)
        {
            const string dontConsolidate = "DontConsolidate";

            IEnumerable<IGrouping<string, XElement>> groupedAdjacentRunsWithIdenticalFormatting =
                runContainer
                    .Elements()
                    .GroupAdjacent(ce =>
                    {
                        if (ce.Name == W.r)
                        {
                            if (ce.Elements().Count(e => e.Name != W.rPr) != 1)
                                return dontConsolidate;

                            XElement rPr = ce.Element(W.rPr);
                            string rPrString = rPr != null ? rPr.ToString(SaveOptions.None) : string.Empty;

                            if (ce.Element(W.t) != null)
                                return "Wt" + rPrString;

                            if (ce.Element(W.instrText) != null)
                                return "WinstrText" + rPrString;

                            return dontConsolidate;
                        }

                        if (ce.Name == W.ins)
                        {
                            if (ce.Elements(W.del).Any())
                            {
                                return dontConsolidate;
#if false
                                // for w:ins/w:del/w:r/w:delText
                                if ((ce.Elements(W.del).Elements(W.r).Elements().Count(e => e.Name != W.rPr) != 1) ||
                                    !ce.Elements().Elements().Elements(W.delText).Any())
                                    return dontConsolidate;

                                XAttribute dateIns = ce.Attribute(W.date);
                                XElement del = ce.Element(W.del);
                                XAttribute dateDel = del.Attribute(W.date);

                                string authorIns = (string) ce.Attribute(W.author) ?? string.Empty;
                                string dateInsString = dateIns != null
                                    ? ((DateTime) dateIns).ToString("s")
                                    : string.Empty;
                                string authorDel = (string) del.Attribute(W.author) ?? string.Empty;
                                string dateDelString = dateDel != null
                                    ? ((DateTime) dateDel).ToString("s")
                                    : string.Empty;

                                return "Wins" +
                                       authorIns +
                                       dateInsString +
                                       authorDel +
                                       dateDelString +
                                       ce.Elements(W.del)
                                           .Elements(W.r)
                                           .Elements(W.rPr)
                                           .Select(rPr => rPr.ToString(SaveOptions.None))
                                           .StringConcatenate();
#endif
                            }

                            // w:ins/w:r/w:t
                            if ((ce.Elements().Elements().Count(e => e.Name != W.rPr) != 1) ||
                                !ce.Elements().Elements(W.t).Any())
                                return dontConsolidate;

                            XAttribute dateIns2 = ce.Attribute(W.date);

                            string authorIns2 = (string) ce.Attribute(W.author) ?? string.Empty;
                            string dateInsString2 = dateIns2 != null
                                ? ((DateTime) dateIns2).ToString("s")
                                : string.Empty;

                            string idIns2 = (string)ce.Attribute(W.id);

                            return "Wins2" +
                                   authorIns2 +
                                   dateInsString2 +
                                   idIns2 +
                                   ce.Elements()
                                       .Elements(W.rPr)
                                       .Select(rPr => rPr.ToString(SaveOptions.None))
                                       .StringConcatenate();
                        }

                        if (ce.Name == W.del)
                        {
                            if ((ce.Elements(W.r).Elements().Count(e => e.Name != W.rPr) != 1) ||
                                !ce.Elements().Elements(W.delText).Any())
                                return dontConsolidate;

                            XAttribute dateDel2 = ce.Attribute(W.date);

                            string authorDel2 = (string) ce.Attribute(W.author) ?? string.Empty;
                            string dateDelString2 = dateDel2 != null ? ((DateTime) dateDel2).ToString("s") : string.Empty;

                            return "Wdel" +
                                   authorDel2 +
                                   dateDelString2 +
                                   ce.Elements(W.r)
                                       .Elements(W.rPr)
                                       .Select(rPr => rPr.ToString(SaveOptions.None))
                                       .StringConcatenate();
                        }

                        return dontConsolidate;
                    });

            var runContainerWithConsolidatedRuns = new XElement(runContainer.Name,
                runContainer.Attributes(),
                groupedAdjacentRunsWithIdenticalFormatting.Select(g =>
                {
                    if (g.Key == dontConsolidate)
                        return (object) g;

                    string textValue = g
                        .Select(r =>
                            r.Descendants()
                                .Where(d => (d.Name == W.t) || (d.Name == W.delText) || (d.Name == W.instrText))
                                .Select(d => d.Value)
                                .StringConcatenate())
                        .StringConcatenate();
                    XAttribute xs = XmlUtil.GetXmlSpaceAttribute(textValue);

                    if (g.First().Name == W.r)
                    {
                        if (g.First().Element(W.t) != null)
                        {
                            IEnumerable<IEnumerable<XAttribute>> statusAtt =
                                g.Select(r => r.Descendants(W.t).Take(1).Attributes(PtOpenXml.Status));
                            return new XElement(W.r,
                                g.First().Elements(W.rPr),
                                new XElement(W.t, statusAtt, xs, textValue));
                        }

                        if (g.First().Element(W.instrText) != null)
                            return new XElement(W.r,
                                g.First().Elements(W.rPr),
                                new XElement(W.instrText, xs, textValue));
                    }

                    if (g.First().Name == W.ins)
                    {
#if false
                        if (g.First().Elements(W.del).Any())
                            return new XElement(W.ins,
                                g.First().Attributes(),
                                new XElement(W.del,
                                    g.First().Elements(W.del).Attributes(),
                                    new XElement(W.r,
                                        g.First().Elements(W.del).Elements(W.r).Elements(W.rPr),
                                        new XElement(W.delText, xs, textValue))));
#endif
                        return new XElement(W.ins,
                            g.First().Attributes(),
                            new XElement(W.r,
                                g.First().Elements(W.r).Elements(W.rPr),
                                new XElement(W.t, xs, textValue)));
                    }

                    if (g.First().Name == W.del)
                        return new XElement(W.del,
                            g.First().Attributes(),
                            new XElement(W.r,
                                g.First().Elements(W.r).Elements(W.rPr),
                                new XElement(W.delText, xs, textValue)));

                    return g;
                }));

            // Process w:txbxContent//w:p
            foreach (XElement txbx in runContainerWithConsolidatedRuns.Descendants(W.txbxContent))
                foreach (XElement txbxPara in txbx.DescendantsTrimmed(W.txbxContent).Where(d => d.Name == W.p))
                {
                    XElement newPara = CoalesceAdjacentRunsWithIdenticalFormatting(txbxPara);
                    txbxPara.ReplaceWith(newPara);
                }

            // Process additional run containers.
            List<XElement> runContainers = runContainerWithConsolidatedRuns
                .Descendants()
                .Where(d => AdditionalRunContainerNames.Contains(d.Name))
                .ToList();
            foreach (XElement container in runContainers)
            {
                XElement newContainer = CoalesceAdjacentRunsWithIdenticalFormatting(container);
                container.ReplaceWith(newContainer);
            }

            return runContainerWithConsolidatedRuns;
        }

        private static Dictionary<XName, int> Order_settings = new Dictionary<XName, int>
        {
            { W.writeProtection, 10}, 
            { W.view, 20}, 
            { W.zoom, 30}, 
            { W.removePersonalInformation, 40}, 
            { W.removeDateAndTime, 50}, 
            { W.doNotDisplayPageBoundaries, 60}, 
            { W.displayBackgroundShape, 70}, 
            { W.printPostScriptOverText, 80}, 
            { W.printFractionalCharacterWidth, 90}, 
            { W.printFormsData, 100}, 
            { W.embedTrueTypeFonts, 110}, 
            { W.embedSystemFonts, 120}, 
            { W.saveSubsetFonts, 130}, 
            { W.saveFormsData, 140}, 
            { W.mirrorMargins, 150}, 
            { W.alignBordersAndEdges, 160}, 
            { W.bordersDoNotSurroundHeader, 170}, 
            { W.bordersDoNotSurroundFooter, 180}, 
            { W.gutterAtTop, 190}, 
            { W.hideSpellingErrors, 200}, 
            { W.hideGrammaticalErrors, 210}, 
            { W.activeWritingStyle, 220}, 
            { W.proofState, 230}, 
            { W.formsDesign, 240}, 
            { W.attachedTemplate, 250}, 
            { W.linkStyles, 260}, 
            { W.stylePaneFormatFilter, 270}, 
            { W.stylePaneSortMethod, 280}, 
            { W.documentType, 290}, 
            { W.mailMerge, 300}, 
            { W.revisionView, 310}, 
            { W.trackRevisions, 320}, 
            { W.doNotTrackMoves, 330}, 
            { W.doNotTrackFormatting, 340}, 
            { W.documentProtection, 350}, 
            { W.autoFormatOverride, 360}, 
            { W.styleLockTheme, 370}, 
            { W.styleLockQFSet, 380}, 
            { W.defaultTabStop, 390}, 
            { W.autoHyphenation, 400}, 
            { W.consecutiveHyphenLimit, 410}, 
            { W.hyphenationZone, 420}, 
            { W.doNotHyphenateCaps, 430}, 
            { W.showEnvelope, 440}, 
            { W.summaryLength, 450}, 
            { W.clickAndTypeStyle, 460}, 
            { W.defaultTableStyle, 470}, 
            { W.evenAndOddHeaders, 480}, 
            { W.bookFoldRevPrinting, 490}, 
            { W.bookFoldPrinting, 500}, 
            { W.bookFoldPrintingSheets, 510}, 
            { W.drawingGridHorizontalSpacing, 520}, 
            { W.drawingGridVerticalSpacing, 530}, 
            { W.displayHorizontalDrawingGridEvery, 540}, 
            { W.displayVerticalDrawingGridEvery, 550}, 
            { W.doNotUseMarginsForDrawingGridOrigin, 560}, 
            { W.drawingGridHorizontalOrigin, 570}, 
            { W.drawingGridVerticalOrigin, 580}, 
            { W.doNotShadeFormData, 590}, 
            { W.noPunctuationKerning, 600}, 
            { W.characterSpacingControl, 610}, 
            { W.printTwoOnOne, 620}, 
            { W.strictFirstAndLastChars, 630}, 
            { W.noLineBreaksAfter, 640}, 
            { W.noLineBreaksBefore, 650}, 
            { W.savePreviewPicture, 660}, 
            { W.doNotValidateAgainstSchema, 670}, 
            { W.saveInvalidXml, 680}, 
            { W.ignoreMixedContent, 690}, 
            { W.alwaysShowPlaceholderText, 700}, 
            { W.doNotDemarcateInvalidXml, 710}, 
            { W.saveXmlDataOnly, 720}, 
            { W.useXSLTWhenSaving, 730}, 
            { W.saveThroughXslt, 740}, 
            { W.showXMLTags, 750}, 
            { W.alwaysMergeEmptyNamespace, 760}, 
            { W.updateFields, 770}, 
            { W.footnotePr, 780}, 
            { W.endnotePr, 790}, 
            { W.compat, 800}, 
            { W.docVars, 810}, 
            { W.rsids, 820}, 
            { M.mathPr, 830}, 
            { W.attachedSchema, 840}, 
            { W.themeFontLang, 850}, 
            { W.clrSchemeMapping, 860}, 
            { W.doNotIncludeSubdocsInStats, 870}, 
            { W.doNotAutoCompressPictures, 880}, 
            { W.forceUpgrade, 890}, 
            //{W.captions, 900}, 
            { W.readModeInkLockDown, 910}, 
            { W.smartTagType, 920}, 
            //{W.sl:schemaLibrary, 930}, 
            { W.doNotEmbedSmartTags, 940}, 
            { W.decimalSymbol, 950}, 
            { W.listSeparator, 960}, 
        };

#if false
// from the schema in the standard
        
writeProtection
view
zoom
removePersonalInformation
removeDateAndTime
doNotDisplayPageBoundaries
displayBackgroundShape
printPostScriptOverText
printFractionalCharacterWidth
printFormsData
embedTrueTypeFonts
embedSystemFonts
saveSubsetFonts
saveFormsData
mirrorMargins
alignBordersAndEdges
bordersDoNotSurroundHeader
bordersDoNotSurroundFooter
gutterAtTop
hideSpellingErrors
hideGrammaticalErrors
activeWritingStyle
proofState
formsDesign
attachedTemplate
linkStyles
stylePaneFormatFilter
stylePaneSortMethod
documentType
mailMerge
revisionView
trackRevisions
doNotTrackMoves
doNotTrackFormatting
documentProtection
autoFormatOverride
styleLockTheme
styleLockQFSet
defaultTabStop
autoHyphenation
consecutiveHyphenLimit
hyphenationZone
doNotHyphenateCaps
showEnvelope
summaryLength
clickAndTypeStyle
defaultTableStyle
evenAndOddHeaders
bookFoldRevPrinting
bookFoldPrinting
bookFoldPrintingSheets
drawingGridHorizontalSpacing
drawingGridVerticalSpacing
displayHorizontalDrawingGridEvery
displayVerticalDrawingGridEvery
doNotUseMarginsForDrawingGridOrigin
drawingGridHorizontalOrigin
drawingGridVerticalOrigin
doNotShadeFormData
noPunctuationKerning
characterSpacingControl
printTwoOnOne
strictFirstAndLastChars
noLineBreaksAfter
noLineBreaksBefore
savePreviewPicture
doNotValidateAgainstSchema
saveInvalidXml
ignoreMixedContent
alwaysShowPlaceholderText
doNotDemarcateInvalidXml
saveXmlDataOnly
useXSLTWhenSaving
saveThroughXslt
showXMLTags
alwaysMergeEmptyNamespace
updateFields
footnotePr
endnotePr
compat
docVars
rsids
m:mathPr
attachedSchema
themeFontLang
clrSchemeMapping
doNotIncludeSubdocsInStats
doNotAutoCompressPictures
forceUpgrade
captions
readModeInkLockDown
smartTagType
sl:schemaLibrary
doNotEmbedSmartTags
decimalSymbol
listSeparator
#endif

        private static Dictionary<XName, int> Order_pPr = new Dictionary<XName, int>
        {
            { W.pStyle, 10 },
            { W.keepNext, 20 },
            { W.keepLines, 30 },
            { W.pageBreakBefore, 40 },
            { W.framePr, 50 },
            { W.widowControl, 60 },
            { W.numPr, 70 },
            { W.suppressLineNumbers, 80 },
            { W.pBdr, 90 },
            { W.shd, 100 },
            { W.tabs, 120 },
            { W.suppressAutoHyphens, 130 },
            { W.kinsoku, 140 },
            { W.wordWrap, 150 },
            { W.overflowPunct, 160 },
            { W.topLinePunct, 170 },
            { W.autoSpaceDE, 180 },
            { W.autoSpaceDN, 190 },
            { W.bidi, 200 },
            { W.adjustRightInd, 210 },
            { W.snapToGrid, 220 },
            { W.spacing, 230 },
            { W.ind, 240 },
            { W.contextualSpacing, 250 },
            { W.mirrorIndents, 260 },
            { W.suppressOverlap, 270 },
            { W.jc, 280 },
            { W.textDirection, 290 },
            { W.textAlignment, 300 },
            { W.textboxTightWrap, 310 },
            { W.outlineLvl, 320 },
            { W.divId, 330 },
            { W.cnfStyle, 340 },
            { W.rPr, 350 },
            { W.sectPr, 360 },
            { W.pPrChange, 370 },
        };

        private static Dictionary<XName, int> Order_rPr = new Dictionary<XName, int>
        {
            { W.ins, 10 },
            { W.del, 20 },
            { W.rStyle, 30 },
            { W.rFonts, 40 },
            { W.b, 50 },
            { W.bCs, 60 },
            { W.i, 70 },
            { W.iCs, 80 },
            { W.caps, 90 },
            { W.smallCaps, 100 },
            { W.strike, 110 },
            { W.dstrike, 120 },
            { W.outline, 130 },
            { W.shadow, 140 },
            { W.emboss, 150 },
            { W.imprint, 160 },
            { W.noProof, 170 },
            { W.snapToGrid, 180 },
            { W.vanish, 190 },
            { W.webHidden, 200 },
            { W.color, 210 },
            { W.spacing, 220 },
            { W._w, 230 },
            { W.kern, 240 },
            { W.position, 250 },
            { W.sz, 260 },
            { W14.wShadow, 270 },
            { W14.wTextOutline, 280 },
            { W14.wTextFill, 290 },
            { W14.wScene3d, 300 },
            { W14.wProps3d, 310 },
            { W.szCs, 320 },
            { W.highlight, 330 },
            { W.u, 340 },
            { W.effect, 350 },
            { W.bdr, 360 },
            { W.shd, 370 },
            { W.fitText, 380 },
            { W.vertAlign, 390 },
            { W.rtl, 400 },
            { W.cs, 410 },
            { W.em, 420 },
            { W.lang, 430 },
            { W.eastAsianLayout, 440 },
            { W.specVanish, 450 },
            { W.oMath, 460 },
        };

        private static Dictionary<XName, int> Order_tblPr = new Dictionary<XName, int>
        {
            { W.tblStyle, 10 },
            { W.tblpPr, 20 },
            { W.tblOverlap, 30 },
            { W.bidiVisual, 40 },
            { W.tblStyleRowBandSize, 50 },
            { W.tblStyleColBandSize, 60 },
            { W.tblW, 70 },
            { W.jc, 80 },
            { W.tblCellSpacing, 90 },
            { W.tblInd, 100 },
            { W.tblBorders, 110 },
            { W.shd, 120 },
            { W.tblLayout, 130 },
            { W.tblCellMar, 140 },
            { W.tblLook, 150 },
            { W.tblCaption, 160 },
            { W.tblDescription, 170 },
        };

        private static Dictionary<XName, int> Order_tblBorders = new Dictionary<XName, int>
        {
            { W.top, 10 },
            { W.left, 20 },
            { W.start, 30 },
            { W.bottom, 40 },
            { W.right, 50 },
            { W.end, 60 },
            { W.insideH, 70 },
            { W.insideV, 80 },
        };

        private static Dictionary<XName, int> Order_tcPr = new Dictionary<XName, int>
        {
            { W.cnfStyle, 10 },
            { W.tcW, 20 },
            { W.gridSpan, 30 },
            { W.hMerge, 40 },
            { W.vMerge, 50 },
            { W.tcBorders, 60 },
            { W.shd, 70 },
            { W.noWrap, 80 },
            { W.tcMar, 90 },
            { W.textDirection, 100 },
            { W.tcFitText, 110 },
            { W.vAlign, 120 },
            { W.hideMark, 130 },
            { W.headers, 140 },
        };

        private static Dictionary<XName, int> Order_tcBorders = new Dictionary<XName, int>
        {
            { W.top, 10 },
            { W.start, 20 },
            { W.left, 30 },
            { W.bottom, 40 },
            { W.right, 50 },
            { W.end, 60 },
            { W.insideH, 70 },
            { W.insideV, 80 },
            { W.tl2br, 90 },
            { W.tr2bl, 100 },
        };

        private static Dictionary<XName, int> Order_pBdr = new Dictionary<XName, int>
        {
            { W.top, 10 },
            { W.left, 20 },
            { W.bottom, 30 },
            { W.right, 40 },
            { W.between, 50 },
            { W.bar, 60 },
        };

        public static object WmlOrderElementsPerStandard(XNode node)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Name == W.pPr)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                        {
                            if (Order_pPr.ContainsKey(e.Name))
                                return Order_pPr[e.Name];
                            return 999;
                        }));

                if (element.Name == W.rPr)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                        {
                            if (Order_rPr.ContainsKey(e.Name))
                                return Order_rPr[e.Name];
                            return 999;
                        }));

                if (element.Name == W.tblPr)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                        {
                            if (Order_tblPr.ContainsKey(e.Name))
                                return Order_tblPr[e.Name];
                            return 999;
                        }));

                if (element.Name == W.tcPr)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                        {
                            if (Order_tcPr.ContainsKey(e.Name))
                                return Order_tcPr[e.Name];
                            return 999;
                        }));

                if (element.Name == W.tcBorders)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                        {
                            if (Order_tcBorders.ContainsKey(e.Name))
                                return Order_tcBorders[e.Name];
                            return 999;
                        }));

                if (element.Name == W.tblBorders)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                        {
                            if (Order_tblBorders.ContainsKey(e.Name))
                                return Order_tblBorders[e.Name];
                            return 999;
                        }));

                if (element.Name == W.pBdr)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                        {
                            if (Order_pBdr.ContainsKey(e.Name))
                                return Order_pBdr[e.Name];
                            return 999;
                        }));

                if (element.Name == W.p)
                {
                    var newP = new XElement(element.Name,
                        element.Attributes(),
                        element.Elements(W.pPr).Select(e => (XElement)WmlOrderElementsPerStandard(e)),
                        element.Elements().Where(e => e.Name != W.pPr).Select(e => (XElement)WmlOrderElementsPerStandard(e)));
                    return newP;
                }

                if (element.Name == W.r)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements(W.rPr).Select(e => (XElement)WmlOrderElementsPerStandard(e)),
                        element.Elements().Where(e => e.Name != W.rPr).Select(e => (XElement)WmlOrderElementsPerStandard(e)));

                if (element.Name == W.settings)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                        {
                            if (Order_settings.ContainsKey(e.Name))
                                return Order_settings[e.Name];
                            return 999;
                        }));

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => WmlOrderElementsPerStandard(n)));
            }
            return node;
        }

        public static WmlDocument BreakLinkToTemplate(WmlDocument source)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(source.DocumentByteArray, 0, source.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    var efpp = wDoc.ExtendedFilePropertiesPart;
                    if (efpp != null)
                    {
                        var xd = efpp.GetXDocument();
                        var template = xd.Descendants(EP.Template).FirstOrDefault();
                        if (template != null)
                            template.Value = "";
                        efpp.PutXDocument();
                    }
                }
                var result = new WmlDocument(source.FileName, ms.ToArray());
                return result;
            }
        }
    }

    public static class PresentationMLUtil
    {
        public static void FixUpPresentationDocument(PresentationDocument pDoc)
        {
            foreach (var part in pDoc.GetAllParts())
            {
                if (part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.slide+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.theme+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.drawingml.chart+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml" ||
                    part.ContentType == "application/vnd.ms-office.drawingml.diagramDrawing+xml")
                {
                    XDocument xd = part.GetXDocument();
                    xd.Descendants().Attributes("smtClean").Remove();
                    xd.Descendants().Attributes("smtId").Remove();
                    part.PutXDocument();
                }
                if (part.ContentType == "application/vnd.openxmlformats-officedocument.vmlDrawing")
                {
                    string fixedContent = null;

                    using (var stream = part.GetStream(FileMode.Open, FileAccess.ReadWrite))
                    {
                        using (var sr = new StreamReader(stream))
                        {
                            //string input = @"    <![if gte mso 9]><v:imagedata o:relid=""rId15""";
                            var input = sr.ReadToEnd();
                            string pattern = @"<!\[(?<test>.*)\]>";
                            //string replacement = "<![CDATA[${test}]]>";
                            //fixedContent = Regex.Replace(input, pattern, replacement, RegexOptions.Multiline);
                            fixedContent = Regex.Replace(input, pattern, m =>
                            {
                                var g = m.Groups[1].Value;
                                if (g.StartsWith("CDATA["))
                                    return "<![" + g + "]>";
                                else
                                    return "<![CDATA[" + g + "]]>";
                            },
                            RegexOptions.Multiline);

                            //var input = @"xxxxx o:relid=""rId1"" o:relid=""rId1"" xxxxx";
                            pattern = @"o:relid=[""'](?<id1>.*)[""'] o:relid=[""'](?<id2>.*)[""']";
                            fixedContent = Regex.Replace(fixedContent, pattern, m =>
                            {
                                var g = m.Groups[1].Value;
                                return @"o:relid=""" + g + @"""";
                            },
                            RegexOptions.Multiline);

                            fixedContent = fixedContent.Replace("</xml>ml>", "</xml>");

                            stream.SetLength(fixedContent.Length);
                        }
                    }
                    using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(fixedContent)))
                        part.FeedData(ms);
                }
            }
        }
    }

    public static class SpreadsheetMLUtil
    {
        public static string GetCellType(string value)
        {
            if (value.Any(c => !Char.IsDigit(c) && c != '.'))
                return "str";
            return null;
        }

        public static string IntToColumnId(int i)
        {
            if (i >= 0 && i <= 25)
                return ((char)(((int)'A') + i)).ToString();
            if (i >= 26 && i <= 701)
            {
                int v = i - 26;
                int h = v / 26;
                int l = v % 26;
                return ((char)(((int)'A') + h)).ToString() + ((char)(((int)'A') + l)).ToString();
            }
            // 17576
            if (i >= 702 && i <= 18277)
            {
                int v = i - 702;
                int h = v / 676;
                int r = v % 676;
                int m = r / 26;
                int l = r % 26;
                return ((char)(((int)'A') + h)).ToString() +
                    ((char)(((int)'A') + m)).ToString() +
                    ((char)(((int)'A') + l)).ToString();
            }
            throw new ColumnReferenceOutOfRange(i.ToString());
        }

        private static int CharToInt(char c)
        {
            return (int)c - (int)'A';
        }

        public static int ColumnIdToInt(string cid)
        {
            if (cid.Length == 1)
                return CharToInt(cid[0]);
            if (cid.Length == 2)
            {
                return CharToInt(cid[0]) * 26 + CharToInt(cid[1]) + 26;
            }
            if (cid.Length == 3)
            {

                return CharToInt(cid[0]) * 676 + CharToInt(cid[1]) * 26 + CharToInt(cid[2]) + 702;
            }
            throw new ColumnReferenceOutOfRange(cid);
        }

        public static IEnumerable<string> ColumnIDs()
        {
            for (var c = (int)'A'; c <= (int)'Z'; ++c)
                yield return ((char)c).ToString();
            for (var c1 = (int)'A'; c1 <= (int)'Z'; ++c1)
                for (var c2 = (int)'A'; c2 <= (int)'Z'; ++c2)
                    yield return ((char)c1).ToString() + ((char)c2).ToString();
            for (var d1 = (int)'A'; d1 <= (int)'Z'; ++d1)
                for (var d2 = (int)'A'; d2 <= (int)'Z'; ++d2)
                    for (var d3 = (int)'A'; d3 <= (int)'Z'; ++d3)
                        yield return ((char)d1).ToString() + ((char)d2).ToString() + ((char)d3).ToString();
        }

        public static string ColumnIdOf(string cellReference)
        {
            string columnIdOf = cellReference.Split('0', '1', '2', '3', '4', '5', '6', '7', '8', '9').First();
            return columnIdOf;
        }
    }

    public class Util
    {
        public static string[] WordprocessingExtensions = new[] {
            ".docx",
            ".docm",
            ".dotx",
            ".dotm",
        };

        public static bool IsWordprocessingML(string ext)
        {
            return WordprocessingExtensions.Contains(ext.ToLower());
        }

        public static string[] SpreadsheetExtensions = new[] {
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

        public static string[] PresentationExtensions = new[] {
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
            XElement propAtt = rPr.Element(propertyName);
            if (propAtt == null)
                return null;

            XAttribute val = propAtt.Attribute(W.val);
            if (val == null)
                return true;

            string s = ((string) val).ToLower();
            if (s == "1")
                return true;
            if (s == "0")
                return false;
            if (s == "true")
                return true;
            if (s == "false")
                return false;
            if (s == "on")
                return true;
            if (s == "off")
                return false;

            return (bool) propAtt.Attribute(W.val);
        }
    }

    public class FieldInfo
    {
        public string FieldType;
        public string[] Switches;
        public string[] Arguments;
    }

    public static class FieldParser
    {
        private enum State
        {
            InToken,
            InWhiteSpace,
            InQuotedToken,
            OnOpeningQuote,
            OnClosingQuote,
            OnBackslash,
        }

        private static string[] GetTokens(string field)
        {
            State state = State.InWhiteSpace;
            int tokenStart = 0;
            char quoteStart = char.MinValue;
            List<string> tokens = new List<string>();
            for (int c = 0; c < field.Length; c++)
            {
                if (Char.IsWhiteSpace(field[c]))
                {
                    if (state == State.InToken)
                    {
                        tokens.Add(field.Substring(tokenStart, c - tokenStart));
                        state = State.InWhiteSpace;
                        continue;
                    }
                    if (state == State.OnOpeningQuote)
                    {
                        tokenStart = c;
                        state = State.InQuotedToken;
                    }
                    if (state == State.OnClosingQuote)
                        state = State.InWhiteSpace;
                    continue;
                }
                if (field[c] == '\\')
                {
                    if (state == State.InQuotedToken)
                    {
                        state = State.OnBackslash;
                        continue;
                    }
                }
                if (state == State.OnBackslash)
                {
                    state = State.InQuotedToken;
                    continue;
                }
                if (field[c] == '"' || field[c] == '\'' || field[c] == 0x201d)
                {
                    if (state == State.InWhiteSpace)
                    {
                        quoteStart = field[c];
                        state = State.OnOpeningQuote;
                        continue;
                    }
                    if (state == State.InQuotedToken)
                    {
                        if (field[c] == quoteStart)
                        {
                            tokens.Add(field.Substring(tokenStart, c - tokenStart));
                            state = State.OnClosingQuote;
                            continue;
                        }
                        continue;
                    }
                    if (state == State.OnOpeningQuote)
                    {
                        if (field[c] == quoteStart)
                        {
                            state = State.OnClosingQuote;
                            continue;
                        }
                        else
                        {
                            tokenStart = c;
                            state = State.InQuotedToken;
                            continue;
                        }
                    }
                    continue;
                }
                if (state == State.InWhiteSpace)
                {
                    tokenStart = c;
                    state = State.InToken;
                    continue;
                }
                if (state == State.OnOpeningQuote)
                {
                    tokenStart = c;
                    state = State.InQuotedToken;
                    continue;
                }
                if (state == State.OnClosingQuote)
                {
                    tokenStart = c;
                    state = State.InToken;
                    continue;
                }
            }
            if (state == State.InToken)
                tokens.Add(field.Substring(tokenStart, field.Length - tokenStart));
            return tokens.ToArray();
        }

        public static FieldInfo ParseField(string field)
        {
            FieldInfo emptyField = new FieldInfo
            {
                FieldType = "",
                Arguments = new string[] { },
                Switches = new string[] { },
            };

            if (field.Length == 0)
                return emptyField;
            string fieldType = field.TrimStart().Split(' ').FirstOrDefault();
            if (fieldType == null || fieldType.ToUpper() != "HYPERLINK" || fieldType.ToUpper() != "REF")
                return emptyField;
            string[] tokens = GetTokens(field);
            if (tokens.Length == 0)
                return emptyField;
            FieldInfo fieldInfo = new FieldInfo()
            {
                FieldType = tokens[0],
                Switches = tokens.Where(t => t[0] == '\\').ToArray(),
                Arguments = tokens.Skip(1).Where(t => t[0] != '\\').ToArray(),
            };
            return fieldInfo;
        }
    }

    class ContentPartRelTypeIdTuple
    {
        public OpenXmlPart ContentPart { get; set; }
        public string RelationshipType { get; set; }
        public string RelationshipId { get; set; }
    }

    // This class is used to prevent duplication of images
    class ImageData
    {
        private string ContentType { get; set; }
        private byte[] Image { get; set; }
        public OpenXmlPart ImagePart { get; set; }
        public List<ContentPartRelTypeIdTuple> ContentPartRelTypeIdList = new List<ContentPartRelTypeIdTuple>();

        public ImageData(ImagePart part)
        {
            ContentType = part.ContentType;
            using (Stream s = part.GetStream(FileMode.Open, FileAccess.Read))
            {
                Image = new byte[s.Length];
                s.Read(Image, 0, (int)s.Length);
            }
        }

        public void AddContentPartRelTypeResourceIdTupple(OpenXmlPart contentPart, string relationshipType, string relationshipId)
        {
            ContentPartRelTypeIdList.Add(
                new ContentPartRelTypeIdTuple()
                {
                    ContentPart = contentPart,
                    RelationshipType = relationshipType,
                    RelationshipId = relationshipId,
                });
        }

        public void WriteImage(ImagePart part)
        {
            using (Stream s = part.GetStream(FileMode.Create, FileAccess.ReadWrite))
                s.Write(Image, 0, Image.GetUpperBound(0) + 1);
        }

        public bool Compare(ImageData arg)
        {
            if (ContentType != arg.ContentType)
                return false;
            if (Image.GetLongLength(0) != arg.Image.GetLongLength(0))
                return false;
            // Compare the arrays byte by byte
            long length = Image.GetLongLength(0);
            byte[] image1 = Image;
            byte[] image2 = arg.Image;
            for (long n = 0; n < length; n++)
                if (image1[n] != image2[n])
                    return false;
            return true;
        }
    }

    // This class is used to prevent duplication of media
    class MediaData
    {
        private string ContentType { get; set; }
        private byte[] Media { get; set; }
        public DataPart DataPart { get; set; }
        public List<ContentPartRelTypeIdTuple> ContentPartRelTypeIdList = new List<ContentPartRelTypeIdTuple>();

        public MediaData(DataPart part)
        {
            ContentType = part.ContentType;
            using (Stream s = part.GetStream(FileMode.Open, FileAccess.Read))
            {
                Media = new byte[s.Length];
                s.Read(Media, 0, (int)s.Length);
            }
        }

        public void AddContentPartRelTypeResourceIdTupple(OpenXmlPart contentPart, string relationshipType, string relationshipId)
        {
            ContentPartRelTypeIdList.Add(
                new ContentPartRelTypeIdTuple()
                {
                    ContentPart = contentPart,
                    RelationshipType = relationshipType,
                    RelationshipId = relationshipId,
                });
        }

        public void WriteMedia(DataPart part)
        {
            using (Stream s = part.GetStream(FileMode.Create, FileAccess.ReadWrite))
                s.Write(Media, 0, Media.GetUpperBound(0) + 1);
        }

        public bool Compare(MediaData arg)
        {
            if (ContentType != arg.ContentType)
                return false;
            if (Media.GetLongLength(0) != arg.Media.GetLongLength(0))
                return false;
            // Compare the arrays byte by byte
            long length = Media.GetLongLength(0);
            byte[] media1 = Media;
            byte[] media2 = arg.Media;
            for (long n = 0; n < length; n++)
                if (media1[n] != media2[n])
                    return false;
            return true;
        }
    }

#if !NET35
    public static class UriFixer
    {
        public static void FixInvalidUri(Stream fs, Func<string, Uri> invalidUriHandler)
        {
            XNamespace relNs = "http://schemas.openxmlformats.org/package/2006/relationships";
            using (ZipArchive za = new ZipArchive(fs, ZipArchiveMode.Update))
            {
                foreach (var entry in za.Entries.ToList())
                {
                    if (!entry.Name.EndsWith(".rels"))
                        continue;
                    bool replaceEntry = false;
                    XDocument entryXDoc = null;
                    using (var entryStream = entry.Open())
                    {
                        try
                        {
                            entryXDoc = XDocument.Load(entryStream);
                            if (entryXDoc.Root != null && entryXDoc.Root.Name.Namespace == relNs)
                            {
                                var urisToCheck = entryXDoc
                                    .Descendants(relNs + "Relationship")
                                    .Where(r => r.Attribute("TargetMode") != null && (string)r.Attribute("TargetMode") == "External");
                                foreach (var rel in urisToCheck)
                                {
                                    var target = (string)rel.Attribute("Target");
                                    if (target != null)
                                    {
                                        try
                                        {
                                            Uri uri = new Uri(target);
                                        }
                                        catch (UriFormatException)
                                        {
                                            Uri newUri = invalidUriHandler(target);
                                            rel.Attribute("Target").Value = newUri.ToString();
                                            replaceEntry = true;
                                        }
                                    }
                                }
                            }
                        }
                        catch (XmlException)
                        {
                            continue;
                        }
                    }
                    if (replaceEntry)
                    {
                        var fullName = entry.FullName;
                        entry.Delete();
                        var newEntry = za.CreateEntry(fullName);
                        using (StreamWriter writer = new StreamWriter(newEntry.Open()))
                        using (XmlWriter xmlWriter = XmlWriter.Create(writer))
                        {
                            entryXDoc.WriteTo(xmlWriter);
                        }
                    }
                }
            }
        }
    }
#endif

    public static class ACTIVEX
    {
        public static readonly XNamespace activex =
            "http://schemas.microsoft.com/office/2006/activeX";
        public static readonly XName classid = activex + "classid";
        public static readonly XName font = activex + "font";
        public static readonly XName license = activex + "license";
        public static readonly XName name = activex + "name";
        public static readonly XName ocx = activex + "ocx";
        public static readonly XName ocxPr = activex + "ocxPr";
        public static readonly XName persistence = activex + "persistence";
        public static readonly XName value = activex + "value";
    }

    public static class BIBLIO
    {
        public static readonly XNamespace biblio =
            "http://schemas.microsoft.com/office/word/2004/10/bibliography";
        public static readonly XName AlbumTitle = biblio + "AlbumTitle";
        public static readonly XName Artist = biblio + "Artist";
        public static readonly XName Author = biblio + "Author";
        public static readonly XName City = biblio + "City";
        public static readonly XName Comments = biblio + "Comments";
        public static readonly XName Composer = biblio + "Composer";
        public static readonly XName Conductor = biblio + "Conductor";
        public static readonly XName ConferenceName = biblio + "ConferenceName";
        public static readonly XName Country = biblio + "Country";
        public static readonly XName Day = biblio + "Day";
        public static readonly XName DayAccessed = biblio + "DayAccessed";
        public static readonly XName Editor = biblio + "Editor";
        public static readonly XName First = biblio + "First";
        public static readonly XName Guid = biblio + "Guid";
        public static readonly XName InternetSiteTitle = biblio + "InternetSiteTitle";
        public static readonly XName Inventor = biblio + "Inventor";
        public static readonly XName Last = biblio + "Last";
        public static readonly XName LCID = biblio + "LCID";
        public static readonly XName Main = biblio + "Main";
        public static readonly XName Medium = biblio + "Medium";
        public static readonly XName Middle = biblio + "Middle";
        public static readonly XName Month = biblio + "Month";
        public static readonly XName MonthAccessed = biblio + "MonthAccessed";
        public static readonly XName NameList = biblio + "NameList";
        public static readonly XName NumberVolumes = biblio + "NumberVolumes";
        public static readonly XName Pages = biblio + "Pages";
        public static readonly XName PatentNumber = biblio + "PatentNumber";
        public static readonly XName Performer = biblio + "Performer";
        public static readonly XName Person = biblio + "Person";
        public static readonly XName ProducerName = biblio + "ProducerName";
        public static readonly XName ProductionCompany = biblio + "ProductionCompany";
        public static readonly XName Publisher = biblio + "Publisher";
        public static readonly XName RefOrder = biblio + "RefOrder";
        public static readonly XName ShortTitle = biblio + "ShortTitle";
        public static readonly XName Source = biblio + "Source";
        public static readonly XName Sources = biblio + "Sources";
        public static readonly XName SourceType = biblio + "SourceType";
        public static readonly XName Tag = biblio + "Tag";
        public static readonly XName Title = biblio + "Title";
        public static readonly XName Translator = biblio + "Translator";
        public static readonly XName Type = biblio + "Type";
        public static readonly XName URL = biblio + "URL";
        public static readonly XName Version = biblio + "Version";
        public static readonly XName Volume = biblio + "Volume";
        public static readonly XName Year = biblio + "Year";
        public static readonly XName YearAccessed = biblio + "YearAccessed";
    }

    public static class INK
    {
        public static readonly XNamespace ink =
            "http://schemas.microsoft.com/ink/2010/main";
        public static readonly XName context = ink + "context";
        public static readonly XName sourceLink = ink + "sourceLink";
    }

    public static class SSNoNamespace
    {
        public static readonly XName _ref = "ref";
        public static readonly XName applyAlignment = "applyAlignment";
        public static readonly XName applyBorder = "applyBorder";
        public static readonly XName applyFont = "applyFont";
        public static readonly XName applyNumberFormat = "applyNumberFormat";
        public static readonly XName appName = "appName";
        public static readonly XName baseType = "baseType";
        public static readonly XName borderId = "borderId";
        public static readonly XName bottom = "bottom";
        public static readonly XName builtinId = "builtinId";
        public static readonly XName calcId = "calcId";
        public static readonly XName count = "count";
        public static readonly XName customHeight = "customHeight";
        public static readonly XName defaultColWidth = "defaultColWidth";
        public static readonly XName defaultPivotStyle = "defaultPivotStyle";
        public static readonly XName defaultRowHeight = "defaultRowHeight";
        public static readonly XName defaultTableStyle = "defaultTableStyle";
        public static readonly XName defaultThemeVersion = "defaultThemeVersion";
        public static readonly XName displayName = "displayName";
        public static readonly XName fillId = "fillId";
        public static readonly XName fontId = "fontId";
        public static readonly XName footer = "footer";
        public static readonly XName formatCode = "formatCode";
        public static readonly XName header = "header";
        public static readonly XName horizontal = "horizontal";
        public static readonly XName ht = "ht";
        public static readonly XName id = "id";
        public static readonly XName lastEdited = "lastEdited";
        public static readonly XName left = "left";
        public static readonly XName lowestEdited = "lowestEdited";
        public static readonly XName max = "max";
        public static readonly XName min = "min";
        public static readonly XName name = "name";
        public static readonly XName numFmtId = "numFmtId";
        public static readonly XName patternType = "patternType";
        public static readonly XName r = "r";
        public static readonly XName rgb = "rgb";
        public static readonly XName right = "right";
        public static readonly XName rupBuild = "rupBuild";
        public static readonly XName s = "s";
        public static readonly XName sheetId = "sheetId";
        public static readonly XName showColumnStripes = "showColumnStripes";
        public static readonly XName showFirstColumn = "showFirstColumn";
        public static readonly XName showLastColumn = "showLastColumn";
        public static readonly XName showRowStripes = "showRowStripes";
        public static readonly XName size = "size";
        public static readonly XName spans = "spans";
        public static readonly XName sqref = "sqref";
        public static readonly XName style = "style";
        public static readonly XName t = "t";
        public static readonly XName tabSelected = "tabSelected";
        public static readonly XName theme = "theme";
        public static readonly XName thickBot = "thickBot";
        public static readonly XName top = "top";
        public static readonly XName totalsRowShown = "totalsRowShown";
        public static readonly XName uniqueCount = "uniqueCount";
        public static readonly XName val = "val";
        public static readonly XName width = "width";
        public static readonly XName windowHeight = "windowHeight";
        public static readonly XName windowWidth = "windowWidth";
        public static readonly XName workbookViewId = "workbookViewId";
        public static readonly XName xfId = "xfId";
        public static readonly XName xWindow = "xWindow";
        public static readonly XName yWindow = "yWindow";
    }

    public static class WNE
    {
        public static readonly XNamespace wne =
            "http://schemas.microsoft.com/office/word/2006/wordml";
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

    public static class WPC
    {
        public static readonly XNamespace wpc =
            "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas";
    }

    public static class WPG
    {
        public static readonly XNamespace wpg =
            "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
    }

    public static class WPI
    {
        public static readonly XNamespace wpi =
            "http://schemas.microsoft.com/office/word/2010/wordprocessingInk";
    }

    public static class A
    {
        public static readonly XNamespace a =
            "http://schemas.openxmlformats.org/drawingml/2006/main";
        public static readonly XName accent1 = a + "accent1";
        public static readonly XName accent2 = a + "accent2";
        public static readonly XName accent3 = a + "accent3";
        public static readonly XName accent4 = a + "accent4";
        public static readonly XName accent5 = a + "accent5";
        public static readonly XName accent6 = a + "accent6";
        public static readonly XName ahLst = a + "ahLst";
        public static readonly XName ahPolar = a + "ahPolar";
        public static readonly XName ahXY = a + "ahXY";
        public static readonly XName alpha = a + "alpha";
        public static readonly XName alphaBiLevel = a + "alphaBiLevel";
        public static readonly XName alphaCeiling = a + "alphaCeiling";
        public static readonly XName alphaFloor = a + "alphaFloor";
        public static readonly XName alphaInv = a + "alphaInv";
        public static readonly XName alphaMod = a + "alphaMod";
        public static readonly XName alphaModFix = a + "alphaModFix";
        public static readonly XName alphaOff = a + "alphaOff";
        public static readonly XName alphaOutset = a + "alphaOutset";
        public static readonly XName alphaRepl = a + "alphaRepl";
        public static readonly XName anchor = a + "anchor";
        public static readonly XName arcTo = a + "arcTo";
        public static readonly XName audioCd = a + "audioCd";
        public static readonly XName audioFile = a + "audioFile";
        public static readonly XName avLst = a + "avLst";
        public static readonly XName backdrop = a + "backdrop";
        public static readonly XName band1H = a + "band1H";
        public static readonly XName band1V = a + "band1V";
        public static readonly XName band2H = a + "band2H";
        public static readonly XName band2V = a + "band2V";
        public static readonly XName bevel = a + "bevel";
        public static readonly XName bevelB = a + "bevelB";
        public static readonly XName bevelT = a + "bevelT";
        public static readonly XName bgClr = a + "bgClr";
        public static readonly XName bgFillStyleLst = a + "bgFillStyleLst";
        public static readonly XName biLevel = a + "biLevel";
        public static readonly XName bldChart = a + "bldChart";
        public static readonly XName bldDgm = a + "bldDgm";
        public static readonly XName blend = a + "blend";
        public static readonly XName blip = a + "blip";
        public static readonly XName blipFill = a + "blipFill";
        public static readonly XName blue = a + "blue";
        public static readonly XName blueMod = a + "blueMod";
        public static readonly XName blueOff = a + "blueOff";
        public static readonly XName blur = a + "blur";
        public static readonly XName bodyPr = a + "bodyPr";
        public static readonly XName bottom = a + "bottom";
        public static readonly XName br = a + "br";
        public static readonly XName buAutoNum = a + "buAutoNum";
        public static readonly XName buBlip = a + "buBlip";
        public static readonly XName buChar = a + "buChar";
        public static readonly XName buClr = a + "buClr";
        public static readonly XName buClrTx = a + "buClrTx";
        public static readonly XName buFont = a + "buFont";
        public static readonly XName buFontTx = a + "buFontTx";
        public static readonly XName buNone = a + "buNone";
        public static readonly XName buSzPct = a + "buSzPct";
        public static readonly XName buSzPts = a + "buSzPts";
        public static readonly XName buSzTx = a + "buSzTx";
        public static readonly XName camera = a + "camera";
        public static readonly XName cell3D = a + "cell3D";
        public static readonly XName chart = a + "chart";
        public static readonly XName chExt = a + "chExt";
        public static readonly XName chOff = a + "chOff";
        public static readonly XName close = a + "close";
        public static readonly XName clrChange = a + "clrChange";
        public static readonly XName clrFrom = a + "clrFrom";
        public static readonly XName clrMap = a + "clrMap";
        public static readonly XName clrRepl = a + "clrRepl";
        public static readonly XName clrScheme = a + "clrScheme";
        public static readonly XName clrTo = a + "clrTo";
        public static readonly XName cNvCxnSpPr = a + "cNvCxnSpPr";
        public static readonly XName cNvGraphicFramePr = a + "cNvGraphicFramePr";
        public static readonly XName cNvGrpSpPr = a + "cNvGrpSpPr";
        public static readonly XName cNvPicPr = a + "cNvPicPr";
        public static readonly XName cNvPr = a + "cNvPr";
        public static readonly XName cNvSpPr = a + "cNvSpPr";
        public static readonly XName comp = a + "comp";
        public static readonly XName cont = a + "cont";
        public static readonly XName contourClr = a + "contourClr";
        public static readonly XName cs = a + "cs";
        public static readonly XName cubicBezTo = a + "cubicBezTo";
        public static readonly XName custClr = a + "custClr";
        public static readonly XName custClrLst = a + "custClrLst";
        public static readonly XName custDash = a + "custDash";
        public static readonly XName custGeom = a + "custGeom";
        public static readonly XName cxn = a + "cxn";
        public static readonly XName cxnLst = a + "cxnLst";
        public static readonly XName cxnSp = a + "cxnSp";
        public static readonly XName cxnSpLocks = a + "cxnSpLocks";
        public static readonly XName defPPr = a + "defPPr";
        public static readonly XName defRPr = a + "defRPr";
        public static readonly XName dgm = a + "dgm";
        public static readonly XName dk1 = a + "dk1";
        public static readonly XName dk2 = a + "dk2";
        public static readonly XName ds = a + "ds";
        public static readonly XName duotone = a + "duotone";
        public static readonly XName ea = a + "ea";
        public static readonly XName effect = a + "effect";
        public static readonly XName effectDag = a + "effectDag";
        public static readonly XName effectLst = a + "effectLst";
        public static readonly XName effectRef = a + "effectRef";
        public static readonly XName effectStyle = a + "effectStyle";
        public static readonly XName effectStyleLst = a + "effectStyleLst";
        public static readonly XName end = a + "end";
        public static readonly XName endCxn = a + "endCxn";
        public static readonly XName endParaRPr = a + "endParaRPr";
        public static readonly XName ext = a + "ext";
        public static readonly XName extLst = a + "extLst";
        public static readonly XName extraClrScheme = a + "extraClrScheme";
        public static readonly XName extraClrSchemeLst = a + "extraClrSchemeLst";
        public static readonly XName extrusionClr = a + "extrusionClr";
        public static readonly XName fgClr = a + "fgClr";
        public static readonly XName fill = a + "fill";
        public static readonly XName fillOverlay = a + "fillOverlay";
        public static readonly XName fillRect = a + "fillRect";
        public static readonly XName fillRef = a + "fillRef";
        public static readonly XName fillStyleLst = a + "fillStyleLst";
        public static readonly XName fillToRect = a + "fillToRect";
        public static readonly XName firstCol = a + "firstCol";
        public static readonly XName firstRow = a + "firstRow";
        public static readonly XName flatTx = a + "flatTx";
        public static readonly XName fld = a + "fld";
        public static readonly XName fmtScheme = a + "fmtScheme";
        public static readonly XName folHlink = a + "folHlink";
        public static readonly XName font = a + "font";
        public static readonly XName fontRef = a + "fontRef";
        public static readonly XName fontScheme = a + "fontScheme";
        public static readonly XName gamma = a + "gamma";
        public static readonly XName gd = a + "gd";
        public static readonly XName gdLst = a + "gdLst";
        public static readonly XName glow = a + "glow";
        public static readonly XName gradFill = a + "gradFill";
        public static readonly XName graphic = a + "graphic";
        public static readonly XName graphicData = a + "graphicData";
        public static readonly XName graphicFrame = a + "graphicFrame";
        public static readonly XName graphicFrameLocks = a + "graphicFrameLocks";
        public static readonly XName gray = a + "gray";
        public static readonly XName grayscl = a + "grayscl";
        public static readonly XName green = a + "green";
        public static readonly XName greenMod = a + "greenMod";
        public static readonly XName greenOff = a + "greenOff";
        public static readonly XName gridCol = a + "gridCol";
        public static readonly XName grpFill = a + "grpFill";
        public static readonly XName grpSp = a + "grpSp";
        public static readonly XName grpSpLocks = a + "grpSpLocks";
        public static readonly XName grpSpPr = a + "grpSpPr";
        public static readonly XName gs = a + "gs";
        public static readonly XName gsLst = a + "gsLst";
        public static readonly XName headEnd = a + "headEnd";
        public static readonly XName highlight = a + "highlight";
        public static readonly XName hlink = a + "hlink";
        public static readonly XName hlinkClick = a + "hlinkClick";
        public static readonly XName hlinkHover = a + "hlinkHover";
        public static readonly XName hlinkMouseOver = a + "hlinkMouseOver";
        public static readonly XName hsl = a + "hsl";
        public static readonly XName hslClr = a + "hslClr";
        public static readonly XName hue = a + "hue";
        public static readonly XName hueMod = a + "hueMod";
        public static readonly XName hueOff = a + "hueOff";
        public static readonly XName innerShdw = a + "innerShdw";
        public static readonly XName insideH = a + "insideH";
        public static readonly XName insideV = a + "insideV";
        public static readonly XName inv = a + "inv";
        public static readonly XName invGamma = a + "invGamma";
        public static readonly XName lastCol = a + "lastCol";
        public static readonly XName lastRow = a + "lastRow";
        public static readonly XName latin = a + "latin";
        public static readonly XName left = a + "left";
        public static readonly XName lightRig = a + "lightRig";
        public static readonly XName lin = a + "lin";
        public static readonly XName ln = a + "ln";
        public static readonly XName lnB = a + "lnB";
        public static readonly XName lnBlToTr = a + "lnBlToTr";
        public static readonly XName lnDef = a + "lnDef";
        public static readonly XName lnL = a + "lnL";
        public static readonly XName lnR = a + "lnR";
        public static readonly XName lnRef = a + "lnRef";
        public static readonly XName lnSpc = a + "lnSpc";
        public static readonly XName lnStyleLst = a + "lnStyleLst";
        public static readonly XName lnT = a + "lnT";
        public static readonly XName lnTlToBr = a + "lnTlToBr";
        public static readonly XName lnTo = a + "lnTo";
        public static readonly XName lstStyle = a + "lstStyle";
        public static readonly XName lt1 = a + "lt1";
        public static readonly XName lt2 = a + "lt2";
        public static readonly XName lum = a + "lum";
        public static readonly XName lumMod = a + "lumMod";
        public static readonly XName lumOff = a + "lumOff";
        public static readonly XName lvl1pPr = a + "lvl1pPr";
        public static readonly XName lvl2pPr = a + "lvl2pPr";
        public static readonly XName lvl3pPr = a + "lvl3pPr";
        public static readonly XName lvl4pPr = a + "lvl4pPr";
        public static readonly XName lvl5pPr = a + "lvl5pPr";
        public static readonly XName lvl6pPr = a + "lvl6pPr";
        public static readonly XName lvl7pPr = a + "lvl7pPr";
        public static readonly XName lvl8pPr = a + "lvl8pPr";
        public static readonly XName lvl9pPr = a + "lvl9pPr";
        public static readonly XName majorFont = a + "majorFont";
        public static readonly XName masterClrMapping = a + "masterClrMapping";
        public static readonly XName minorFont = a + "minorFont";
        public static readonly XName miter = a + "miter";
        public static readonly XName moveTo = a + "moveTo";
        public static readonly XName neCell = a + "neCell";
        public static readonly XName noAutofit = a + "noAutofit";
        public static readonly XName noFill = a + "noFill";
        public static readonly XName norm = a + "norm";
        public static readonly XName normAutofit = a + "normAutofit";
        public static readonly XName nvCxnSpPr = a + "nvCxnSpPr";
        public static readonly XName nvGraphicFramePr = a + "nvGraphicFramePr";
        public static readonly XName nvGrpSpPr = a + "nvGrpSpPr";
        public static readonly XName nvPicPr = a + "nvPicPr";
        public static readonly XName nvSpPr = a + "nvSpPr";
        public static readonly XName nwCell = a + "nwCell";
        public static readonly XName objectDefaults = a + "objectDefaults";
        public static readonly XName off = a + "off";
        public static readonly XName outerShdw = a + "outerShdw";
        public static readonly XName overrideClrMapping = a + "overrideClrMapping";
        public static readonly XName p = a + "p";
        public static readonly XName path = a + "path";
        public static readonly XName pathLst = a + "pathLst";
        public static readonly XName pattFill = a + "pattFill";
        public static readonly XName pic = a + "pic";
        public static readonly XName picLocks = a + "picLocks";
        public static readonly XName pos = a + "pos";
        public static readonly XName pPr = a + "pPr";
        public static readonly XName prstClr = a + "prstClr";
        public static readonly XName prstDash = a + "prstDash";
        public static readonly XName prstGeom = a + "prstGeom";
        public static readonly XName prstShdw = a + "prstShdw";
        public static readonly XName prstTxWarp = a + "prstTxWarp";
        public static readonly XName pt = a + "pt";
        public static readonly XName quadBezTo = a + "quadBezTo";
        public static readonly XName quickTimeFile = a + "quickTimeFile";
        public static readonly XName r = a + "r";
        public static readonly XName rect = a + "rect";
        public static readonly XName red = a + "red";
        public static readonly XName redMod = a + "redMod";
        public static readonly XName redOff = a + "redOff";
        public static readonly XName reflection = a + "reflection";
        public static readonly XName relIds = a + "relIds";
        public static readonly XName relOff = a + "relOff";
        public static readonly XName right = a + "right";
        public static readonly XName rot = a + "rot";
        public static readonly XName round = a + "round";
        public static readonly XName rPr = a + "rPr";
        public static readonly XName sat = a + "sat";
        public static readonly XName satMod = a + "satMod";
        public static readonly XName satOff = a + "satOff";
        public static readonly XName scene3d = a + "scene3d";
        public static readonly XName schemeClr = a + "schemeClr";
        public static readonly XName scrgbClr = a + "scrgbClr";
        public static readonly XName seCell = a + "seCell";
        public static readonly XName shade = a + "shade";
        public static readonly XName snd = a + "snd";
        public static readonly XName softEdge = a + "softEdge";
        public static readonly XName solidFill = a + "solidFill";
        public static readonly XName sp = a + "sp";
        public static readonly XName sp3d = a + "sp3d";
        public static readonly XName spAutoFit = a + "spAutoFit";
        public static readonly XName spcAft = a + "spcAft";
        public static readonly XName spcBef = a + "spcBef";
        public static readonly XName spcPct = a + "spcPct";
        public static readonly XName spcPts = a + "spcPts";
        public static readonly XName spDef = a + "spDef";
        public static readonly XName spLocks = a + "spLocks";
        public static readonly XName spPr = a + "spPr";
        public static readonly XName srcRect = a + "srcRect";
        public static readonly XName srgbClr = a + "srgbClr";
        public static readonly XName st = a + "st";
        public static readonly XName stCxn = a + "stCxn";
        public static readonly XName stretch = a + "stretch";
        public static readonly XName style = a + "style";
        public static readonly XName swCell = a + "swCell";
        public static readonly XName sx = a + "sx";
        public static readonly XName sy = a + "sy";
        public static readonly XName sym = a + "sym";
        public static readonly XName sysClr = a + "sysClr";
        public static readonly XName t = a + "t";
        public static readonly XName tab = a + "tab";
        public static readonly XName tableStyle = a + "tableStyle";
        public static readonly XName tableStyleId = a + "tableStyleId";
        public static readonly XName tabLst = a + "tabLst";
        public static readonly XName tailEnd = a + "tailEnd";
        public static readonly XName tbl = a + "tbl";
        public static readonly XName tblBg = a + "tblBg";
        public static readonly XName tblGrid = a + "tblGrid";
        public static readonly XName tblPr = a + "tblPr";
        public static readonly XName tblStyle = a + "tblStyle";
        public static readonly XName tblStyleLst = a + "tblStyleLst";
        public static readonly XName tc = a + "tc";
        public static readonly XName tcBdr = a + "tcBdr";
        public static readonly XName tcPr = a + "tcPr";
        public static readonly XName tcStyle = a + "tcStyle";
        public static readonly XName tcTxStyle = a + "tcTxStyle";
        public static readonly XName theme = a + "theme";
        public static readonly XName themeElements = a + "themeElements";
        public static readonly XName themeOverride = a + "themeOverride";
        public static readonly XName tile = a + "tile";
        public static readonly XName tileRect = a + "tileRect";
        public static readonly XName tint = a + "tint";
        public static readonly XName top = a + "top";
        public static readonly XName tr = a + "tr";
        public static readonly XName txBody = a + "txBody";
        public static readonly XName txDef = a + "txDef";
        public static readonly XName txSp = a + "txSp";
        public static readonly XName uFill = a + "uFill";
        public static readonly XName uFillTx = a + "uFillTx";
        public static readonly XName uLn = a + "uLn";
        public static readonly XName uLnTx = a + "uLnTx";
        public static readonly XName up = a + "up";
        public static readonly XName useSpRect = a + "useSpRect";
        public static readonly XName videoFile = a + "videoFile";
        public static readonly XName wavAudioFile = a + "wavAudioFile";
        public static readonly XName wholeTbl = a + "wholeTbl";
        public static readonly XName xfrm = a + "xfrm";
    }

    public static class A14
    {
        public static readonly XNamespace a14 =
            "http://schemas.microsoft.com/office/drawing/2010/main";
        public static readonly XName artisticChalkSketch = a14 + "artisticChalkSketch";
        public static readonly XName artisticGlass = a14 + "artisticGlass";
        public static readonly XName artisticGlowEdges = a14 + "artisticGlowEdges";
        public static readonly XName artisticLightScreen = a14 + "artisticLightScreen";
        public static readonly XName artisticPlasticWrap = a14 + "artisticPlasticWrap";
        public static readonly XName artisticTexturizer = a14 + "artisticTexturizer";
        public static readonly XName backgroundMark = a14 + "backgroundMark";
        public static readonly XName backgroundRemoval = a14 + "backgroundRemoval";
        public static readonly XName brightnessContrast = a14 + "brightnessContrast";
        public static readonly XName cameraTool = a14 + "cameraTool";
        public static readonly XName colorTemperature = a14 + "colorTemperature";
        public static readonly XName compatExt = a14 + "compatExt";
        public static readonly XName cpLocks = a14 + "cpLocks";
        public static readonly XName extLst = a14 + "extLst";
        public static readonly XName foregroundMark = a14 + "foregroundMark";
        public static readonly XName hiddenEffects = a14 + "hiddenEffects";
        public static readonly XName hiddenFill = a14 + "hiddenFill";
        public static readonly XName hiddenLine = a14 + "hiddenLine";
        public static readonly XName hiddenScene3d = a14 + "hiddenScene3d";
        public static readonly XName hiddenSp3d = a14 + "hiddenSp3d";
        public static readonly XName imgEffect = a14 + "imgEffect";
        public static readonly XName imgLayer = a14 + "imgLayer";
        public static readonly XName imgProps = a14 + "imgProps";
        public static readonly XName legacySpreadsheetColorIndex = a14 + "legacySpreadsheetColorIndex";
        public static readonly XName m = a14 + "m";
        public static readonly XName saturation = a14 + "saturation";
        public static readonly XName shadowObscured = a14 + "shadowObscured";
        public static readonly XName sharpenSoften = a14 + "sharpenSoften";
        public static readonly XName useLocalDpi = a14 + "useLocalDpi";
    }

    public static class C
    {
        public static readonly XNamespace c =
            "http://schemas.openxmlformats.org/drawingml/2006/chart";
        public static readonly XName applyToEnd = c + "applyToEnd";
        public static readonly XName applyToFront = c + "applyToFront";
        public static readonly XName applyToSides = c + "applyToSides";
        public static readonly XName area3DChart = c + "area3DChart";
        public static readonly XName areaChart = c + "areaChart";
        public static readonly XName auto = c + "auto";
        public static readonly XName autoTitleDeleted = c + "autoTitleDeleted";
        public static readonly XName autoUpdate = c + "autoUpdate";
        public static readonly XName axId = c + "axId";
        public static readonly XName axPos = c + "axPos";
        public static readonly XName backWall = c + "backWall";
        public static readonly XName backward = c + "backward";
        public static readonly XName bandFmt = c + "bandFmt";
        public static readonly XName bandFmts = c + "bandFmts";
        public static readonly XName bar3DChart = c + "bar3DChart";
        public static readonly XName barChart = c + "barChart";
        public static readonly XName barDir = c + "barDir";
        public static readonly XName baseTimeUnit = c + "baseTimeUnit";
        public static readonly XName bubble3D = c + "bubble3D";
        public static readonly XName bubbleChart = c + "bubbleChart";
        public static readonly XName bubbleScale = c + "bubbleScale";
        public static readonly XName bubbleSize = c + "bubbleSize";
        public static readonly XName builtInUnit = c + "builtInUnit";
        public static readonly XName cat = c + "cat";
        public static readonly XName catAx = c + "catAx";
        public static readonly XName chart = c + "chart";
        public static readonly XName chartObject = c + "chartObject";
        public static readonly XName chartSpace = c + "chartSpace";
        public static readonly XName clrMapOvr = c + "clrMapOvr";
        public static readonly XName crossAx = c + "crossAx";
        public static readonly XName crossBetween = c + "crossBetween";
        public static readonly XName crosses = c + "crosses";
        public static readonly XName crossesAt = c + "crossesAt";
        public static readonly XName custSplit = c + "custSplit";
        public static readonly XName custUnit = c + "custUnit";
        public static readonly XName data = c + "data";
        public static readonly XName date1904 = c + "date1904";
        public static readonly XName dateAx = c + "dateAx";
        public static readonly XName delete = c + "delete";
        public static readonly XName depthPercent = c + "depthPercent";
        public static readonly XName dispBlanksAs = c + "dispBlanksAs";
        public static readonly XName dispEq = c + "dispEq";
        public static readonly XName dispRSqr = c + "dispRSqr";
        public static readonly XName dispUnits = c + "dispUnits";
        public static readonly XName dispUnitsLbl = c + "dispUnitsLbl";
        public static readonly XName dLbl = c + "dLbl";
        public static readonly XName dLblPos = c + "dLblPos";
        public static readonly XName dLbls = c + "dLbls";
        public static readonly XName doughnutChart = c + "doughnutChart";
        public static readonly XName downBars = c + "downBars";
        public static readonly XName dPt = c + "dPt";
        public static readonly XName dropLines = c + "dropLines";
        public static readonly XName dTable = c + "dTable";
        public static readonly XName errBars = c + "errBars";
        public static readonly XName errBarType = c + "errBarType";
        public static readonly XName errDir = c + "errDir";
        public static readonly XName errValType = c + "errValType";
        public static readonly XName explosion = c + "explosion";
        public static readonly XName ext = c + "ext";
        public static readonly XName externalData = c + "externalData";
        public static readonly XName extLst = c + "extLst";
        public static readonly XName f = c + "f";
        public static readonly XName firstSliceAng = c + "firstSliceAng";
        public static readonly XName floor = c + "floor";
        public static readonly XName fmtId = c + "fmtId";
        public static readonly XName formatCode = c + "formatCode";
        public static readonly XName formatting = c + "formatting";
        public static readonly XName forward = c + "forward";
        public static readonly XName gapDepth = c + "gapDepth";
        public static readonly XName gapWidth = c + "gapWidth";
        public static readonly XName grouping = c + "grouping";
        public static readonly XName h = c + "h";
        public static readonly XName headerFooter = c + "headerFooter";
        public static readonly XName hiLowLines = c + "hiLowLines";
        public static readonly XName hMode = c + "hMode";
        public static readonly XName holeSize = c + "holeSize";
        public static readonly XName hPercent = c + "hPercent";
        public static readonly XName idx = c + "idx";
        public static readonly XName intercept = c + "intercept";
        public static readonly XName invertIfNegative = c + "invertIfNegative";
        public static readonly XName lang = c + "lang";
        public static readonly XName layout = c + "layout";
        public static readonly XName layoutTarget = c + "layoutTarget";
        public static readonly XName lblAlgn = c + "lblAlgn";
        public static readonly XName lblOffset = c + "lblOffset";
        public static readonly XName leaderLines = c + "leaderLines";
        public static readonly XName legend = c + "legend";
        public static readonly XName legendEntry = c + "legendEntry";
        public static readonly XName legendPos = c + "legendPos";
        public static readonly XName line3DChart = c + "line3DChart";
        public static readonly XName lineChart = c + "lineChart";
        public static readonly XName logBase = c + "logBase";
        public static readonly XName lvl = c + "lvl";
        public static readonly XName majorGridlines = c + "majorGridlines";
        public static readonly XName majorTickMark = c + "majorTickMark";
        public static readonly XName majorTimeUnit = c + "majorTimeUnit";
        public static readonly XName majorUnit = c + "majorUnit";
        public static readonly XName manualLayout = c + "manualLayout";
        public static readonly XName marker = c + "marker";
        public static readonly XName max = c + "max";
        public static readonly XName min = c + "min";
        public static readonly XName minorGridlines = c + "minorGridlines";
        public static readonly XName minorTickMark = c + "minorTickMark";
        public static readonly XName minorTimeUnit = c + "minorTimeUnit";
        public static readonly XName minorUnit = c + "minorUnit";
        public static readonly XName minus = c + "minus";
        public static readonly XName multiLvlStrCache = c + "multiLvlStrCache";
        public static readonly XName multiLvlStrRef = c + "multiLvlStrRef";
        public static readonly XName name = c + "name";
        public static readonly XName noEndCap = c + "noEndCap";
        public static readonly XName noMultiLvlLbl = c + "noMultiLvlLbl";
        public static readonly XName numCache = c + "numCache";
        public static readonly XName numFmt = c + "numFmt";
        public static readonly XName numLit = c + "numLit";
        public static readonly XName numRef = c + "numRef";
        public static readonly XName oddFooter = c + "oddFooter";
        public static readonly XName oddHeader = c + "oddHeader";
        public static readonly XName ofPieChart = c + "ofPieChart";
        public static readonly XName ofPieType = c + "ofPieType";
        public static readonly XName order = c + "order";
        public static readonly XName orientation = c + "orientation";
        public static readonly XName overlap = c + "overlap";
        public static readonly XName overlay = c + "overlay";
        public static readonly XName pageMargins = c + "pageMargins";
        public static readonly XName pageSetup = c + "pageSetup";
        public static readonly XName period = c + "period";
        public static readonly XName perspective = c + "perspective";
        public static readonly XName pictureFormat = c + "pictureFormat";
        public static readonly XName pictureOptions = c + "pictureOptions";
        public static readonly XName pictureStackUnit = c + "pictureStackUnit";
        public static readonly XName pie3DChart = c + "pie3DChart";
        public static readonly XName pieChart = c + "pieChart";
        public static readonly XName pivotFmt = c + "pivotFmt";
        public static readonly XName pivotFmts = c + "pivotFmts";
        public static readonly XName pivotSource = c + "pivotSource";
        public static readonly XName plotArea = c + "plotArea";
        public static readonly XName plotVisOnly = c + "plotVisOnly";
        public static readonly XName plus = c + "plus";
        public static readonly XName printSettings = c + "printSettings";
        public static readonly XName protection = c + "protection";
        public static readonly XName pt = c + "pt";
        public static readonly XName ptCount = c + "ptCount";
        public static readonly XName radarChart = c + "radarChart";
        public static readonly XName radarStyle = c + "radarStyle";
        public static readonly XName rAngAx = c + "rAngAx";
        public static readonly XName rich = c + "rich";
        public static readonly XName rotX = c + "rotX";
        public static readonly XName rotY = c + "rotY";
        public static readonly XName roundedCorners = c + "roundedCorners";
        public static readonly XName scaling = c + "scaling";
        public static readonly XName scatterChart = c + "scatterChart";
        public static readonly XName scatterStyle = c + "scatterStyle";
        public static readonly XName secondPiePt = c + "secondPiePt";
        public static readonly XName secondPieSize = c + "secondPieSize";
        public static readonly XName selection = c + "selection";
        public static readonly XName separator = c + "separator";
        public static readonly XName ser = c + "ser";
        public static readonly XName serAx = c + "serAx";
        public static readonly XName serLines = c + "serLines";
        public static readonly XName shape = c + "shape";
        public static readonly XName showBubbleSize = c + "showBubbleSize";
        public static readonly XName showCatName = c + "showCatName";
        public static readonly XName showDLblsOverMax = c + "showDLblsOverMax";
        public static readonly XName showHorzBorder = c + "showHorzBorder";
        public static readonly XName showKeys = c + "showKeys";
        public static readonly XName showLeaderLines = c + "showLeaderLines";
        public static readonly XName showLegendKey = c + "showLegendKey";
        public static readonly XName showNegBubbles = c + "showNegBubbles";
        public static readonly XName showOutline = c + "showOutline";
        public static readonly XName showPercent = c + "showPercent";
        public static readonly XName showSerName = c + "showSerName";
        public static readonly XName showVal = c + "showVal";
        public static readonly XName showVertBorder = c + "showVertBorder";
        public static readonly XName sideWall = c + "sideWall";
        public static readonly XName size = c + "size";
        public static readonly XName sizeRepresents = c + "sizeRepresents";
        public static readonly XName smooth = c + "smooth";
        public static readonly XName splitPos = c + "splitPos";
        public static readonly XName splitType = c + "splitType";
        public static readonly XName spPr = c + "spPr";
        public static readonly XName stockChart = c + "stockChart";
        public static readonly XName strCache = c + "strCache";
        public static readonly XName strLit = c + "strLit";
        public static readonly XName strRef = c + "strRef";
        public static readonly XName style = c + "style";
        public static readonly XName surface3DChart = c + "surface3DChart";
        public static readonly XName surfaceChart = c + "surfaceChart";
        public static readonly XName symbol = c + "symbol";
        public static readonly XName thickness = c + "thickness";
        public static readonly XName tickLblPos = c + "tickLblPos";
        public static readonly XName tickLblSkip = c + "tickLblSkip";
        public static readonly XName tickMarkSkip = c + "tickMarkSkip";
        public static readonly XName title = c + "title";
        public static readonly XName trendline = c + "trendline";
        public static readonly XName trendlineLbl = c + "trendlineLbl";
        public static readonly XName trendlineType = c + "trendlineType";
        public static readonly XName tx = c + "tx";
        public static readonly XName txPr = c + "txPr";
        public static readonly XName upBars = c + "upBars";
        public static readonly XName upDownBars = c + "upDownBars";
        public static readonly XName userInterface = c + "userInterface";
        public static readonly XName userShapes = c + "userShapes";
        public static readonly XName v = c + "v";
        public static readonly XName val = c + "val";
        public static readonly XName valAx = c + "valAx";
        public static readonly XName varyColors = c + "varyColors";
        public static readonly XName view3D = c + "view3D";
        public static readonly XName w = c + "w";
        public static readonly XName wireframe = c + "wireframe";
        public static readonly XName wMode = c + "wMode";
        public static readonly XName x = c + "x";
        public static readonly XName xMode = c + "xMode";
        public static readonly XName xVal = c + "xVal";
        public static readonly XName y = c + "y";
        public static readonly XName yMode = c + "yMode";
        public static readonly XName yVal = c + "yVal";
    }

    public static class CDR
    {
        public static readonly XNamespace cdr =
            "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing";
        public static readonly XName absSizeAnchor = cdr + "absSizeAnchor";
        public static readonly XName blipFill = cdr + "blipFill";
        public static readonly XName cNvCxnSpPr = cdr + "cNvCxnSpPr";
        public static readonly XName cNvGraphicFramePr = cdr + "cNvGraphicFramePr";
        public static readonly XName cNvGrpSpPr = cdr + "cNvGrpSpPr";
        public static readonly XName cNvPicPr = cdr + "cNvPicPr";
        public static readonly XName cNvPr = cdr + "cNvPr";
        public static readonly XName cNvSpPr = cdr + "cNvSpPr";
        public static readonly XName cxnSp = cdr + "cxnSp";
        public static readonly XName ext = cdr + "ext";
        public static readonly XName from = cdr + "from";
        public static readonly XName graphicFrame = cdr + "graphicFrame";
        public static readonly XName grpSp = cdr + "grpSp";
        public static readonly XName grpSpPr = cdr + "grpSpPr";
        public static readonly XName nvCxnSpPr = cdr + "nvCxnSpPr";
        public static readonly XName nvGraphicFramePr = cdr + "nvGraphicFramePr";
        public static readonly XName nvGrpSpPr = cdr + "nvGrpSpPr";
        public static readonly XName nvPicPr = cdr + "nvPicPr";
        public static readonly XName nvSpPr = cdr + "nvSpPr";
        public static readonly XName pic = cdr + "pic";
        public static readonly XName relSizeAnchor = cdr + "relSizeAnchor";
        public static readonly XName sp = cdr + "sp";
        public static readonly XName spPr = cdr + "spPr";
        public static readonly XName style = cdr + "style";
        public static readonly XName to = cdr + "to";
        public static readonly XName txBody = cdr + "txBody";
        public static readonly XName x = cdr + "x";
        public static readonly XName xfrm = cdr + "xfrm";
        public static readonly XName y = cdr + "y";
    }

    public static class COM
    {
        public static readonly XNamespace com =
            "http://schemas.openxmlformats.org/drawingml/2006/compatibility";
        public static readonly XName legacyDrawing = com + "legacyDrawing";
    }

    public static class CP
    {
        public static readonly XNamespace cp =
            "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        public static readonly XName category = cp + "category";
        public static readonly XName contentStatus = cp + "contentStatus";
        public static readonly XName contentType = cp + "contentType";
        public static readonly XName coreProperties = cp + "coreProperties";
        public static readonly XName keywords = cp + "keywords";
        public static readonly XName lastModifiedBy = cp + "lastModifiedBy";
        public static readonly XName lastPrinted = cp + "lastPrinted";
        public static readonly XName revision = cp + "revision";
    }

    public static class CUSTPRO
    {
        public static readonly XNamespace custpro =
            "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
        public static readonly XName Properties = custpro + "Properties";
        public static readonly XName property = custpro + "property";
    }

    public static class DC
    {
        public static readonly XNamespace dc =
            "http://purl.org/dc/elements/1.1/";
        public static readonly XName creator = dc + "creator";
        public static readonly XName description = dc + "description";
        public static readonly XName subject = dc + "subject";
        public static readonly XName title = dc + "title";
    }

    public static class DCTERMS
    {
        public static readonly XNamespace dcterms =
            "http://purl.org/dc/terms/";
        public static readonly XName created = dcterms + "created";
        public static readonly XName modified = dcterms + "modified";
    }

    public static class DGM
    {
        public static readonly XNamespace dgm =
            "http://schemas.openxmlformats.org/drawingml/2006/diagram";
        public static readonly XName adj = dgm + "adj";
        public static readonly XName adjLst = dgm + "adjLst";
        public static readonly XName alg = dgm + "alg";
        public static readonly XName animLvl = dgm + "animLvl";
        public static readonly XName animOne = dgm + "animOne";
        public static readonly XName bg = dgm + "bg";
        public static readonly XName bulletEnabled = dgm + "bulletEnabled";
        public static readonly XName cat = dgm + "cat";
        public static readonly XName catLst = dgm + "catLst";
        public static readonly XName chMax = dgm + "chMax";
        public static readonly XName choose = dgm + "choose";
        public static readonly XName chPref = dgm + "chPref";
        public static readonly XName clrData = dgm + "clrData";
        public static readonly XName colorsDef = dgm + "colorsDef";
        public static readonly XName constr = dgm + "constr";
        public static readonly XName constrLst = dgm + "constrLst";
        public static readonly XName cxn = dgm + "cxn";
        public static readonly XName cxnLst = dgm + "cxnLst";
        public static readonly XName dataModel = dgm + "dataModel";
        public static readonly XName desc = dgm + "desc";
        public static readonly XName dir = dgm + "dir";
        public static readonly XName effectClrLst = dgm + "effectClrLst";
        public static readonly XName _else = dgm + "else";
        public static readonly XName extLst = dgm + "extLst";
        public static readonly XName fillClrLst = dgm + "fillClrLst";
        public static readonly XName forEach = dgm + "forEach";
        public static readonly XName hierBranch = dgm + "hierBranch";
        public static readonly XName _if = dgm + "if";
        public static readonly XName layoutDef = dgm + "layoutDef";
        public static readonly XName layoutNode = dgm + "layoutNode";
        public static readonly XName linClrLst = dgm + "linClrLst";
        public static readonly XName orgChart = dgm + "orgChart";
        public static readonly XName param = dgm + "param";
        public static readonly XName presLayoutVars = dgm + "presLayoutVars";
        public static readonly XName presOf = dgm + "presOf";
        public static readonly XName prSet = dgm + "prSet";
        public static readonly XName pt = dgm + "pt";
        public static readonly XName ptLst = dgm + "ptLst";
        public static readonly XName relIds = dgm + "relIds";
        public static readonly XName resizeHandles = dgm + "resizeHandles";
        public static readonly XName rule = dgm + "rule";
        public static readonly XName ruleLst = dgm + "ruleLst";
        public static readonly XName sampData = dgm + "sampData";
        public static readonly XName scene3d = dgm + "scene3d";
        public static readonly XName shape = dgm + "shape";
        public static readonly XName sp3d = dgm + "sp3d";
        public static readonly XName spPr = dgm + "spPr";
        public static readonly XName style = dgm + "style";
        public static readonly XName styleData = dgm + "styleData";
        public static readonly XName styleDef = dgm + "styleDef";
        public static readonly XName styleLbl = dgm + "styleLbl";
        public static readonly XName t = dgm + "t";
        public static readonly XName title = dgm + "title";
        public static readonly XName txEffectClrLst = dgm + "txEffectClrLst";
        public static readonly XName txFillClrLst = dgm + "txFillClrLst";
        public static readonly XName txLinClrLst = dgm + "txLinClrLst";
        public static readonly XName txPr = dgm + "txPr";
        public static readonly XName varLst = dgm + "varLst";
        public static readonly XName whole = dgm + "whole";
    }

    public static class DGM14
    {
        public static readonly XNamespace dgm14 =
            "http://schemas.microsoft.com/office/drawing/2010/diagram";
        public static readonly XName cNvPr = dgm14 + "cNvPr";
        public static readonly XName recolorImg = dgm14 + "recolorImg";
    }

    public static class DIGSIG
    {
        public static readonly XNamespace digsig =
            "http://schemas.microsoft.com/office/2006/digsig";
        public static readonly XName ApplicationVersion = digsig + "ApplicationVersion";
        public static readonly XName ColorDepth = digsig + "ColorDepth";
        public static readonly XName HorizontalResolution = digsig + "HorizontalResolution";
        public static readonly XName ManifestHashAlgorithm = digsig + "ManifestHashAlgorithm";
        public static readonly XName Monitors = digsig + "Monitors";
        public static readonly XName OfficeVersion = digsig + "OfficeVersion";
        public static readonly XName SetupID = digsig + "SetupID";
        public static readonly XName SignatureComments = digsig + "SignatureComments";
        public static readonly XName SignatureImage = digsig + "SignatureImage";
        public static readonly XName SignatureInfoV1 = digsig + "SignatureInfoV1";
        public static readonly XName SignatureProviderDetails = digsig + "SignatureProviderDetails";
        public static readonly XName SignatureProviderId = digsig + "SignatureProviderId";
        public static readonly XName SignatureProviderUrl = digsig + "SignatureProviderUrl";
        public static readonly XName SignatureText = digsig + "SignatureText";
        public static readonly XName SignatureType = digsig + "SignatureType";
        public static readonly XName VerticalResolution = digsig + "VerticalResolution";
        public static readonly XName WindowsVersion = digsig + "WindowsVersion";
    }

    public static class DS
    {
        public static readonly XNamespace ds =
            "http://schemas.openxmlformats.org/officeDocument/2006/customXml";
        public static readonly XName datastoreItem = ds + "datastoreItem";
        public static readonly XName itemID = ds + "itemID";
        public static readonly XName schemaRef = ds + "schemaRef";
        public static readonly XName schemaRefs = ds + "schemaRefs";
        public static readonly XName uri = ds + "uri";
    }

    public static class DSP
    {
        public static readonly XNamespace dsp =
            "http://schemas.microsoft.com/office/drawing/2008/diagram";
        public static readonly XName dataModelExt = dsp + "dataModelExt";
    }

    public static class EP
    {
        public static readonly XNamespace ep =
            "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
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

    public static class LC
    {
        public static readonly XNamespace lc =
            "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas";
        public static readonly XName lockedCanvas = lc + "lockedCanvas";
    }

    public static class M
    {
        public static readonly XNamespace m =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";
        public static readonly XName acc = m + "acc";
        public static readonly XName accPr = m + "accPr";
        public static readonly XName aln = m + "aln";
        public static readonly XName alnAt = m + "alnAt";
        public static readonly XName alnScr = m + "alnScr";
        public static readonly XName argPr = m + "argPr";
        public static readonly XName argSz = m + "argSz";
        public static readonly XName bar = m + "bar";
        public static readonly XName barPr = m + "barPr";
        public static readonly XName baseJc = m + "baseJc";
        public static readonly XName begChr = m + "begChr";
        public static readonly XName borderBox = m + "borderBox";
        public static readonly XName borderBoxPr = m + "borderBoxPr";
        public static readonly XName box = m + "box";
        public static readonly XName boxPr = m + "boxPr";
        public static readonly XName brk = m + "brk";
        public static readonly XName brkBin = m + "brkBin";
        public static readonly XName brkBinSub = m + "brkBinSub";
        public static readonly XName cGp = m + "cGp";
        public static readonly XName cGpRule = m + "cGpRule";
        public static readonly XName chr = m + "chr";
        public static readonly XName count = m + "count";
        public static readonly XName cSp = m + "cSp";
        public static readonly XName ctrlPr = m + "ctrlPr";
        public static readonly XName d = m + "d";
        public static readonly XName defJc = m + "defJc";
        public static readonly XName deg = m + "deg";
        public static readonly XName degHide = m + "degHide";
        public static readonly XName den = m + "den";
        public static readonly XName diff = m + "diff";
        public static readonly XName dispDef = m + "dispDef";
        public static readonly XName dPr = m + "dPr";
        public static readonly XName e = m + "e";
        public static readonly XName endChr = m + "endChr";
        public static readonly XName eqArr = m + "eqArr";
        public static readonly XName eqArrPr = m + "eqArrPr";
        public static readonly XName f = m + "f";
        public static readonly XName fName = m + "fName";
        public static readonly XName fPr = m + "fPr";
        public static readonly XName func = m + "func";
        public static readonly XName funcPr = m + "funcPr";
        public static readonly XName groupChr = m + "groupChr";
        public static readonly XName groupChrPr = m + "groupChrPr";
        public static readonly XName grow = m + "grow";
        public static readonly XName hideBot = m + "hideBot";
        public static readonly XName hideLeft = m + "hideLeft";
        public static readonly XName hideRight = m + "hideRight";
        public static readonly XName hideTop = m + "hideTop";
        public static readonly XName interSp = m + "interSp";
        public static readonly XName intLim = m + "intLim";
        public static readonly XName intraSp = m + "intraSp";
        public static readonly XName jc = m + "jc";
        public static readonly XName lim = m + "lim";
        public static readonly XName limLoc = m + "limLoc";
        public static readonly XName limLow = m + "limLow";
        public static readonly XName limLowPr = m + "limLowPr";
        public static readonly XName limUpp = m + "limUpp";
        public static readonly XName limUppPr = m + "limUppPr";
        public static readonly XName lit = m + "lit";
        public static readonly XName lMargin = m + "lMargin";
        public static readonly XName _m = m + "m";
        public static readonly XName mathFont = m + "mathFont";
        public static readonly XName mathPr = m + "mathPr";
        public static readonly XName maxDist = m + "maxDist";
        public static readonly XName mc = m + "mc";
        public static readonly XName mcJc = m + "mcJc";
        public static readonly XName mcPr = m + "mcPr";
        public static readonly XName mcs = m + "mcs";
        public static readonly XName mPr = m + "mPr";
        public static readonly XName mr = m + "mr";
        public static readonly XName nary = m + "nary";
        public static readonly XName naryLim = m + "naryLim";
        public static readonly XName naryPr = m + "naryPr";
        public static readonly XName noBreak = m + "noBreak";
        public static readonly XName nor = m + "nor";
        public static readonly XName num = m + "num";
        public static readonly XName objDist = m + "objDist";
        public static readonly XName oMath = m + "oMath";
        public static readonly XName oMathPara = m + "oMathPara";
        public static readonly XName oMathParaPr = m + "oMathParaPr";
        public static readonly XName opEmu = m + "opEmu";
        public static readonly XName phant = m + "phant";
        public static readonly XName phantPr = m + "phantPr";
        public static readonly XName plcHide = m + "plcHide";
        public static readonly XName pos = m + "pos";
        public static readonly XName postSp = m + "postSp";
        public static readonly XName preSp = m + "preSp";
        public static readonly XName r = m + "r";
        public static readonly XName rad = m + "rad";
        public static readonly XName radPr = m + "radPr";
        public static readonly XName rMargin = m + "rMargin";
        public static readonly XName rPr = m + "rPr";
        public static readonly XName rSp = m + "rSp";
        public static readonly XName rSpRule = m + "rSpRule";
        public static readonly XName scr = m + "scr";
        public static readonly XName sepChr = m + "sepChr";
        public static readonly XName show = m + "show";
        public static readonly XName shp = m + "shp";
        public static readonly XName smallFrac = m + "smallFrac";
        public static readonly XName sPre = m + "sPre";
        public static readonly XName sPrePr = m + "sPrePr";
        public static readonly XName sSub = m + "sSub";
        public static readonly XName sSubPr = m + "sSubPr";
        public static readonly XName sSubSup = m + "sSubSup";
        public static readonly XName sSubSupPr = m + "sSubSupPr";
        public static readonly XName sSup = m + "sSup";
        public static readonly XName sSupPr = m + "sSupPr";
        public static readonly XName strikeBLTR = m + "strikeBLTR";
        public static readonly XName strikeH = m + "strikeH";
        public static readonly XName strikeTLBR = m + "strikeTLBR";
        public static readonly XName strikeV = m + "strikeV";
        public static readonly XName sty = m + "sty";
        public static readonly XName sub = m + "sub";
        public static readonly XName subHide = m + "subHide";
        public static readonly XName sup = m + "sup";
        public static readonly XName supHide = m + "supHide";
        public static readonly XName t = m + "t";
        public static readonly XName transp = m + "transp";
        public static readonly XName type = m + "type";
        public static readonly XName val = m + "val";
        public static readonly XName vertJc = m + "vertJc";
        public static readonly XName wrapIndent = m + "wrapIndent";
        public static readonly XName wrapRight = m + "wrapRight";
        public static readonly XName zeroAsc = m + "zeroAsc";
        public static readonly XName zeroDesc = m + "zeroDesc";
        public static readonly XName zeroWid = m + "zeroWid";
    }

    public static class MC
    {
        public static readonly XNamespace mc =
            "http://schemas.openxmlformats.org/markup-compatibility/2006";
        public static readonly XName AlternateContent = mc + "AlternateContent";
        public static readonly XName Choice = mc + "Choice";
        public static readonly XName Fallback = mc + "Fallback";
        public static readonly XName Ignorable = mc + "Ignorable";
        public static readonly XName PreserveAttributes = mc + "PreserveAttributes";
    }

    public static class MDSSI
    {
        public static readonly XNamespace mdssi =
            "http://schemas.openxmlformats.org/package/2006/digital-signature";
        public static readonly XName Format = mdssi + "Format";
        public static readonly XName RelationshipReference = mdssi + "RelationshipReference";
        public static readonly XName SignatureTime = mdssi + "SignatureTime";
        public static readonly XName Value = mdssi + "Value";
    }

    public static class MP
    {
        public static readonly XNamespace mp =
            "http://schemas.microsoft.com/office/mac/powerpoint/2008/main";
        public static readonly XName cube = mp + "cube";
        public static readonly XName flip = mp + "flip";
        public static readonly XName transition = mp + "transition";
    }

    public static class MV
    {
        public static readonly XNamespace mv =
            "urn:schemas-microsoft-com:mac:vml";
        public static readonly XName blur = mv + "blur";
        public static readonly XName complextextbox = mv + "complextextbox";
    }

    public static class NoNamespace
    {
        public static readonly XName a = "a";
        public static readonly XName accent1 = "accent1";
        public static readonly XName accent2 = "accent2";
        public static readonly XName accent3 = "accent3";
        public static readonly XName accent4 = "accent4";
        public static readonly XName accent5 = "accent5";
        public static readonly XName accent6 = "accent6";
        public static readonly XName action = "action";
        public static readonly XName activeCell = "activeCell";
        public static readonly XName activeCol = "activeCol";
        public static readonly XName activePane = "activePane";
        public static readonly XName activeRow = "activeRow";
        public static readonly XName advise = "advise";
        public static readonly XName algn = "algn";
        public static readonly XName Algorithm = "Algorithm";
        public static readonly XName alignWithMargins = "alignWithMargins";
        public static readonly XName allowcomments = "allowcomments";
        public static readonly XName allowOverlap = "allowOverlap";
        public static readonly XName allUniqueName = "allUniqueName";
        public static readonly XName alt = "alt";
        public static readonly XName alwaysShow = "alwaysShow";
        public static readonly XName amount = "amount";
        public static readonly XName amt = "amt";
        public static readonly XName anchor = "anchor";
        public static readonly XName anchorCtr = "anchorCtr";
        public static readonly XName ang = "ang";
        public static readonly XName animBg = "animBg";
        public static readonly XName annotation = "annotation";
        public static readonly XName applyAlignment = "applyAlignment";
        public static readonly XName applyAlignmentFormats = "applyAlignmentFormats";
        public static readonly XName applyBorder = "applyBorder";
        public static readonly XName applyBorderFormats = "applyBorderFormats";
        public static readonly XName applyFill = "applyFill";
        public static readonly XName applyFont = "applyFont";
        public static readonly XName applyFontFormats = "applyFontFormats";
        public static readonly XName applyNumberFormat = "applyNumberFormat";
        public static readonly XName applyNumberFormats = "applyNumberFormats";
        public static readonly XName applyPatternFormats = "applyPatternFormats";
        public static readonly XName applyProtection = "applyProtection";
        public static readonly XName applyWidthHeightFormats = "applyWidthHeightFormats";
        public static readonly XName arcsize = "arcsize";
        public static readonly XName arg = "arg";
        public static readonly XName aspectratio = "aspectratio";
        public static readonly XName assign = "assign";
        public static readonly XName attribute = "attribute";
        public static readonly XName author = "author";
        public static readonly XName authorId = "authorId";
        public static readonly XName auto = "auto";
        public static readonly XName autoEnd = "autoEnd";
        public static readonly XName autoFormatId = "autoFormatId";
        public static readonly XName autoLine = "autoLine";
        public static readonly XName autoStart = "autoStart";
        public static readonly XName axis = "axis";
        public static readonly XName b = "b";
        public static readonly XName backdepth = "backdepth";
        public static readonly XName bandRow = "bandRow";
        public static readonly XName _base = "base";
        public static readonly XName baseField = "baseField";
        public static readonly XName baseItem = "baseItem";
        public static readonly XName baseline = "baseline";
        public static readonly XName baseType = "baseType";
        public static readonly XName behindDoc = "behindDoc";
        public static readonly XName bestFit = "bestFit";
        public static readonly XName bg1 = "bg1";
        public static readonly XName bg2 = "bg2";
        public static readonly XName bIns = "bIns";
        public static readonly XName bld = "bld";
        public static readonly XName bldStep = "bldStep";
        public static readonly XName blend = "blend";
        public static readonly XName blurRad = "blurRad";
        public static readonly XName bmkName = "bmkName";
        public static readonly XName borderId = "borderId";
        public static readonly XName bottom = "bottom";
        public static readonly XName bright = "bright";
        public static readonly XName brightness = "brightness";
        public static readonly XName builtinId = "builtinId";
        public static readonly XName bwMode = "bwMode";
        public static readonly XName by = "by";
        public static readonly XName c = "c";
        public static readonly XName cacheId = "cacheId";
        public static readonly XName cacheIndex = "cacheIndex";
        public static readonly XName calcmode = "calcmode";
        public static readonly XName cap = "cap";
        public static readonly XName caption = "caption";
        public static readonly XName categoryIdx = "categoryIdx";
        public static readonly XName cell = "cell";
        public static readonly XName cellColor = "cellColor";
        public static readonly XName cellRange = "cellRange";
        public static readonly XName _char = "char";
        public static readonly XName charset = "charset";
        public static readonly XName chart = "chart";
        public static readonly XName clearComments = "clearComments";
        public static readonly XName clearFormats = "clearFormats";
        public static readonly XName click = "click";
        public static readonly XName clientInsertedTime = "clientInsertedTime";
        public static readonly XName clrIdx = "clrIdx";
        public static readonly XName clrSpc = "clrSpc";
        public static readonly XName cmd = "cmd";
        public static readonly XName cmpd = "cmpd";
        public static readonly XName codeName = "codeName";
        public static readonly XName coerce = "coerce";
        public static readonly XName colId = "colId";
        public static readonly XName color = "color";
        public static readonly XName colors = "colors";
        public static readonly XName colorTemp = "colorTemp";
        public static readonly XName colPageCount = "colPageCount";
        public static readonly XName cols = "cols";
        public static readonly XName comma = "comma";
        public static readonly XName command = "command";
        public static readonly XName commandType = "commandType";
        public static readonly XName comment = "comment";
        public static readonly XName compatLnSpc = "compatLnSpc";
        public static readonly XName concurrent = "concurrent";
        public static readonly XName connection = "connection";
        public static readonly XName connectionId = "connectionId";
        public static readonly XName connectloc = "connectloc";
        public static readonly XName consecutive = "consecutive";
        public static readonly XName constrainbounds = "constrainbounds";
        public static readonly XName containsInteger = "containsInteger";
        public static readonly XName containsNumber = "containsNumber";
        public static readonly XName containsSemiMixedTypes = "containsSemiMixedTypes";
        public static readonly XName containsString = "containsString";
        public static readonly XName contrast = "contrast";
        public static readonly XName control1 = "control1";
        public static readonly XName control2 = "control2";
        public static readonly XName coordorigin = "coordorigin";
        public static readonly XName coordsize = "coordsize";
        public static readonly XName copy = "copy";
        public static readonly XName count = "count";
        public static readonly XName createdVersion = "createdVersion";
        public static readonly XName cryptAlgorithmClass = "cryptAlgorithmClass";
        public static readonly XName cryptAlgorithmSid = "cryptAlgorithmSid";
        public static readonly XName cryptAlgorithmType = "cryptAlgorithmType";
        public static readonly XName cryptProviderType = "cryptProviderType";
        public static readonly XName csCatId = "csCatId";
        public static readonly XName cstate = "cstate";
        public static readonly XName csTypeId = "csTypeId";
        public static readonly XName culture = "culture";
        public static readonly XName current = "current";
        public static readonly XName customFormat = "customFormat";
        public static readonly XName customList = "customList";
        public static readonly XName customWidth = "customWidth";
        public static readonly XName cx = "cx";
        public static readonly XName cy = "cy";
        public static readonly XName d = "d";
        public static readonly XName data = "data";
        public static readonly XName dataCaption = "dataCaption";
        public static readonly XName dataDxfId = "dataDxfId";
        public static readonly XName dataField = "dataField";
        public static readonly XName dateTime = "dateTime";
        public static readonly XName dateTimeGrouping = "dateTimeGrouping";
        public static readonly XName dde = "dde";
        public static readonly XName ddeService = "ddeService";
        public static readonly XName ddeTopic = "ddeTopic";
        public static readonly XName def = "def";
        public static readonly XName defaultMemberUniqueName = "defaultMemberUniqueName";
        public static readonly XName defaultPivotStyle = "defaultPivotStyle";
        public static readonly XName defaultRowHeight = "defaultRowHeight";
        public static readonly XName defaultSize = "defaultSize";
        public static readonly XName defaultTableStyle = "defaultTableStyle";
        public static readonly XName defStyle = "defStyle";
        public static readonly XName defTabSz = "defTabSz";
        public static readonly XName degree = "degree";
        public static readonly XName delay = "delay";
        public static readonly XName descending = "descending";
        public static readonly XName descr = "descr";
        public static readonly XName destId = "destId";
        public static readonly XName destination = "destination";
        public static readonly XName destinationFile = "destinationFile";
        public static readonly XName destOrd = "destOrd";
        public static readonly XName dgmfontsize = "dgmfontsize";
        public static readonly XName dgmstyle = "dgmstyle";
        public static readonly XName diagonalDown = "diagonalDown";
        public static readonly XName diagonalUp = "diagonalUp";
        public static readonly XName dimension = "dimension";
        public static readonly XName dimensionUniqueName = "dimensionUniqueName";
        public static readonly XName dir = "dir";
        public static readonly XName dirty = "dirty";
        public static readonly XName display = "display";
        public static readonly XName displayFolder = "displayFolder";
        public static readonly XName displayName = "displayName";
        public static readonly XName dist = "dist";
        public static readonly XName distB = "distB";
        public static readonly XName distL = "distL";
        public static readonly XName distR = "distR";
        public static readonly XName distT = "distT";
        public static readonly XName divId = "divId";
        public static readonly XName dpi = "dpi";
        public static readonly XName dr = "dr";
        public static readonly XName DrawAspect = "DrawAspect";
        public static readonly XName dt = "dt";
        public static readonly XName dur = "dur";
        public static readonly XName dx = "dx";
        public static readonly XName dxfId = "dxfId";
        public static readonly XName dy = "dy";
        public static readonly XName dz = "dz";
        public static readonly XName eaLnBrk = "eaLnBrk";
        public static readonly XName eb = "eb";
        public static readonly XName edited = "edited";
        public static readonly XName editPage = "editPage";
        public static readonly XName end = "end";
        public static readonly XName endA = "endA";
        public static readonly XName endangle = "endangle";
        public static readonly XName endDate = "endDate";
        public static readonly XName endPos = "endPos";
        public static readonly XName endSnd = "endSnd";
        public static readonly XName eqn = "eqn";
        public static readonly XName evalOrder = "evalOrder";
        public static readonly XName evt = "evt";
        public static readonly XName exp = "exp";
        public static readonly XName extProperty = "extProperty";
        public static readonly XName f = "f";
        public static readonly XName fact = "fact";
        public static readonly XName field = "field";
        public static readonly XName fieldId = "fieldId";
        public static readonly XName fieldListSortAscending = "fieldListSortAscending";
        public static readonly XName fieldPosition = "fieldPosition";
        public static readonly XName fileType = "fileType";
        public static readonly XName fillcolor = "fillcolor";
        public static readonly XName filled = "filled";
        public static readonly XName fillId = "fillId";
        public static readonly XName filter = "filter";
        public static readonly XName filterVal = "filterVal";
        public static readonly XName first = "first";
        public static readonly XName firstDataCol = "firstDataCol";
        public static readonly XName firstDataRow = "firstDataRow";
        public static readonly XName firstHeaderRow = "firstHeaderRow";
        public static readonly XName firstRow = "firstRow";
        public static readonly XName fitshape = "fitshape";
        public static readonly XName fitToPage = "fitToPage";
        public static readonly XName fld = "fld";
        public static readonly XName flip = "flip";
        public static readonly XName fmla = "fmla";
        public static readonly XName fmtid = "fmtid";
        public static readonly XName folHlink = "folHlink";
        public static readonly XName followColorScheme = "followColorScheme";
        public static readonly XName fontId = "fontId";
        public static readonly XName footer = "footer";
        public static readonly XName _for = "for";
        public static readonly XName forceAA = "forceAA";
        public static readonly XName format = "format";
        public static readonly XName formatCode = "formatCode";
        public static readonly XName formula = "formula";
        public static readonly XName forName = "forName";
        public static readonly XName fov = "fov";
        public static readonly XName frame = "frame";
        public static readonly XName from = "from";
        public static readonly XName fromWordArt = "fromWordArt";
        public static readonly XName fullCalcOnLoad = "fullCalcOnLoad";
        public static readonly XName func = "func";
        public static readonly XName g = "g";
        public static readonly XName gdRefAng = "gdRefAng";
        public static readonly XName gdRefR = "gdRefR";
        public static readonly XName gdRefX = "gdRefX";
        public static readonly XName gdRefY = "gdRefY";
        public static readonly XName goal = "goal";
        public static readonly XName gradientshapeok = "gradientshapeok";
        public static readonly XName groupBy = "groupBy";
        public static readonly XName grpId = "grpId";
        public static readonly XName guid = "guid";
        public static readonly XName h = "h";
        public static readonly XName hangingPunct = "hangingPunct";
        public static readonly XName hashData = "hashData";
        public static readonly XName header = "header";
        public static readonly XName headerRowBorderDxfId = "headerRowBorderDxfId";
        public static readonly XName headerRowDxfId = "headerRowDxfId";
        public static readonly XName hidden = "hidden";
        public static readonly XName hier = "hier";
        public static readonly XName hierarchy = "hierarchy";
        public static readonly XName hierarchyUsage = "hierarchyUsage";
        public static readonly XName highlightClick = "highlightClick";
        public static readonly XName hlink = "hlink";
        public static readonly XName horizontal = "horizontal";
        public static readonly XName horizontalCentered = "horizontalCentered";
        public static readonly XName horizontalDpi = "horizontalDpi";
        public static readonly XName horzOverflow = "horzOverflow";
        public static readonly XName href = "href";
        public static readonly XName hR = "hR";
        public static readonly XName htmlFormat = "htmlFormat";
        public static readonly XName htmlTables = "htmlTables";
        public static readonly XName hue = "hue";
        public static readonly XName i = "i";
        public static readonly XName i1 = "i1";
        public static readonly XName iconId = "iconId";
        public static readonly XName iconSet = "iconSet";
        public static readonly XName id = "id";
        public static readonly XName Id = "Id";
        public static readonly XName iddest = "iddest";
        public static readonly XName idref = "idref";
        public static readonly XName idsrc = "idsrc";
        public static readonly XName idx = "idx";
        public static readonly XName imgH = "imgH";
        public static readonly XName imgW = "imgW";
        public static readonly XName _in = "in";
        public static readonly XName includeNewItemsInFilter = "includeNewItemsInFilter";
        public static readonly XName indent = "indent";
        public static readonly XName index = "index";
        public static readonly XName indexed = "indexed";
        public static readonly XName initials = "initials";
        public static readonly XName insetpen = "insetpen";
        public static readonly XName invalEndChars = "invalEndChars";
        public static readonly XName invalidUrl = "invalidUrl";
        public static readonly XName invalStChars = "invalStChars";
        public static readonly XName isInverted = "isInverted";
        public static readonly XName issignatureline = "issignatureline";
        public static readonly XName item = "item";
        public static readonly XName itemPrintTitles = "itemPrintTitles";
        public static readonly XName joinstyle = "joinstyle";
        public static readonly XName justifyLastLine = "justifyLastLine";
        public static readonly XName key = "key";
        public static readonly XName keyAttribute = "keyAttribute";
        public static readonly XName l = "l";
        public static readonly XName lang = "lang";
        public static readonly XName lastClr = "lastClr";
        public static readonly XName lastIdx = "lastIdx";
        public static readonly XName lat = "lat";
        public static readonly XName latinLnBrk = "latinLnBrk";
        public static readonly XName layout = "layout";
        public static readonly XName layoutInCell = "layoutInCell";
        public static readonly XName left = "left";
        public static readonly XName len = "len";
        public static readonly XName length = "length";
        public static readonly XName level = "level";
        public static readonly XName lightharsh2 = "lightharsh2";
        public static readonly XName lightlevel = "lightlevel";
        public static readonly XName lightlevel2 = "lightlevel2";
        public static readonly XName lightposition = "lightposition";
        public static readonly XName lightposition2 = "lightposition2";
        public static readonly XName lim = "lim";
        public static readonly XName link = "link";
        public static readonly XName lIns = "lIns";
        public static readonly XName loCatId = "loCatId";
        public static readonly XName locked = "locked";
        public static readonly XName lon = "lon";
        public static readonly XName loop = "loop";
        public static readonly XName loTypeId = "loTypeId";
        public static readonly XName lum = "lum";
        public static readonly XName lvl = "lvl";
        public static readonly XName macro = "macro";
        public static readonly XName man = "man";
        public static readonly XName manualBreakCount = "manualBreakCount";
        public static readonly XName mapId = "mapId";
        public static readonly XName marL = "marL";
        public static readonly XName max = "max";
        public static readonly XName maxAng = "maxAng";
        public static readonly XName maxR = "maxR";
        public static readonly XName maxRank = "maxRank";
        public static readonly XName maxSheetId = "maxSheetId";
        public static readonly XName maxValue = "maxValue";
        public static readonly XName maxX = "maxX";
        public static readonly XName maxY = "maxY";
        public static readonly XName mdx = "mdx";
        public static readonly XName measureGroup = "measureGroup";
        public static readonly XName memberName = "memberName";
        public static readonly XName merge = "merge";
        public static readonly XName meth = "meth";
        public static readonly XName min = "min";
        public static readonly XName minAng = "minAng";
        public static readonly XName minR = "minR";
        public static readonly XName minRefreshableVersion = "minRefreshableVersion";
        public static readonly XName minSupportedVersion = "minSupportedVersion";
        public static readonly XName minValue = "minValue";
        public static readonly XName minVer = "minVer";
        public static readonly XName minX = "minX";
        public static readonly XName minY = "minY";
        public static readonly XName modelId = "modelId";
        public static readonly XName moveWithCells = "moveWithCells";
        public static readonly XName n = "n";
        public static readonly XName name = "name";
        public static readonly XName _new = "new";
        public static readonly XName newLength = "newLength";
        public static readonly XName newName = "newName";
        public static readonly XName nextAc = "nextAc";
        public static readonly XName nextId = "nextId";
        public static readonly XName noChangeArrowheads = "noChangeArrowheads";
        public static readonly XName noChangeAspect = "noChangeAspect";
        public static readonly XName noChangeShapeType = "noChangeShapeType";
        public static readonly XName nodeType = "nodeType";
        public static readonly XName noEditPoints = "noEditPoints";
        public static readonly XName noGrp = "noGrp";
        public static readonly XName noRot = "noRot";
        public static readonly XName noUngrp = "noUngrp";
        public static readonly XName np = "np";
        public static readonly XName ns = "ns";
        public static readonly XName numCol = "numCol";
        public static readonly XName numFmtId = "numFmtId";
        public static readonly XName o = "o";
        public static readonly XName ObjectID = "ObjectID";
        public static readonly XName objects = "objects";
        public static readonly XName ObjectType = "ObjectType";
        public static readonly XName objId = "objId";
        public static readonly XName offset = "offset";
        public static readonly XName old = "old";
        public static readonly XName oldComment = "oldComment";
        public static readonly XName oldName = "oldName";
        public static readonly XName oleUpdate = "oleUpdate";
        public static readonly XName on = "on";
        public static readonly XName op = "op";
        public static readonly XName orient = "orient";
        public static readonly XName orientation = "orientation";
        public static readonly XName origin = "origin";
        public static readonly XName _out = "out";
        public static readonly XName outline = "outline";
        public static readonly XName outlineData = "outlineData";
        public static readonly XName p = "p";
        public static readonly XName pane = "pane";
        public static readonly XName panose = "panose";
        public static readonly XName paperSize = "paperSize";
        public static readonly XName par = "par";
        public static readonly XName parameterType = "parameterType";
        public static readonly XName parent = "parent";
        public static readonly XName password = "password";
        public static readonly XName pasteAll = "pasteAll";
        public static readonly XName pasteValues = "pasteValues";
        public static readonly XName path = "path";
        public static readonly XName pathEditMode = "pathEditMode";
        public static readonly XName patternType = "patternType";
        public static readonly XName phldr = "phldr";
        public static readonly XName pid = "pid";
        public static readonly XName pitchFamily = "pitchFamily";
        public static readonly XName pivot = "pivot";
        public static readonly XName points = "points";
        public static readonly XName pos = "pos";
        public static readonly XName position = "position";
        public static readonly XName post = "post";
        public static readonly XName preferPic = "preferPic";
        public static readonly XName preserve = "preserve";
        public static readonly XName pressure = "pressure";
        public static readonly XName previousCol = "previousCol";
        public static readonly XName previousRow = "previousRow";
        public static readonly XName pri = "pri";
        public static readonly XName priority = "priority";
        public static readonly XName progId = "progId";
        public static readonly XName ProgID = "ProgID";
        public static readonly XName provid = "provid";
        public static readonly XName prst = "prst";
        public static readonly XName prstMaterial = "prstMaterial";
        public static readonly XName ptsTypes = "ptsTypes";
        public static readonly XName ptType = "ptType";
        public static readonly XName qsCatId = "qsCatId";
        public static readonly XName qsTypeId = "qsTypeId";
        public static readonly XName r = "r";
        public static readonly XName rad = "rad";
        public static readonly XName readingOrder = "readingOrder";
        public static readonly XName recordCount = "recordCount";
        public static readonly XName _ref = "ref";
        public static readonly XName ref3D = "ref3D";
        public static readonly XName refersTo = "refersTo";
        public static readonly XName refreshedBy = "refreshedBy";
        public static readonly XName refreshedDate = "refreshedDate";
        public static readonly XName refreshedVersion = "refreshedVersion";
        public static readonly XName refreshOnLoad = "refreshOnLoad";
        public static readonly XName refType = "refType";
        public static readonly XName relativeFrom = "relativeFrom";
        public static readonly XName relativeHeight = "relativeHeight";
        public static readonly XName relId = "relId";
        public static readonly XName Requires = "Requires";
        public static readonly XName restart = "restart";
        public static readonly XName rev = "rev";
        public static readonly XName rgb = "rgb";
        public static readonly XName rId = "rId";
        public static readonly XName rig = "rig";
        public static readonly XName right = "right";
        public static readonly XName rIns = "rIns";
        public static readonly XName rot = "rot";
        public static readonly XName rotWithShape = "rotWithShape";
        public static readonly XName rowColShift = "rowColShift";
        public static readonly XName rowDrillCount = "rowDrillCount";
        public static readonly XName rowPageCount = "rowPageCount";
        public static readonly XName rows = "rows";
        public static readonly XName rtl = "rtl";
        public static readonly XName rtlCol = "rtlCol";
        public static readonly XName s = "s";
        public static readonly XName saltData = "saltData";
        public static readonly XName sat = "sat";
        public static readonly XName saveData = "saveData";
        public static readonly XName saveSubsetFonts = "saveSubsetFonts";
        public static readonly XName sb = "sb";
        public static readonly XName scaled = "scaled";
        public static readonly XName scaling = "scaling";
        public static readonly XName scenarios = "scenarios";
        public static readonly XName scope = "scope";
        public static readonly XName script = "script";
        public static readonly XName securityDescriptor = "securityDescriptor";
        public static readonly XName seek = "seek";
        public static readonly XName sendLocale = "sendLocale";
        public static readonly XName series = "series";
        public static readonly XName seriesIdx = "seriesIdx";
        public static readonly XName serverSldId = "serverSldId";
        public static readonly XName serverSldModifiedTime = "serverSldModifiedTime";
        public static readonly XName setDefinition = "setDefinition";
        public static readonly XName shapeId = "shapeId";
        public static readonly XName ShapeID = "ShapeID";
        public static readonly XName sheet = "sheet";
        public static readonly XName sheetId = "sheetId";
        public static readonly XName sheetPosition = "sheetPosition";
        public static readonly XName show = "show";
        public static readonly XName showAll = "showAll";
        public static readonly XName showCaptions = "showCaptions";
        public static readonly XName showColHeaders = "showColHeaders";
        public static readonly XName showColStripes = "showColStripes";
        public static readonly XName showColumnStripes = "showColumnStripes";
        public static readonly XName showErrorMessage = "showErrorMessage";
        public static readonly XName showFirstColumn = "showFirstColumn";
        public static readonly XName showHeader = "showHeader";
        public static readonly XName showInputMessage = "showInputMessage";
        public static readonly XName showLastColumn = "showLastColumn";
        public static readonly XName showRowHeaders = "showRowHeaders";
        public static readonly XName showRowStripes = "showRowStripes";
        public static readonly XName showValue = "showValue";
        public static readonly XName shrinkToFit = "shrinkToFit";
        public static readonly XName si = "si";
        public static readonly XName sId = "sId";
        public static readonly XName simplePos = "simplePos";
        public static readonly XName size = "size";
        public static readonly XName skewangle = "skewangle";
        public static readonly XName smoothness = "smoothness";
        public static readonly XName smtClean = "smtClean";
        public static readonly XName source = "source";
        public static readonly XName sourceFile = "sourceFile";
        public static readonly XName SourceId = "SourceId";
        public static readonly XName sourceLinked = "sourceLinked";
        public static readonly XName sourceSheetId = "sourceSheetId";
        public static readonly XName sourceType = "sourceType";
        public static readonly XName sp = "sp";
        public static readonly XName spans = "spans";
        public static readonly XName spcCol = "spcCol";
        public static readonly XName spcFirstLastPara = "spcFirstLastPara";
        public static readonly XName spid = "spid";
        public static readonly XName spidmax = "spidmax";
        public static readonly XName spinCount = "spinCount";
        public static readonly XName splitFirst = "splitFirst";
        public static readonly XName spokes = "spokes";
        public static readonly XName sqlType = "sqlType";
        public static readonly XName sqref = "sqref";
        public static readonly XName src = "src";
        public static readonly XName srcId = "srcId";
        public static readonly XName srcOrd = "srcOrd";
        public static readonly XName st = "st";
        public static readonly XName stA = "stA";
        public static readonly XName stAng = "stAng";
        public static readonly XName start = "start";
        public static readonly XName startangle = "startangle";
        public static readonly XName startDate = "startDate";
        public static readonly XName status = "status";
        public static readonly XName strike = "strike";
        public static readonly XName _string = "string";
        public static readonly XName strokecolor = "strokecolor";
        public static readonly XName stroked = "stroked";
        public static readonly XName strokeweight = "strokeweight";
        public static readonly XName style = "style";
        public static readonly XName styleId = "styleId";
        public static readonly XName styleName = "styleName";
        public static readonly XName subtotal = "subtotal";
        public static readonly XName summaryBelow = "summaryBelow";
        public static readonly XName swAng = "swAng";
        public static readonly XName sx = "sx";
        public static readonly XName sy = "sy";
        public static readonly XName sz = "sz";
        public static readonly XName t = "t";
        public static readonly XName tab = "tab";
        public static readonly XName tableBorderDxfId = "tableBorderDxfId";
        public static readonly XName tableColumnId = "tableColumnId";
        public static readonly XName Target = "Target";
        public static readonly XName textlink = "textlink";
        public static readonly XName textRotation = "textRotation";
        public static readonly XName theme = "theme";
        public static readonly XName thresh = "thresh";
        public static readonly XName thruBlk = "thruBlk";
        public static readonly XName time = "time";
        public static readonly XName tIns = "tIns";
        public static readonly XName tint = "tint";
        public static readonly XName tm = "tm";
        public static readonly XName to = "to";
        public static readonly XName tooltip = "tooltip";
        public static readonly XName top = "top";
        public static readonly XName topLabels = "topLabels";
        public static readonly XName topLeftCell = "topLeftCell";
        public static readonly XName totalsRowShown = "totalsRowShown";
        public static readonly XName track = "track";
        public static readonly XName trans = "trans";
        public static readonly XName transition = "transition";
        public static readonly XName trend = "trend";
        public static readonly XName twoDigitTextYear = "twoDigitTextYear";
        public static readonly XName tx = "tx";
        public static readonly XName tx1 = "tx1";
        public static readonly XName tx2 = "tx2";
        public static readonly XName txBox = "txBox";
        public static readonly XName txbxSeq = "txbxSeq";
        public static readonly XName txbxStory = "txbxStory";
        public static readonly XName ty = "ty";
        public static readonly XName type = "type";
        public static readonly XName Type = "Type";
        public static readonly XName typeface = "typeface";
        public static readonly XName u = "u";
        public static readonly XName ua = "ua";
        public static readonly XName uiExpand = "uiExpand";
        public static readonly XName unbalanced = "unbalanced";
        public static readonly XName uniqueCount = "uniqueCount";
        public static readonly XName uniqueId = "uniqueId";
        public static readonly XName uniqueName = "uniqueName";
        public static readonly XName uniqueParent = "uniqueParent";
        public static readonly XName updateAutomatic = "updateAutomatic";
        public static readonly XName updatedVersion = "updatedVersion";
        public static readonly XName uri = "uri";
        public static readonly XName URI = "URI";
        public static readonly XName url = "url";
        public static readonly XName useAutoFormatting = "useAutoFormatting";
        public static readonly XName useDef = "useDef";
        public static readonly XName user = "user";
        public static readonly XName userName = "userName";
        public static readonly XName v = "v";
        public static readonly XName val = "val";
        public static readonly XName value = "value";
        public static readonly XName valueType = "valueType";
        public static readonly XName varScale = "varScale";
        public static readonly XName vert = "vert";
        public static readonly XName vertical = "vertical";
        public static readonly XName verticalCentered = "verticalCentered";
        public static readonly XName verticalDpi = "verticalDpi";
        public static readonly XName vertOverflow = "vertOverflow";
        public static readonly XName viewpoint = "viewpoint";
        public static readonly XName viewpointorigin = "viewpointorigin";
        public static readonly XName w = "w";
        public static readonly XName weight = "weight";
        public static readonly XName width = "width";
        public static readonly XName workbookViewId = "workbookViewId";
        public static readonly XName wR = "wR";
        public static readonly XName wrap = "wrap";
        public static readonly XName wrapText = "wrapText";
        public static readonly XName x = "x";
        public static readonly XName x1 = "x1";
        public static readonly XName x2 = "x2";
        public static readonly XName xfId = "xfId";
        public static readonly XName xl97 = "xl97";
        public static readonly XName xmlDataType = "xmlDataType";
        public static readonly XName xpath = "xpath";
        public static readonly XName xSplit = "xSplit";
        public static readonly XName y = "y";
        public static readonly XName y1 = "y1";
        public static readonly XName y2 = "y2";
        public static readonly XName year = "year";
        public static readonly XName yrange = "yrange";
        public static readonly XName ySplit = "ySplit";
        public static readonly XName z = "z";
    }

    public static class O
    {
        public static readonly XNamespace o =
            "urn:schemas-microsoft-com:office:office";
        public static readonly XName allowincell = o + "allowincell";
        public static readonly XName allowoverlap = o + "allowoverlap";
        public static readonly XName althref = o + "althref";
        public static readonly XName borderbottomcolor = o + "borderbottomcolor";
        public static readonly XName borderleftcolor = o + "borderleftcolor";
        public static readonly XName borderrightcolor = o + "borderrightcolor";
        public static readonly XName bordertopcolor = o + "bordertopcolor";
        public static readonly XName bottom = o + "bottom";
        public static readonly XName bullet = o + "bullet";
        public static readonly XName button = o + "button";
        public static readonly XName bwmode = o + "bwmode";
        public static readonly XName bwnormal = o + "bwnormal";
        public static readonly XName bwpure = o + "bwpure";
        public static readonly XName callout = o + "callout";
        public static readonly XName clip = o + "clip";
        public static readonly XName clippath = o + "clippath";
        public static readonly XName cliptowrap = o + "cliptowrap";
        public static readonly XName colormenu = o + "colormenu";
        public static readonly XName colormru = o + "colormru";
        public static readonly XName column = o + "column";
        public static readonly XName complex = o + "complex";
        public static readonly XName connectangles = o + "connectangles";
        public static readonly XName connectlocs = o + "connectlocs";
        public static readonly XName connectortype = o + "connectortype";
        public static readonly XName connecttype = o + "connecttype";
        public static readonly XName detectmouseclick = o + "detectmouseclick";
        public static readonly XName dgmlayout = o + "dgmlayout";
        public static readonly XName dgmlayoutmru = o + "dgmlayoutmru";
        public static readonly XName dgmnodekind = o + "dgmnodekind";
        public static readonly XName diagram = o + "diagram";
        public static readonly XName doubleclicknotify = o + "doubleclicknotify";
        public static readonly XName entry = o + "entry";
        public static readonly XName extrusion = o + "extrusion";
        public static readonly XName extrusionok = o + "extrusionok";
        public static readonly XName FieldCodes = o + "FieldCodes";
        public static readonly XName fill = o + "fill";
        public static readonly XName forcedash = o + "forcedash";
        public static readonly XName gfxdata = o + "gfxdata";
        public static readonly XName hr = o + "hr";
        public static readonly XName hralign = o + "hralign";
        public static readonly XName href = o + "href";
        public static readonly XName hrnoshade = o + "hrnoshade";
        public static readonly XName hrpct = o + "hrpct";
        public static readonly XName hrstd = o + "hrstd";
        public static readonly XName idmap = o + "idmap";
        public static readonly XName ink = o + "ink";
        public static readonly XName insetmode = o + "insetmode";
        public static readonly XName left = o + "left";
        public static readonly XName LinkType = o + "LinkType";
        public static readonly XName _lock = o + "lock";
        public static readonly XName LockedField = o + "LockedField";
        public static readonly XName master = o + "master";
        public static readonly XName ole = o + "ole";
        public static readonly XName oleicon = o + "oleicon";
        public static readonly XName OLEObject = o + "OLEObject";
        public static readonly XName oned = o + "oned";
        public static readonly XName opacity2 = o + "opacity2";
        public static readonly XName preferrelative = o + "preferrelative";
        public static readonly XName proxy = o + "proxy";
        public static readonly XName r = o + "r";
        public static readonly XName regroupid = o + "regroupid";
        public static readonly XName regrouptable = o + "regrouptable";
        public static readonly XName rel = o + "rel";
        public static readonly XName relationtable = o + "relationtable";
        public static readonly XName relid = o + "relid";
        public static readonly XName right = o + "right";
        public static readonly XName rules = o + "rules";
        public static readonly XName shapedefaults = o + "shapedefaults";
        public static readonly XName shapelayout = o + "shapelayout";
        public static readonly XName signatureline = o + "signatureline";
        public static readonly XName singleclick = o + "singleclick";
        public static readonly XName skew = o + "skew";
        public static readonly XName spid = o + "spid";
        public static readonly XName spt = o + "spt";
        public static readonly XName suggestedsigner = o + "suggestedsigner";
        public static readonly XName suggestedsigner2 = o + "suggestedsigner2";
        public static readonly XName suggestedsigneremail = o + "suggestedsigneremail";
        public static readonly XName tablelimits = o + "tablelimits";
        public static readonly XName tableproperties = o + "tableproperties";
        public static readonly XName targetscreensize = o + "targetscreensize";
        public static readonly XName title = o + "title";
        public static readonly XName top = o + "top";
        public static readonly XName userdrawn = o + "userdrawn";
        public static readonly XName userhidden = o + "userhidden";
        public static readonly XName v = o + "v";
    }

    public static class P
    {
        public static readonly XNamespace p =
            "http://schemas.openxmlformats.org/presentationml/2006/main";
        public static readonly XName anim = p + "anim";
        public static readonly XName animClr = p + "animClr";
        public static readonly XName animEffect = p + "animEffect";
        public static readonly XName animMotion = p + "animMotion";
        public static readonly XName animRot = p + "animRot";
        public static readonly XName animScale = p + "animScale";
        public static readonly XName attrName = p + "attrName";
        public static readonly XName attrNameLst = p + "attrNameLst";
        public static readonly XName audio = p + "audio";
        public static readonly XName bg = p + "bg";
        public static readonly XName bgPr = p + "bgPr";
        public static readonly XName bgRef = p + "bgRef";
        public static readonly XName bldAsOne = p + "bldAsOne";
        public static readonly XName bldDgm = p + "bldDgm";
        public static readonly XName bldGraphic = p + "bldGraphic";
        public static readonly XName bldLst = p + "bldLst";
        public static readonly XName bldOleChart = p + "bldOleChart";
        public static readonly XName bldP = p + "bldP";
        public static readonly XName bldSub = p + "bldSub";
        public static readonly XName blinds = p + "blinds";
        public static readonly XName blipFill = p + "blipFill";
        public static readonly XName bodyStyle = p + "bodyStyle";
        public static readonly XName bold = p + "bold";
        public static readonly XName boldItalic = p + "boldItalic";
        public static readonly XName boolVal = p + "boolVal";
        public static readonly XName by = p + "by";
        public static readonly XName cBhvr = p + "cBhvr";
        public static readonly XName charRg = p + "charRg";
        public static readonly XName checker = p + "checker";
        public static readonly XName childTnLst = p + "childTnLst";
        public static readonly XName circle = p + "circle";
        public static readonly XName clrMap = p + "clrMap";
        public static readonly XName clrMapOvr = p + "clrMapOvr";
        public static readonly XName clrVal = p + "clrVal";
        public static readonly XName cm = p + "cm";
        public static readonly XName cmAuthor = p + "cmAuthor";
        public static readonly XName cmAuthorLst = p + "cmAuthorLst";
        public static readonly XName cmd = p + "cmd";
        public static readonly XName cMediaNode = p + "cMediaNode";
        public static readonly XName cmLst = p + "cmLst";
        public static readonly XName cNvCxnSpPr = p + "cNvCxnSpPr";
        public static readonly XName cNvGraphicFramePr = p + "cNvGraphicFramePr";
        public static readonly XName cNvGrpSpPr = p + "cNvGrpSpPr";
        public static readonly XName cNvPicPr = p + "cNvPicPr";
        public static readonly XName cNvPr = p + "cNvPr";
        public static readonly XName cNvSpPr = p + "cNvSpPr";
        public static readonly XName comb = p + "comb";
        public static readonly XName cond = p + "cond";
        public static readonly XName contentPart = p + "contentPart";
        public static readonly XName control = p + "control";
        public static readonly XName controls = p + "controls";
        public static readonly XName cover = p + "cover";
        public static readonly XName cSld = p + "cSld";
        public static readonly XName cSldViewPr = p + "cSldViewPr";
        public static readonly XName cTn = p + "cTn";
        public static readonly XName custData = p + "custData";
        public static readonly XName custDataLst = p + "custDataLst";
        public static readonly XName custShow = p + "custShow";
        public static readonly XName custShowLst = p + "custShowLst";
        public static readonly XName cut = p + "cut";
        public static readonly XName cViewPr = p + "cViewPr";
        public static readonly XName cxnSp = p + "cxnSp";
        public static readonly XName defaultTextStyle = p + "defaultTextStyle";
        public static readonly XName diamond = p + "diamond";
        public static readonly XName dissolve = p + "dissolve";
        public static readonly XName embed = p + "embed";
        public static readonly XName embeddedFont = p + "embeddedFont";
        public static readonly XName embeddedFontLst = p + "embeddedFontLst";
        public static readonly XName endCondLst = p + "endCondLst";
        public static readonly XName endSnd = p + "endSnd";
        public static readonly XName endSync = p + "endSync";
        public static readonly XName ext = p + "ext";
        public static readonly XName externalData = p + "externalData";
        public static readonly XName extLst = p + "extLst";
        public static readonly XName fade = p + "fade";
        public static readonly XName fltVal = p + "fltVal";
        public static readonly XName font = p + "font";
        public static readonly XName from = p + "from";
        public static readonly XName graphicEl = p + "graphicEl";
        public static readonly XName graphicFrame = p + "graphicFrame";
        public static readonly XName gridSpacing = p + "gridSpacing";
        public static readonly XName grpSp = p + "grpSp";
        public static readonly XName grpSpPr = p + "grpSpPr";
        public static readonly XName guide = p + "guide";
        public static readonly XName guideLst = p + "guideLst";
        public static readonly XName handoutMaster = p + "handoutMaster";
        public static readonly XName handoutMasterId = p + "handoutMasterId";
        public static readonly XName handoutMasterIdLst = p + "handoutMasterIdLst";
        public static readonly XName hf = p + "hf";
        public static readonly XName hsl = p + "hsl";
        public static readonly XName inkTgt = p + "inkTgt";
        public static readonly XName italic = p + "italic";
        public static readonly XName iterate = p + "iterate";
        public static readonly XName kinsoku = p + "kinsoku";
        public static readonly XName link = p + "link";
        public static readonly XName modifyVerifier = p + "modifyVerifier";
        public static readonly XName newsflash = p + "newsflash";
        public static readonly XName nextCondLst = p + "nextCondLst";
        public static readonly XName normalViewPr = p + "normalViewPr";
        public static readonly XName notes = p + "notes";
        public static readonly XName notesMaster = p + "notesMaster";
        public static readonly XName notesMasterId = p + "notesMasterId";
        public static readonly XName notesMasterIdLst = p + "notesMasterIdLst";
        public static readonly XName notesStyle = p + "notesStyle";
        public static readonly XName notesSz = p + "notesSz";
        public static readonly XName notesTextViewPr = p + "notesTextViewPr";
        public static readonly XName notesViewPr = p + "notesViewPr";
        public static readonly XName nvCxnSpPr = p + "nvCxnSpPr";
        public static readonly XName nvGraphicFramePr = p + "nvGraphicFramePr";
        public static readonly XName nvGrpSpPr = p + "nvGrpSpPr";
        public static readonly XName nvPicPr = p + "nvPicPr";
        public static readonly XName nvPr = p + "nvPr";
        public static readonly XName nvSpPr = p + "nvSpPr";
        public static readonly XName oleChartEl = p + "oleChartEl";
        public static readonly XName oleObj = p + "oleObj";
        public static readonly XName origin = p + "origin";
        public static readonly XName otherStyle = p + "otherStyle";
        public static readonly XName outlineViewPr = p + "outlineViewPr";
        public static readonly XName par = p + "par";
        public static readonly XName ph = p + "ph";
        public static readonly XName photoAlbum = p + "photoAlbum";
        public static readonly XName pic = p + "pic";
        public static readonly XName plus = p + "plus";
        public static readonly XName pos = p + "pos";
        public static readonly XName presentation = p + "presentation";
        public static readonly XName prevCondLst = p + "prevCondLst";
        public static readonly XName pRg = p + "pRg";
        public static readonly XName pull = p + "pull";
        public static readonly XName push = p + "push";
        public static readonly XName random = p + "random";
        public static readonly XName randomBar = p + "randomBar";
        public static readonly XName rCtr = p + "rCtr";
        public static readonly XName regular = p + "regular";
        public static readonly XName restoredLeft = p + "restoredLeft";
        public static readonly XName restoredTop = p + "restoredTop";
        public static readonly XName rgb = p + "rgb";
        public static readonly XName rtn = p + "rtn";
        public static readonly XName scale = p + "scale";
        public static readonly XName seq = p + "seq";
        public static readonly XName set = p + "set";
        public static readonly XName sld = p + "sld";
        public static readonly XName sldId = p + "sldId";
        public static readonly XName sldIdLst = p + "sldIdLst";
        public static readonly XName sldLayout = p + "sldLayout";
        public static readonly XName sldLayoutId = p + "sldLayoutId";
        public static readonly XName sldLayoutIdLst = p + "sldLayoutIdLst";
        public static readonly XName sldLst = p + "sldLst";
        public static readonly XName sldMaster = p + "sldMaster";
        public static readonly XName sldMasterId = p + "sldMasterId";
        public static readonly XName sldMasterIdLst = p + "sldMasterIdLst";
        public static readonly XName sldSyncPr = p + "sldSyncPr";
        public static readonly XName sldSz = p + "sldSz";
        public static readonly XName sldTgt = p + "sldTgt";
        public static readonly XName slideViewPr = p + "slideViewPr";
        public static readonly XName snd = p + "snd";
        public static readonly XName sndAc = p + "sndAc";
        public static readonly XName sndTgt = p + "sndTgt";
        public static readonly XName sorterViewPr = p + "sorterViewPr";
        public static readonly XName sp = p + "sp";
        public static readonly XName split = p + "split";
        public static readonly XName spPr = p + "spPr";
        public static readonly XName spTgt = p + "spTgt";
        public static readonly XName spTree = p + "spTree";
        public static readonly XName stCondLst = p + "stCondLst";
        public static readonly XName strips = p + "strips";
        public static readonly XName strVal = p + "strVal";
        public static readonly XName stSnd = p + "stSnd";
        public static readonly XName style = p + "style";
        public static readonly XName subSp = p + "subSp";
        public static readonly XName subTnLst = p + "subTnLst";
        public static readonly XName tag = p + "tag";
        public static readonly XName tagLst = p + "tagLst";
        public static readonly XName tags = p + "tags";
        public static readonly XName tav = p + "tav";
        public static readonly XName tavLst = p + "tavLst";
        public static readonly XName text = p + "text";
        public static readonly XName tgtEl = p + "tgtEl";
        public static readonly XName timing = p + "timing";
        public static readonly XName titleStyle = p + "titleStyle";
        public static readonly XName tmAbs = p + "tmAbs";
        public static readonly XName tmPct = p + "tmPct";
        public static readonly XName tmpl = p + "tmpl";
        public static readonly XName tmplLst = p + "tmplLst";
        public static readonly XName tn = p + "tn";
        public static readonly XName tnLst = p + "tnLst";
        public static readonly XName to = p + "to";
        public static readonly XName transition = p + "transition";
        public static readonly XName txBody = p + "txBody";
        public static readonly XName txEl = p + "txEl";
        public static readonly XName txStyles = p + "txStyles";
        public static readonly XName val = p + "val";
        public static readonly XName video = p + "video";
        public static readonly XName viewPr = p + "viewPr";
        public static readonly XName wedge = p + "wedge";
        public static readonly XName wheel = p + "wheel";
        public static readonly XName wipe = p + "wipe";
        public static readonly XName xfrm = p + "xfrm";
        public static readonly XName zoom = p + "zoom";
    }

    public static class P14
    {
        public static readonly XNamespace p14 =
            "http://schemas.microsoft.com/office/powerpoint/2010/main";
        public static readonly XName bmk = p14 + "bmk";
        public static readonly XName bmkLst = p14 + "bmkLst";
        public static readonly XName bmkTgt = p14 + "bmkTgt";
        public static readonly XName bounceEnd = p14 + "bounceEnd";
        public static readonly XName bwMode = p14 + "bwMode";
        public static readonly XName cNvContentPartPr = p14 + "cNvContentPartPr";
        public static readonly XName cNvPr = p14 + "cNvPr";
        public static readonly XName conveyor = p14 + "conveyor";
        public static readonly XName creationId = p14 + "creationId";
        public static readonly XName doors = p14 + "doors";
        public static readonly XName dur = p14 + "dur";
        public static readonly XName extLst = p14 + "extLst";
        public static readonly XName fade = p14 + "fade";
        public static readonly XName ferris = p14 + "ferris";
        public static readonly XName flash = p14 + "flash";
        public static readonly XName flip = p14 + "flip";
        public static readonly XName flythrough = p14 + "flythrough";
        public static readonly XName gallery = p14 + "gallery";
        public static readonly XName glitter = p14 + "glitter";
        public static readonly XName honeycomb = p14 + "honeycomb";
        public static readonly XName laserTraceLst = p14 + "laserTraceLst";
        public static readonly XName media = p14 + "media";
        public static readonly XName modId = p14 + "modId";
        public static readonly XName nvContentPartPr = p14 + "nvContentPartPr";
        public static readonly XName nvPr = p14 + "nvPr";
        public static readonly XName pan = p14 + "pan";
        public static readonly XName pauseEvt = p14 + "pauseEvt";
        public static readonly XName playEvt = p14 + "playEvt";
        public static readonly XName presetBounceEnd = p14 + "presetBounceEnd";
        public static readonly XName prism = p14 + "prism";
        public static readonly XName resumeEvt = p14 + "resumeEvt";
        public static readonly XName reveal = p14 + "reveal";
        public static readonly XName ripple = p14 + "ripple";
        public static readonly XName section = p14 + "section";
        public static readonly XName sectionLst = p14 + "sectionLst";
        public static readonly XName seekEvt = p14 + "seekEvt";
        public static readonly XName showEvtLst = p14 + "showEvtLst";
        public static readonly XName shred = p14 + "shred";
        public static readonly XName sldId = p14 + "sldId";
        public static readonly XName sldIdLst = p14 + "sldIdLst";
        public static readonly XName stopEvt = p14 + "stopEvt";
        public static readonly XName _switch = p14 + "switch";
        public static readonly XName tracePt = p14 + "tracePt";
        public static readonly XName tracePtLst = p14 + "tracePtLst";
        public static readonly XName triggerEvt = p14 + "triggerEvt";
        public static readonly XName trim = p14 + "trim";
        public static readonly XName vortex = p14 + "vortex";
        public static readonly XName warp = p14 + "warp";
        public static readonly XName wheelReverse = p14 + "wheelReverse";
        public static readonly XName window = p14 + "window";
        public static readonly XName xfrm = p14 + "xfrm";
    }

    public static class P15
    {
        public static readonly XNamespace p15 =
            "http://schemas.microsoft.com/office15/powerpoint";
        public static readonly XName extElement = p15 + "extElement";
    }

    public static class PAV
    {
        public static readonly XNamespace pav = "http://schemas.microsoft.com/office/2007/6/19/audiovideo";
        public static readonly XName media = pav + "media";
        public static readonly XName srcMedia = pav + "srcMedia";
        public static readonly XName bmkLst = pav + "bmkLst";
    }

    public static class Pic
    {
        public static readonly XNamespace pic =
            "http://schemas.openxmlformats.org/drawingml/2006/picture";
        public static readonly XName blipFill = pic + "blipFill";
        public static readonly XName cNvPicPr = pic + "cNvPicPr";
        public static readonly XName cNvPr = pic + "cNvPr";
        public static readonly XName nvPicPr = pic + "nvPicPr";
        public static readonly XName _pic = pic + "pic";
        public static readonly XName spPr = pic + "spPr";
    }

    public static class Plegacy
    {
        public static readonly XNamespace plegacy = "urn:schemas-microsoft-com:office:powerpoint";
        public static readonly XName textdata = plegacy + "textdata";
    }

    public static class R
    {
        public static readonly XNamespace r =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        public static readonly XName blip = r + "blip";
        public static readonly XName cs = r + "cs";
        public static readonly XName dm = r + "dm";
        public static readonly XName embed = r + "embed";
        public static readonly XName href = r + "href";
        public static readonly XName id = r + "id";
        public static readonly XName link = r + "link";
        public static readonly XName lo = r + "lo";
        public static readonly XName pict = r + "pict";
        public static readonly XName qs = r + "qs";
        public static readonly XName verticalDpi = r + "verticalDpi";
    }

    public static class S
    {
        public static readonly XNamespace s =
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        public static readonly XName alignment = s + "alignment";
        public static readonly XName anchor = s + "anchor";
        public static readonly XName author = s + "author";
        public static readonly XName authors = s + "authors";
        public static readonly XName autoFilter = s + "autoFilter";
        public static readonly XName autoSortScope = s + "autoSortScope";
        public static readonly XName b = s + "b";
        public static readonly XName bgColor = s + "bgColor";
        public static readonly XName bk = s + "bk";
        public static readonly XName border = s + "border";
        public static readonly XName borders = s + "borders";
        public static readonly XName bottom = s + "bottom";
        public static readonly XName brk = s + "brk";
        public static readonly XName c = s + "c";
        public static readonly XName cacheField = s + "cacheField";
        public static readonly XName cacheFields = s + "cacheFields";
        public static readonly XName cacheHierarchies = s + "cacheHierarchies";
        public static readonly XName cacheHierarchy = s + "cacheHierarchy";
        public static readonly XName cacheSource = s + "cacheSource";
        public static readonly XName calcChain = s + "calcChain";
        public static readonly XName calcPr = s + "calcPr";
        public static readonly XName calculatedColumnFormula = s + "calculatedColumnFormula";
        public static readonly XName calculatedItem = s + "calculatedItem";
        public static readonly XName calculatedItems = s + "calculatedItems";
        public static readonly XName calculatedMember = s + "calculatedMember";
        public static readonly XName calculatedMembers = s + "calculatedMembers";
        public static readonly XName cell = s + "cell";
        public static readonly XName cellMetadata = s + "cellMetadata";
        public static readonly XName cellSmartTag = s + "cellSmartTag";
        public static readonly XName cellSmartTagPr = s + "cellSmartTagPr";
        public static readonly XName cellSmartTags = s + "cellSmartTags";
        public static readonly XName cellStyle = s + "cellStyle";
        public static readonly XName cellStyles = s + "cellStyles";
        public static readonly XName cellStyleXfs = s + "cellStyleXfs";
        public static readonly XName cellWatch = s + "cellWatch";
        public static readonly XName cellWatches = s + "cellWatches";
        public static readonly XName cellXfs = s + "cellXfs";
        public static readonly XName cfRule = s + "cfRule";
        public static readonly XName cfvo = s + "cfvo";
        public static readonly XName charset = s + "charset";
        public static readonly XName chartFormat = s + "chartFormat";
        public static readonly XName chartFormats = s + "chartFormats";
        public static readonly XName chartsheet = s + "chartsheet";
        public static readonly XName col = s + "col";
        public static readonly XName colBreaks = s + "colBreaks";
        public static readonly XName colFields = s + "colFields";
        public static readonly XName colHierarchiesUsage = s + "colHierarchiesUsage";
        public static readonly XName colHierarchyUsage = s + "colHierarchyUsage";
        public static readonly XName colItems = s + "colItems";
        public static readonly XName color = s + "color";
        public static readonly XName colorFilter = s + "colorFilter";
        public static readonly XName colors = s + "colors";
        public static readonly XName colorScale = s + "colorScale";
        public static readonly XName cols = s + "cols";
        public static readonly XName comment = s + "comment";
        public static readonly XName commentList = s + "commentList";
        public static readonly XName comments = s + "comments";
        public static readonly XName condense = s + "condense";
        public static readonly XName conditionalFormat = s + "conditionalFormat";
        public static readonly XName conditionalFormats = s + "conditionalFormats";
        public static readonly XName conditionalFormatting = s + "conditionalFormatting";
        public static readonly XName connection = s + "connection";
        public static readonly XName connections = s + "connections";
        public static readonly XName consolidation = s + "consolidation";
        public static readonly XName control = s + "control";
        public static readonly XName controlPr = s + "controlPr";
        public static readonly XName controls = s + "controls";
        public static readonly XName customFilter = s + "customFilter";
        public static readonly XName customFilters = s + "customFilters";
        public static readonly XName customPr = s + "customPr";
        public static readonly XName customProperties = s + "customProperties";
        public static readonly XName customSheetView = s + "customSheetView";
        public static readonly XName customSheetViews = s + "customSheetViews";
        public static readonly XName d = s + "d";
        public static readonly XName dataBar = s + "dataBar";
        public static readonly XName dataConsolidate = s + "dataConsolidate";
        public static readonly XName dataField = s + "dataField";
        public static readonly XName dataFields = s + "dataFields";
        public static readonly XName dataRef = s + "dataRef";
        public static readonly XName dataRefs = s + "dataRefs";
        public static readonly XName dataValidation = s + "dataValidation";
        public static readonly XName dataValidations = s + "dataValidations";
        public static readonly XName dateGroupItem = s + "dateGroupItem";
        public static readonly XName dbPr = s + "dbPr";
        public static readonly XName ddeItem = s + "ddeItem";
        public static readonly XName ddeItems = s + "ddeItems";
        public static readonly XName ddeLink = s + "ddeLink";
        public static readonly XName definedName = s + "definedName";
        public static readonly XName definedNames = s + "definedNames";
        public static readonly XName deletedField = s + "deletedField";
        public static readonly XName diagonal = s + "diagonal";
        public static readonly XName dialogsheet = s + "dialogsheet";
        public static readonly XName dimension = s + "dimension";
        public static readonly XName dimensions = s + "dimensions";
        public static readonly XName discretePr = s + "discretePr";
        public static readonly XName drawing = s + "drawing";
        public static readonly XName dxf = s + "dxf";
        public static readonly XName dxfs = s + "dxfs";
        public static readonly XName dynamicFilter = s + "dynamicFilter";
        public static readonly XName e = s + "e";
        public static readonly XName entries = s + "entries";
        public static readonly XName evenFooter = s + "evenFooter";
        public static readonly XName evenHeader = s + "evenHeader";
        public static readonly XName ext = s + "ext";
        public static readonly XName extend = s + "extend";
        public static readonly XName externalBook = s + "externalBook";
        public static readonly XName externalLink = s + "externalLink";
        public static readonly XName extLst = s + "extLst";
        public static readonly XName f = s + "f";
        public static readonly XName family = s + "family";
        public static readonly XName fgColor = s + "fgColor";
        public static readonly XName field = s + "field";
        public static readonly XName fieldGroup = s + "fieldGroup";
        public static readonly XName fieldsUsage = s + "fieldsUsage";
        public static readonly XName fieldUsage = s + "fieldUsage";
        public static readonly XName fill = s + "fill";
        public static readonly XName fills = s + "fills";
        public static readonly XName filter = s + "filter";
        public static readonly XName filterColumn = s + "filterColumn";
        public static readonly XName filters = s + "filters";
        public static readonly XName firstFooter = s + "firstFooter";
        public static readonly XName firstHeader = s + "firstHeader";
        public static readonly XName font = s + "font";
        public static readonly XName fonts = s + "fonts";
        public static readonly XName foo = s + "foo";
        public static readonly XName format = s + "format";
        public static readonly XName formats = s + "formats";
        public static readonly XName formula = s + "formula";
        public static readonly XName formula1 = s + "formula1";
        public static readonly XName formula2 = s + "formula2";
        public static readonly XName from = s + "from";
        public static readonly XName futureMetadata = s + "futureMetadata";
        public static readonly XName gradientFill = s + "gradientFill";
        public static readonly XName group = s + "group";
        public static readonly XName groupItems = s + "groupItems";
        public static readonly XName groupLevel = s + "groupLevel";
        public static readonly XName groupLevels = s + "groupLevels";
        public static readonly XName groupMember = s + "groupMember";
        public static readonly XName groupMembers = s + "groupMembers";
        public static readonly XName groups = s + "groups";
        public static readonly XName header = s + "header";
        public static readonly XName headerFooter = s + "headerFooter";
        public static readonly XName headers = s + "headers";
        public static readonly XName horizontal = s + "horizontal";
        public static readonly XName hyperlink = s + "hyperlink";
        public static readonly XName hyperlinks = s + "hyperlinks";
        public static readonly XName i = s + "i";
        public static readonly XName iconFilter = s + "iconFilter";
        public static readonly XName iconSet = s + "iconSet";
        public static readonly XName ignoredError = s + "ignoredError";
        public static readonly XName ignoredErrors = s + "ignoredErrors";
        public static readonly XName indexedColors = s + "indexedColors";
        public static readonly XName inputCells = s + "inputCells";
        public static readonly XName _is = s + "is";
        public static readonly XName item = s + "item";
        public static readonly XName items = s + "items";
        public static readonly XName k = s + "k";
        public static readonly XName kpi = s + "kpi";
        public static readonly XName kpis = s + "kpis";
        public static readonly XName left = s + "left";
        public static readonly XName legacyDrawing = s + "legacyDrawing";
        public static readonly XName legacyDrawingHF = s + "legacyDrawingHF";
        public static readonly XName location = s + "location";
        public static readonly XName m = s + "m";
        public static readonly XName main = s + "main";
        public static readonly XName map = s + "map";
        public static readonly XName maps = s + "maps";
        public static readonly XName mdx = s + "mdx";
        public static readonly XName mdxMetadata = s + "mdxMetadata";
        public static readonly XName measureGroup = s + "measureGroup";
        public static readonly XName measureGroups = s + "measureGroups";
        public static readonly XName member = s + "member";
        public static readonly XName members = s + "members";
        public static readonly XName mergeCell = s + "mergeCell";
        public static readonly XName mergeCells = s + "mergeCells";
        public static readonly XName metadata = s + "metadata";
        public static readonly XName metadataStrings = s + "metadataStrings";
        public static readonly XName metadataType = s + "metadataType";
        public static readonly XName metadataTypes = s + "metadataTypes";
        public static readonly XName mp = s + "mp";
        public static readonly XName mpMap = s + "mpMap";
        public static readonly XName mps = s + "mps";
        public static readonly XName mruColors = s + "mruColors";
        public static readonly XName ms = s + "ms";
        public static readonly XName n = s + "n";
        public static readonly XName name = s + "name";
        public static readonly XName nc = s + "nc";
        public static readonly XName ndxf = s + "ndxf";
        public static readonly XName numFmt = s + "numFmt";
        public static readonly XName numFmts = s + "numFmts";
        public static readonly XName objectPr = s + "objectPr";
        public static readonly XName oc = s + "oc";
        public static readonly XName oddFooter = s + "oddFooter";
        public static readonly XName oddHeader = s + "oddHeader";
        public static readonly XName odxf = s + "odxf";
        public static readonly XName olapPr = s + "olapPr";
        public static readonly XName oldFormula = s + "oldFormula";
        public static readonly XName oleItem = s + "oleItem";
        public static readonly XName oleItems = s + "oleItems";
        public static readonly XName oleLink = s + "oleLink";
        public static readonly XName oleObject = s + "oleObject";
        public static readonly XName oleObjects = s + "oleObjects";
        public static readonly XName outline = s + "outline";
        public static readonly XName outlinePr = s + "outlinePr";
        public static readonly XName p = s + "p";
        public static readonly XName page = s + "page";
        public static readonly XName pageField = s + "pageField";
        public static readonly XName pageFields = s + "pageFields";
        public static readonly XName pageItem = s + "pageItem";
        public static readonly XName pageMargins = s + "pageMargins";
        public static readonly XName pages = s + "pages";
        public static readonly XName pageSetup = s + "pageSetup";
        public static readonly XName pageSetUpPr = s + "pageSetUpPr";
        public static readonly XName pane = s + "pane";
        public static readonly XName parameter = s + "parameter";
        public static readonly XName parameters = s + "parameters";
        public static readonly XName patternFill = s + "patternFill";
        public static readonly XName phoneticPr = s + "phoneticPr";
        public static readonly XName picture = s + "picture";
        public static readonly XName pivotArea = s + "pivotArea";
        public static readonly XName pivotAreas = s + "pivotAreas";
        public static readonly XName pivotCache = s + "pivotCache";
        public static readonly XName pivotCacheDefinition = s + "pivotCacheDefinition";
        public static readonly XName pivotCacheRecords = s + "pivotCacheRecords";
        public static readonly XName pivotCaches = s + "pivotCaches";
        public static readonly XName pivotField = s + "pivotField";
        public static readonly XName pivotFields = s + "pivotFields";
        public static readonly XName pivotHierarchies = s + "pivotHierarchies";
        public static readonly XName pivotHierarchy = s + "pivotHierarchy";
        public static readonly XName pivotSelection = s + "pivotSelection";
        public static readonly XName pivotTableDefinition = s + "pivotTableDefinition";
        public static readonly XName pivotTableStyleInfo = s + "pivotTableStyleInfo";
        public static readonly XName printOptions = s + "printOptions";
        public static readonly XName protectedRange = s + "protectedRange";
        public static readonly XName protectedRanges = s + "protectedRanges";
        public static readonly XName protection = s + "protection";
        public static readonly XName query = s + "query";
        public static readonly XName queryCache = s + "queryCache";
        public static readonly XName queryTable = s + "queryTable";
        public static readonly XName queryTableDeletedFields = s + "queryTableDeletedFields";
        public static readonly XName queryTableField = s + "queryTableField";
        public static readonly XName queryTableFields = s + "queryTableFields";
        public static readonly XName queryTableRefresh = s + "queryTableRefresh";
        public static readonly XName r = s + "r";
        public static readonly XName raf = s + "raf";
        public static readonly XName rangePr = s + "rangePr";
        public static readonly XName rangeSet = s + "rangeSet";
        public static readonly XName rangeSets = s + "rangeSets";
        public static readonly XName rc = s + "rc";
        public static readonly XName rcc = s + "rcc";
        public static readonly XName rcft = s + "rcft";
        public static readonly XName rcmt = s + "rcmt";
        public static readonly XName rcv = s + "rcv";
        public static readonly XName rdn = s + "rdn";
        public static readonly XName reference = s + "reference";
        public static readonly XName references = s + "references";
        public static readonly XName reviewed = s + "reviewed";
        public static readonly XName reviewedList = s + "reviewedList";
        public static readonly XName revisions = s + "revisions";
        public static readonly XName rfmt = s + "rfmt";
        public static readonly XName rFont = s + "rFont";
        public static readonly XName rgbColor = s + "rgbColor";
        public static readonly XName right = s + "right";
        public static readonly XName ris = s + "ris";
        public static readonly XName rm = s + "rm";
        public static readonly XName row = s + "row";
        public static readonly XName rowBreaks = s + "rowBreaks";
        public static readonly XName rowFields = s + "rowFields";
        public static readonly XName rowHierarchiesUsage = s + "rowHierarchiesUsage";
        public static readonly XName rowHierarchyUsage = s + "rowHierarchyUsage";
        public static readonly XName rowItems = s + "rowItems";
        public static readonly XName rPh = s + "rPh";
        public static readonly XName rPr = s + "rPr";
        public static readonly XName rqt = s + "rqt";
        public static readonly XName rrc = s + "rrc";
        public static readonly XName rsnm = s + "rsnm";
        public static readonly XName _s = s + "s";
        public static readonly XName scenario = s + "scenario";
        public static readonly XName scenarios = s + "scenarios";
        public static readonly XName scheme = s + "scheme";
        public static readonly XName selection = s + "selection";
        public static readonly XName serverFormat = s + "serverFormat";
        public static readonly XName serverFormats = s + "serverFormats";
        public static readonly XName set = s + "set";
        public static readonly XName sets = s + "sets";
        public static readonly XName shadow = s + "shadow";
        public static readonly XName sharedItems = s + "sharedItems";
        public static readonly XName sheet = s + "sheet";
        public static readonly XName sheetCalcPr = s + "sheetCalcPr";
        public static readonly XName sheetData = s + "sheetData";
        public static readonly XName sheetDataSet = s + "sheetDataSet";
        public static readonly XName sheetFormatPr = s + "sheetFormatPr";
        public static readonly XName sheetId = s + "sheetId";
        public static readonly XName sheetIdMap = s + "sheetIdMap";
        public static readonly XName sheetName = s + "sheetName";
        public static readonly XName sheetNames = s + "sheetNames";
        public static readonly XName sheetPr = s + "sheetPr";
        public static readonly XName sheetProtection = s + "sheetProtection";
        public static readonly XName sheets = s + "sheets";
        public static readonly XName sheetView = s + "sheetView";
        public static readonly XName sheetViews = s + "sheetViews";
        public static readonly XName si = s + "si";
        public static readonly XName singleXmlCell = s + "singleXmlCell";
        public static readonly XName singleXmlCells = s + "singleXmlCells";
        public static readonly XName smartTags = s + "smartTags";
        public static readonly XName sortByTuple = s + "sortByTuple";
        public static readonly XName sortCondition = s + "sortCondition";
        public static readonly XName sortState = s + "sortState";
        public static readonly XName sst = s + "sst";
        public static readonly XName stop = s + "stop";
        public static readonly XName stp = s + "stp";
        public static readonly XName strike = s + "strike";
        public static readonly XName styleSheet = s + "styleSheet";
        public static readonly XName sz = s + "sz";
        public static readonly XName t = s + "t";
        public static readonly XName tabColor = s + "tabColor";
        public static readonly XName table = s + "table";
        public static readonly XName tableColumn = s + "tableColumn";
        public static readonly XName tableColumns = s + "tableColumns";
        public static readonly XName tablePart = s + "tablePart";
        public static readonly XName tableParts = s + "tableParts";
        public static readonly XName tables = s + "tables";
        public static readonly XName tableStyle = s + "tableStyle";
        public static readonly XName tableStyleElement = s + "tableStyleElement";
        public static readonly XName tableStyleInfo = s + "tableStyleInfo";
        public static readonly XName tableStyles = s + "tableStyles";
        public static readonly XName text = s + "text";
        public static readonly XName textField = s + "textField";
        public static readonly XName textFields = s + "textFields";
        public static readonly XName textPr = s + "textPr";
        public static readonly XName to = s + "to";
        public static readonly XName top = s + "top";
        public static readonly XName top10 = s + "top10";
        public static readonly XName totalsRowFormula = s + "totalsRowFormula";
        public static readonly XName tp = s + "tp";
        public static readonly XName tpl = s + "tpl";
        public static readonly XName tpls = s + "tpls";
        public static readonly XName tr = s + "tr";
        public static readonly XName tupleCache = s + "tupleCache";
        public static readonly XName u = s + "u";
        public static readonly XName undo = s + "undo";
        public static readonly XName userInfo = s + "userInfo";
        public static readonly XName users = s + "users";
        public static readonly XName v = s + "v";
        public static readonly XName val = s + "val";
        public static readonly XName value = s + "value";
        public static readonly XName valueMetadata = s + "valueMetadata";
        public static readonly XName values = s + "values";
        public static readonly XName vertAlign = s + "vertAlign";
        public static readonly XName vertical = s + "vertical";
        public static readonly XName volType = s + "volType";
        public static readonly XName volTypes = s + "volTypes";
        public static readonly XName webPr = s + "webPr";
        public static readonly XName webPublishItem = s + "webPublishItem";
        public static readonly XName webPublishItems = s + "webPublishItems";
        public static readonly XName worksheet = s + "worksheet";
        public static readonly XName worksheetEx14 = s + "worksheetEx14";
        public static readonly XName worksheetSource = s + "worksheetSource";
        public static readonly XName x = s + "x";
        public static readonly XName xf = s + "xf";
        public static readonly XName xmlCellPr = s + "xmlCellPr";
        public static readonly XName xmlColumnPr = s + "xmlColumnPr";
        public static readonly XName xmlPr = s + "xmlPr";
    }

    public static class SL
    {
        public static readonly XNamespace sl =
            "http://schemas.openxmlformats.org/schemaLibrary/2006/main";
        public static readonly XName manifestLocation = sl + "manifestLocation";
        public static readonly XName schema = sl + "schema";
        public static readonly XName schemaLibrary = sl + "schemaLibrary";
        public static readonly XName uri = sl + "uri";
    }

    public static class SLE
    {
        public static readonly XNamespace sle =
            "http://schemas.microsoft.com/office/drawing/2010/slicer";
        public static readonly XName slicer = sle + "slicer";
    }

    public static class VML
    {
        public static readonly XNamespace vml =
            "urn:schemas-microsoft-com:vml";
        public static readonly XName arc = vml + "arc";
        public static readonly XName background = vml + "background";
        public static readonly XName curve = vml + "curve";
        public static readonly XName ext = vml + "ext";
        public static readonly XName f = vml + "f";
        public static readonly XName fill = vml + "fill";
        public static readonly XName formulas = vml + "formulas";
        public static readonly XName group = vml + "group";
        public static readonly XName h = vml + "h";
        public static readonly XName handles = vml + "handles";
        public static readonly XName image = vml + "image";
        public static readonly XName imagedata = vml + "imagedata";
        public static readonly XName line = vml + "line";
        public static readonly XName oval = vml + "oval";
        public static readonly XName path = vml + "path";
        public static readonly XName polyline = vml + "polyline";
        public static readonly XName rect = vml + "rect";
        public static readonly XName roundrect = vml + "roundrect";
        public static readonly XName shadow = vml + "shadow";
        public static readonly XName shape = vml + "shape";
        public static readonly XName shapetype = vml + "shapetype";
        public static readonly XName stroke = vml + "stroke";
        public static readonly XName textbox = vml + "textbox";
        public static readonly XName textpath = vml + "textpath";
    }

    public static class VT
    {
        public static readonly XNamespace vt =
            "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
        public static readonly XName _bool = vt + "bool";
        public static readonly XName filetime = vt + "filetime";
        public static readonly XName i4 = vt + "i4";
        public static readonly XName lpstr = vt + "lpstr";
        public static readonly XName lpwstr = vt + "lpwstr";
        public static readonly XName r8 = vt + "r8";
        public static readonly XName variant = vt + "variant";
        public static readonly XName vector = vt + "vector";
    }

    public static class W
    {
        public static readonly XNamespace w =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        public static readonly XName abstractNum = w + "abstractNum";
        public static readonly XName abstractNumId = w + "abstractNumId";
        public static readonly XName accent1 = w + "accent1";
        public static readonly XName accent2 = w + "accent2";
        public static readonly XName accent3 = w + "accent3";
        public static readonly XName accent4 = w + "accent4";
        public static readonly XName accent5 = w + "accent5";
        public static readonly XName accent6 = w + "accent6";
        public static readonly XName activeRecord = w + "activeRecord";
        public static readonly XName activeWritingStyle = w + "activeWritingStyle";
        public static readonly XName actualPg = w + "actualPg";
        public static readonly XName addressFieldName = w + "addressFieldName";
        public static readonly XName adjustLineHeightInTable = w + "adjustLineHeightInTable";
        public static readonly XName adjustRightInd = w + "adjustRightInd";
        public static readonly XName after = w + "after";
        public static readonly XName afterAutospacing = w + "afterAutospacing";
        public static readonly XName afterLines = w + "afterLines";
        public static readonly XName algIdExt = w + "algIdExt";
        public static readonly XName algIdExtSource = w + "algIdExtSource";
        public static readonly XName alias = w + "alias";
        public static readonly XName aliases = w + "aliases";
        public static readonly XName alignBordersAndEdges = w + "alignBordersAndEdges";
        public static readonly XName alignment = w + "alignment";
        public static readonly XName alignTablesRowByRow = w + "alignTablesRowByRow";
        public static readonly XName allowPNG = w + "allowPNG";
        public static readonly XName allowSpaceOfSameStyleInTable = w + "allowSpaceOfSameStyleInTable";
        public static readonly XName altChunk = w + "altChunk";
        public static readonly XName altChunkPr = w + "altChunkPr";
        public static readonly XName altName = w + "altName";
        public static readonly XName alwaysMergeEmptyNamespace = w + "alwaysMergeEmptyNamespace";
        public static readonly XName alwaysShowPlaceholderText = w + "alwaysShowPlaceholderText";
        public static readonly XName anchor = w + "anchor";
        public static readonly XName anchorLock = w + "anchorLock";
        public static readonly XName annotationRef = w + "annotationRef";
        public static readonly XName applyBreakingRules = w + "applyBreakingRules";
        public static readonly XName appName = w + "appName";
        public static readonly XName ascii = w + "ascii";
        public static readonly XName asciiTheme = w + "asciiTheme";
        public static readonly XName attachedSchema = w + "attachedSchema";
        public static readonly XName attachedTemplate = w + "attachedTemplate";
        public static readonly XName attr = w + "attr";
        public static readonly XName author = w + "author";
        public static readonly XName autofitToFirstFixedWidthCell = w + "autofitToFirstFixedWidthCell";
        public static readonly XName autoFormatOverride = w + "autoFormatOverride";
        public static readonly XName autoHyphenation = w + "autoHyphenation";
        public static readonly XName autoRedefine = w + "autoRedefine";
        public static readonly XName autoSpaceDE = w + "autoSpaceDE";
        public static readonly XName autoSpaceDN = w + "autoSpaceDN";
        public static readonly XName autoSpaceLikeWord95 = w + "autoSpaceLikeWord95";
        public static readonly XName b = w + "b";
        public static readonly XName background = w + "background";
        public static readonly XName balanceSingleByteDoubleByteWidth = w + "balanceSingleByteDoubleByteWidth";
        public static readonly XName bar = w + "bar";
        public static readonly XName basedOn = w + "basedOn";
        public static readonly XName bCs = w + "bCs";
        public static readonly XName bdr = w + "bdr";
        public static readonly XName before = w + "before";
        public static readonly XName beforeAutospacing = w + "beforeAutospacing";
        public static readonly XName beforeLines = w + "beforeLines";
        public static readonly XName behavior = w + "behavior";
        public static readonly XName behaviors = w + "behaviors";
        public static readonly XName between = w + "between";
        public static readonly XName bg1 = w + "bg1";
        public static readonly XName bg2 = w + "bg2";
        public static readonly XName bibliography = w + "bibliography";
        public static readonly XName bidi = w + "bidi";
        public static readonly XName bidiVisual = w + "bidiVisual";
        public static readonly XName blockQuote = w + "blockQuote";
        public static readonly XName body = w + "body";
        public static readonly XName bodyDiv = w + "bodyDiv";
        public static readonly XName bookFoldPrinting = w + "bookFoldPrinting";
        public static readonly XName bookFoldPrintingSheets = w + "bookFoldPrintingSheets";
        public static readonly XName bookFoldRevPrinting = w + "bookFoldRevPrinting";
        public static readonly XName bookmarkEnd = w + "bookmarkEnd";
        public static readonly XName bookmarkStart = w + "bookmarkStart";
        public static readonly XName bordersDoNotSurroundFooter = w + "bordersDoNotSurroundFooter";
        public static readonly XName bordersDoNotSurroundHeader = w + "bordersDoNotSurroundHeader";
        public static readonly XName bottom = w + "bottom";
        public static readonly XName bottomFromText = w + "bottomFromText";
        public static readonly XName br = w + "br";
        public static readonly XName cachedColBalance = w + "cachedColBalance";
        public static readonly XName calcOnExit = w + "calcOnExit";
        public static readonly XName calendar = w + "calendar";
        public static readonly XName cantSplit = w + "cantSplit";
        public static readonly XName caps = w + "caps";
        public static readonly XName category = w + "category";
        public static readonly XName cellDel = w + "cellDel";
        public static readonly XName cellIns = w + "cellIns";
        public static readonly XName cellMerge = w + "cellMerge";
        public static readonly XName chapSep = w + "chapSep";
        public static readonly XName chapStyle = w + "chapStyle";
        public static readonly XName _char = w + "char";
        public static readonly XName characterSpacingControl = w + "characterSpacingControl";
        public static readonly XName charset = w + "charset";
        public static readonly XName charSpace = w + "charSpace";
        public static readonly XName checkBox = w + "checkBox";
        public static readonly XName _checked = w + "checked";
        public static readonly XName checkErrors = w + "checkErrors";
        public static readonly XName checkStyle = w + "checkStyle";
        public static readonly XName citation = w + "citation";
        public static readonly XName clear = w + "clear";
        public static readonly XName clickAndTypeStyle = w + "clickAndTypeStyle";
        public static readonly XName clrSchemeMapping = w + "clrSchemeMapping";
        public static readonly XName cnfStyle = w + "cnfStyle";
        public static readonly XName code = w + "code";
        public static readonly XName col = w + "col";
        public static readonly XName colDelim = w + "colDelim";
        public static readonly XName colFirst = w + "colFirst";
        public static readonly XName colLast = w + "colLast";
        public static readonly XName color = w + "color";
        public static readonly XName cols = w + "cols";
        public static readonly XName column = w + "column";
        public static readonly XName combine = w + "combine";
        public static readonly XName combineBrackets = w + "combineBrackets";
        public static readonly XName comboBox = w + "comboBox";
        public static readonly XName comment = w + "comment";
        public static readonly XName commentRangeEnd = w + "commentRangeEnd";
        public static readonly XName commentRangeStart = w + "commentRangeStart";
        public static readonly XName commentReference = w + "commentReference";
        public static readonly XName comments = w + "comments";
        public static readonly XName compat = w + "compat";
        public static readonly XName compatSetting = w + "compatSetting";
        public static readonly XName connectString = w + "connectString";
        public static readonly XName consecutiveHyphenLimit = w + "consecutiveHyphenLimit";
        public static readonly XName contentPart = w + "contentPart";
        public static readonly XName contextualSpacing = w + "contextualSpacing";
        public static readonly XName continuationSeparator = w + "continuationSeparator";
        public static readonly XName control = w + "control";
        public static readonly XName convMailMergeEsc = w + "convMailMergeEsc";
        public static readonly XName count = w + "count";
        public static readonly XName countBy = w + "countBy";
        public static readonly XName cr = w + "cr";
        public static readonly XName cryptAlgorithmClass = w + "cryptAlgorithmClass";
        public static readonly XName cryptAlgorithmSid = w + "cryptAlgorithmSid";
        public static readonly XName cryptAlgorithmType = w + "cryptAlgorithmType";
        public static readonly XName cryptProvider = w + "cryptProvider";
        public static readonly XName cryptProviderType = w + "cryptProviderType";
        public static readonly XName cryptProviderTypeExt = w + "cryptProviderTypeExt";
        public static readonly XName cryptProviderTypeExtSource = w + "cryptProviderTypeExtSource";
        public static readonly XName cryptSpinCount = w + "cryptSpinCount";
        public static readonly XName cs = w + "cs";
        public static readonly XName csb0 = w + "csb0";
        public static readonly XName csb1 = w + "csb1";
        public static readonly XName cstheme = w + "cstheme";
        public static readonly XName customMarkFollows = w + "customMarkFollows";
        public static readonly XName customStyle = w + "customStyle";
        public static readonly XName customXml = w + "customXml";
        public static readonly XName customXmlDelRangeEnd = w + "customXmlDelRangeEnd";
        public static readonly XName customXmlDelRangeStart = w + "customXmlDelRangeStart";
        public static readonly XName customXmlInsRangeEnd = w + "customXmlInsRangeEnd";
        public static readonly XName customXmlInsRangeStart = w + "customXmlInsRangeStart";
        public static readonly XName customXmlMoveFromRangeEnd = w + "customXmlMoveFromRangeEnd";
        public static readonly XName customXmlMoveFromRangeStart = w + "customXmlMoveFromRangeStart";
        public static readonly XName customXmlMoveToRangeEnd = w + "customXmlMoveToRangeEnd";
        public static readonly XName customXmlMoveToRangeStart = w + "customXmlMoveToRangeStart";
        public static readonly XName customXmlPr = w + "customXmlPr";
        public static readonly XName dataBinding = w + "dataBinding";
        public static readonly XName dataSource = w + "dataSource";
        public static readonly XName dataType = w + "dataType";
        public static readonly XName date = w + "date";
        public static readonly XName dateFormat = w + "dateFormat";
        public static readonly XName dayLong = w + "dayLong";
        public static readonly XName dayShort = w + "dayShort";
        public static readonly XName ddList = w + "ddList";
        public static readonly XName decimalSymbol = w + "decimalSymbol";
        public static readonly XName _default = w + "default";
        public static readonly XName defaultTableStyle = w + "defaultTableStyle";
        public static readonly XName defaultTabStop = w + "defaultTabStop";
        public static readonly XName defLockedState = w + "defLockedState";
        public static readonly XName defQFormat = w + "defQFormat";
        public static readonly XName defSemiHidden = w + "defSemiHidden";
        public static readonly XName defUIPriority = w + "defUIPriority";
        public static readonly XName defUnhideWhenUsed = w + "defUnhideWhenUsed";
        public static readonly XName del = w + "del";
        public static readonly XName delInstrText = w + "delInstrText";
        public static readonly XName delText = w + "delText";
        public static readonly XName description = w + "description";
        public static readonly XName destination = w + "destination";
        public static readonly XName dir = w + "dir";
        public static readonly XName dirty = w + "dirty";
        public static readonly XName displacedByCustomXml = w + "displacedByCustomXml";
        public static readonly XName display = w + "display";
        public static readonly XName displayBackgroundShape = w + "displayBackgroundShape";
        public static readonly XName displayHangulFixedWidth = w + "displayHangulFixedWidth";
        public static readonly XName displayHorizontalDrawingGridEvery = w + "displayHorizontalDrawingGridEvery";
        public static readonly XName displayText = w + "displayText";
        public static readonly XName displayVerticalDrawingGridEvery = w + "displayVerticalDrawingGridEvery";
        public static readonly XName distance = w + "distance";
        public static readonly XName div = w + "div";
        public static readonly XName divBdr = w + "divBdr";
        public static readonly XName divId = w + "divId";
        public static readonly XName divs = w + "divs";
        public static readonly XName divsChild = w + "divsChild";
        public static readonly XName dllVersion = w + "dllVersion";
        public static readonly XName docDefaults = w + "docDefaults";
        public static readonly XName docGrid = w + "docGrid";
        public static readonly XName docLocation = w + "docLocation";
        public static readonly XName docPart = w + "docPart";
        public static readonly XName docPartBody = w + "docPartBody";
        public static readonly XName docPartCategory = w + "docPartCategory";
        public static readonly XName docPartGallery = w + "docPartGallery";
        public static readonly XName docPartList = w + "docPartList";
        public static readonly XName docPartObj = w + "docPartObj";
        public static readonly XName docPartPr = w + "docPartPr";
        public static readonly XName docParts = w + "docParts";
        public static readonly XName docPartUnique = w + "docPartUnique";
        public static readonly XName document = w + "document";
        public static readonly XName documentProtection = w + "documentProtection";
        public static readonly XName documentType = w + "documentType";
        public static readonly XName docVar = w + "docVar";
        public static readonly XName docVars = w + "docVars";
        public static readonly XName doNotAutoCompressPictures = w + "doNotAutoCompressPictures";
        public static readonly XName doNotAutofitConstrainedTables = w + "doNotAutofitConstrainedTables";
        public static readonly XName doNotBreakConstrainedForcedTable = w + "doNotBreakConstrainedForcedTable";
        public static readonly XName doNotBreakWrappedTables = w + "doNotBreakWrappedTables";
        public static readonly XName doNotDemarcateInvalidXml = w + "doNotDemarcateInvalidXml";
        public static readonly XName doNotDisplayPageBoundaries = w + "doNotDisplayPageBoundaries";
        public static readonly XName doNotEmbedSmartTags = w + "doNotEmbedSmartTags";
        public static readonly XName doNotExpandShiftReturn = w + "doNotExpandShiftReturn";
        public static readonly XName doNotHyphenateCaps = w + "doNotHyphenateCaps";
        public static readonly XName doNotIncludeSubdocsInStats = w + "doNotIncludeSubdocsInStats";
        public static readonly XName doNotLeaveBackslashAlone = w + "doNotLeaveBackslashAlone";
        public static readonly XName doNotOrganizeInFolder = w + "doNotOrganizeInFolder";
        public static readonly XName doNotRelyOnCSS = w + "doNotRelyOnCSS";
        public static readonly XName doNotSaveAsSingleFile = w + "doNotSaveAsSingleFile";
        public static readonly XName doNotShadeFormData = w + "doNotShadeFormData";
        public static readonly XName doNotSnapToGridInCell = w + "doNotSnapToGridInCell";
        public static readonly XName doNotSuppressBlankLines = w + "doNotSuppressBlankLines";
        public static readonly XName doNotSuppressIndentation = w + "doNotSuppressIndentation";
        public static readonly XName doNotSuppressParagraphBorders = w + "doNotSuppressParagraphBorders";
        public static readonly XName doNotTrackFormatting = w + "doNotTrackFormatting";
        public static readonly XName doNotTrackMoves = w + "doNotTrackMoves";
        public static readonly XName doNotUseEastAsianBreakRules = w + "doNotUseEastAsianBreakRules";
        public static readonly XName doNotUseHTMLParagraphAutoSpacing = w + "doNotUseHTMLParagraphAutoSpacing";
        public static readonly XName doNotUseIndentAsNumberingTabStop = w + "doNotUseIndentAsNumberingTabStop";
        public static readonly XName doNotUseLongFileNames = w + "doNotUseLongFileNames";
        public static readonly XName doNotUseMarginsForDrawingGridOrigin = w + "doNotUseMarginsForDrawingGridOrigin";
        public static readonly XName doNotValidateAgainstSchema = w + "doNotValidateAgainstSchema";
        public static readonly XName doNotVertAlignCellWithSp = w + "doNotVertAlignCellWithSp";
        public static readonly XName doNotVertAlignInTxbx = w + "doNotVertAlignInTxbx";
        public static readonly XName doNotWrapTextWithPunct = w + "doNotWrapTextWithPunct";
        public static readonly XName drawing = w + "drawing";
        public static readonly XName drawingGridHorizontalOrigin = w + "drawingGridHorizontalOrigin";
        public static readonly XName drawingGridHorizontalSpacing = w + "drawingGridHorizontalSpacing";
        public static readonly XName drawingGridVerticalOrigin = w + "drawingGridVerticalOrigin";
        public static readonly XName drawingGridVerticalSpacing = w + "drawingGridVerticalSpacing";
        public static readonly XName dropCap = w + "dropCap";
        public static readonly XName dropDownList = w + "dropDownList";
        public static readonly XName dstrike = w + "dstrike";
        public static readonly XName dxaOrig = w + "dxaOrig";
        public static readonly XName dyaOrig = w + "dyaOrig";
        public static readonly XName dynamicAddress = w + "dynamicAddress";
        public static readonly XName eastAsia = w + "eastAsia";
        public static readonly XName eastAsianLayout = w + "eastAsianLayout";
        public static readonly XName eastAsiaTheme = w + "eastAsiaTheme";
        public static readonly XName ed = w + "ed";
        public static readonly XName edGrp = w + "edGrp";
        public static readonly XName edit = w + "edit";
        public static readonly XName effect = w + "effect";
        public static readonly XName element = w + "element";
        public static readonly XName em = w + "em";
        public static readonly XName embedBold = w + "embedBold";
        public static readonly XName embedBoldItalic = w + "embedBoldItalic";
        public static readonly XName embedItalic = w + "embedItalic";
        public static readonly XName embedRegular = w + "embedRegular";
        public static readonly XName embedSystemFonts = w + "embedSystemFonts";
        public static readonly XName embedTrueTypeFonts = w + "embedTrueTypeFonts";
        public static readonly XName emboss = w + "emboss";
        public static readonly XName enabled = w + "enabled";
        public static readonly XName encoding = w + "encoding";
        public static readonly XName endnote = w + "endnote";
        public static readonly XName endnotePr = w + "endnotePr";
        public static readonly XName endnoteRef = w + "endnoteRef";
        public static readonly XName endnoteReference = w + "endnoteReference";
        public static readonly XName endnotes = w + "endnotes";
        public static readonly XName enforcement = w + "enforcement";
        public static readonly XName entryMacro = w + "entryMacro";
        public static readonly XName equalWidth = w + "equalWidth";
        public static readonly XName equation = w + "equation";
        public static readonly XName evenAndOddHeaders = w + "evenAndOddHeaders";
        public static readonly XName exitMacro = w + "exitMacro";
        public static readonly XName family = w + "family";
        public static readonly XName ffData = w + "ffData";
        public static readonly XName fHdr = w + "fHdr";
        public static readonly XName fieldMapData = w + "fieldMapData";
        public static readonly XName fill = w + "fill";
        public static readonly XName first = w + "first";
        public static readonly XName firstColumn = w + "firstColumn";
        public static readonly XName firstLine = w + "firstLine";
        public static readonly XName firstLineChars = w + "firstLineChars";
        public static readonly XName firstRow = w + "firstRow";
        public static readonly XName fitText = w + "fitText";
        public static readonly XName flatBorders = w + "flatBorders";
        public static readonly XName fldChar = w + "fldChar";
        public static readonly XName fldCharType = w + "fldCharType";
        public static readonly XName fldData = w + "fldData";
        public static readonly XName fldLock = w + "fldLock";
        public static readonly XName fldSimple = w + "fldSimple";
        public static readonly XName fmt = w + "fmt";
        public static readonly XName followedHyperlink = w + "followedHyperlink";
        public static readonly XName font = w + "font";
        public static readonly XName fontKey = w + "fontKey";
        public static readonly XName fonts = w + "fonts";
        public static readonly XName fontSz = w + "fontSz";
        public static readonly XName footer = w + "footer";
        public static readonly XName footerReference = w + "footerReference";
        public static readonly XName footnote = w + "footnote";
        public static readonly XName footnoteLayoutLikeWW8 = w + "footnoteLayoutLikeWW8";
        public static readonly XName footnotePr = w + "footnotePr";
        public static readonly XName footnoteRef = w + "footnoteRef";
        public static readonly XName footnoteReference = w + "footnoteReference";
        public static readonly XName footnotes = w + "footnotes";
        public static readonly XName forceUpgrade = w + "forceUpgrade";
        public static readonly XName forgetLastTabAlignment = w + "forgetLastTabAlignment";
        public static readonly XName format = w + "format";
        public static readonly XName formatting = w + "formatting";
        public static readonly XName formProt = w + "formProt";
        public static readonly XName formsDesign = w + "formsDesign";
        public static readonly XName frame = w + "frame";
        public static readonly XName frameLayout = w + "frameLayout";
        public static readonly XName framePr = w + "framePr";
        public static readonly XName frameset = w + "frameset";
        public static readonly XName framesetSplitbar = w + "framesetSplitbar";
        public static readonly XName ftr = w + "ftr";
        public static readonly XName fullDate = w + "fullDate";
        public static readonly XName gallery = w + "gallery";
        public static readonly XName glossaryDocument = w + "glossaryDocument";
        public static readonly XName grammar = w + "grammar";
        public static readonly XName gridAfter = w + "gridAfter";
        public static readonly XName gridBefore = w + "gridBefore";
        public static readonly XName gridCol = w + "gridCol";
        public static readonly XName gridSpan = w + "gridSpan";
        public static readonly XName group = w + "group";
        public static readonly XName growAutofit = w + "growAutofit";
        public static readonly XName guid = w + "guid";
        public static readonly XName gutter = w + "gutter";
        public static readonly XName gutterAtTop = w + "gutterAtTop";
        public static readonly XName h = w + "h";
        public static readonly XName hAnchor = w + "hAnchor";
        public static readonly XName hanging = w + "hanging";
        public static readonly XName hangingChars = w + "hangingChars";
        public static readonly XName hAnsi = w + "hAnsi";
        public static readonly XName hAnsiTheme = w + "hAnsiTheme";
        public static readonly XName hash = w + "hash";
        public static readonly XName hdr = w + "hdr";
        public static readonly XName hdrShapeDefaults = w + "hdrShapeDefaults";
        public static readonly XName header = w + "header";
        public static readonly XName headerReference = w + "headerReference";
        public static readonly XName headerSource = w + "headerSource";
        public static readonly XName helpText = w + "helpText";
        public static readonly XName hidden = w + "hidden";
        public static readonly XName hideGrammaticalErrors = w + "hideGrammaticalErrors";
        public static readonly XName hideMark = w + "hideMark";
        public static readonly XName hideSpellingErrors = w + "hideSpellingErrors";
        public static readonly XName highlight = w + "highlight";
        public static readonly XName hint = w + "hint";
        public static readonly XName history = w + "history";
        public static readonly XName hMerge = w + "hMerge";
        public static readonly XName horzAnchor = w + "horzAnchor";
        public static readonly XName hps = w + "hps";
        public static readonly XName hpsBaseText = w + "hpsBaseText";
        public static readonly XName hpsRaise = w + "hpsRaise";
        public static readonly XName hRule = w + "hRule";
        public static readonly XName hSpace = w + "hSpace";
        public static readonly XName hyperlink = w + "hyperlink";
        public static readonly XName hyphenationZone = w + "hyphenationZone";
        public static readonly XName i = w + "i";
        public static readonly XName iCs = w + "iCs";
        public static readonly XName id = w + "id";
        public static readonly XName ignoreMixedContent = w + "ignoreMixedContent";
        public static readonly XName ilvl = w + "ilvl";
        public static readonly XName imprint = w + "imprint";
        public static readonly XName ind = w + "ind";
        public static readonly XName initials = w + "initials";
        public static readonly XName inkAnnotations = w + "inkAnnotations";
        public static readonly XName ins = w + "ins";
        public static readonly XName insDel = w + "insDel";
        public static readonly XName insideH = w + "insideH";
        public static readonly XName insideV = w + "insideV";
        public static readonly XName instr = w + "instr";
        public static readonly XName instrText = w + "instrText";
        public static readonly XName isLgl = w + "isLgl";
        public static readonly XName jc = w + "jc";
        public static readonly XName keepLines = w + "keepLines";
        public static readonly XName keepNext = w + "keepNext";
        public static readonly XName kern = w + "kern";
        public static readonly XName kinsoku = w + "kinsoku";
        public static readonly XName lang = w + "lang";
        public static readonly XName lastColumn = w + "lastColumn";
        public static readonly XName lastRenderedPageBreak = w + "lastRenderedPageBreak";
        public static readonly XName lastValue = w + "lastValue";
        public static readonly XName lastRow = w + "lastRow";
        public static readonly XName latentStyles = w + "latentStyles";
        public static readonly XName layoutRawTableWidth = w + "layoutRawTableWidth";
        public static readonly XName layoutTableRowsApart = w + "layoutTableRowsApart";
        public static readonly XName leader = w + "leader";
        public static readonly XName left = w + "left";
        public static readonly XName leftChars = w + "leftChars";
        public static readonly XName leftFromText = w + "leftFromText";
        public static readonly XName legacy = w + "legacy";
        public static readonly XName legacyIndent = w + "legacyIndent";
        public static readonly XName legacySpace = w + "legacySpace";
        public static readonly XName lid = w + "lid";
        public static readonly XName line = w + "line";
        public static readonly XName linePitch = w + "linePitch";
        public static readonly XName lineRule = w + "lineRule";
        public static readonly XName lines = w + "lines";
        public static readonly XName lineWrapLikeWord6 = w + "lineWrapLikeWord6";
        public static readonly XName link = w + "link";
        public static readonly XName linkedToFile = w + "linkedToFile";
        public static readonly XName linkStyles = w + "linkStyles";
        public static readonly XName linkToQuery = w + "linkToQuery";
        public static readonly XName listEntry = w + "listEntry";
        public static readonly XName listItem = w + "listItem";
        public static readonly XName listSeparator = w + "listSeparator";
        public static readonly XName lnNumType = w + "lnNumType";
        public static readonly XName _lock = w + "lock";
        public static readonly XName locked = w + "locked";
        public static readonly XName lsdException = w + "lsdException";
        public static readonly XName lvl = w + "lvl";
        public static readonly XName lvlJc = w + "lvlJc";
        public static readonly XName lvlOverride = w + "lvlOverride";
        public static readonly XName lvlPicBulletId = w + "lvlPicBulletId";
        public static readonly XName lvlRestart = w + "lvlRestart";
        public static readonly XName lvlText = w + "lvlText";
        public static readonly XName mailAsAttachment = w + "mailAsAttachment";
        public static readonly XName mailMerge = w + "mailMerge";
        public static readonly XName mailSubject = w + "mailSubject";
        public static readonly XName mainDocumentType = w + "mainDocumentType";
        public static readonly XName mappedName = w + "mappedName";
        public static readonly XName marBottom = w + "marBottom";
        public static readonly XName marH = w + "marH";
        public static readonly XName markup = w + "markup";
        public static readonly XName marLeft = w + "marLeft";
        public static readonly XName marRight = w + "marRight";
        public static readonly XName marTop = w + "marTop";
        public static readonly XName marW = w + "marW";
        public static readonly XName matchSrc = w + "matchSrc";
        public static readonly XName maxLength = w + "maxLength";
        public static readonly XName mirrorIndents = w + "mirrorIndents";
        public static readonly XName mirrorMargins = w + "mirrorMargins";
        public static readonly XName monthLong = w + "monthLong";
        public static readonly XName monthShort = w + "monthShort";
        public static readonly XName moveFrom = w + "moveFrom";
        public static readonly XName moveFromRangeEnd = w + "moveFromRangeEnd";
        public static readonly XName moveFromRangeStart = w + "moveFromRangeStart";
        public static readonly XName moveTo = w + "moveTo";
        public static readonly XName moveToRangeEnd = w + "moveToRangeEnd";
        public static readonly XName moveToRangeStart = w + "moveToRangeStart";
        public static readonly XName multiLevelType = w + "multiLevelType";
        public static readonly XName multiLine = w + "multiLine";
        public static readonly XName mwSmallCaps = w + "mwSmallCaps";
        public static readonly XName name = w + "name";
        public static readonly XName namespaceuri = w + "namespaceuri";
        public static readonly XName next = w + "next";
        public static readonly XName nlCheck = w + "nlCheck";
        public static readonly XName noBorder = w + "noBorder";
        public static readonly XName noBreakHyphen = w + "noBreakHyphen";
        public static readonly XName noColumnBalance = w + "noColumnBalance";
        public static readonly XName noEndnote = w + "noEndnote";
        public static readonly XName noExtraLineSpacing = w + "noExtraLineSpacing";
        public static readonly XName noHBand = w + "noHBand";
        public static readonly XName noLeading = w + "noLeading";
        public static readonly XName noLineBreaksAfter = w + "noLineBreaksAfter";
        public static readonly XName noLineBreaksBefore = w + "noLineBreaksBefore";
        public static readonly XName noProof = w + "noProof";
        public static readonly XName noPunctuationKerning = w + "noPunctuationKerning";
        public static readonly XName noResizeAllowed = w + "noResizeAllowed";
        public static readonly XName noSpaceRaiseLower = w + "noSpaceRaiseLower";
        public static readonly XName noTabHangInd = w + "noTabHangInd";
        public static readonly XName notTrueType = w + "notTrueType";
        public static readonly XName noVBand = w + "noVBand";
        public static readonly XName noWrap = w + "noWrap";
        public static readonly XName nsid = w + "nsid";
        public static readonly XName _null = w + "null";
        public static readonly XName num = w + "num";
        public static readonly XName numbering = w + "numbering";
        public static readonly XName numberingChange = w + "numberingChange";
        public static readonly XName numFmt = w + "numFmt";
        public static readonly XName numId = w + "numId";
        public static readonly XName numIdMacAtCleanup = w + "numIdMacAtCleanup";
        public static readonly XName numPicBullet = w + "numPicBullet";
        public static readonly XName numPicBulletId = w + "numPicBulletId";
        public static readonly XName numPr = w + "numPr";
        public static readonly XName numRestart = w + "numRestart";
        public static readonly XName numStart = w + "numStart";
        public static readonly XName numStyleLink = w + "numStyleLink";
        public static readonly XName _object = w + "object";
        public static readonly XName odso = w + "odso";
        public static readonly XName offsetFrom = w + "offsetFrom";
        public static readonly XName oMath = w + "oMath";
        public static readonly XName optimizeForBrowser = w + "optimizeForBrowser";
        public static readonly XName orient = w + "orient";
        public static readonly XName original = w + "original";
        public static readonly XName other = w + "other";
        public static readonly XName outline = w + "outline";
        public static readonly XName outlineLvl = w + "outlineLvl";
        public static readonly XName overflowPunct = w + "overflowPunct";
        public static readonly XName p = w + "p";
        public static readonly XName pageBreakBefore = w + "pageBreakBefore";
        public static readonly XName panose1 = w + "panose1";
        public static readonly XName paperSrc = w + "paperSrc";
        public static readonly XName pBdr = w + "pBdr";
        public static readonly XName percent = w + "percent";
        public static readonly XName permEnd = w + "permEnd";
        public static readonly XName permStart = w + "permStart";
        public static readonly XName personal = w + "personal";
        public static readonly XName personalCompose = w + "personalCompose";
        public static readonly XName personalReply = w + "personalReply";
        public static readonly XName pgBorders = w + "pgBorders";
        public static readonly XName pgMar = w + "pgMar";
        public static readonly XName pgNum = w + "pgNum";
        public static readonly XName pgNumType = w + "pgNumType";
        public static readonly XName pgSz = w + "pgSz";
        public static readonly XName pict = w + "pict";
        public static readonly XName picture = w + "picture";
        public static readonly XName pitch = w + "pitch";
        public static readonly XName pixelsPerInch = w + "pixelsPerInch";
        public static readonly XName placeholder = w + "placeholder";
        public static readonly XName pos = w + "pos";
        public static readonly XName position = w + "position";
        public static readonly XName pPr = w + "pPr";
        public static readonly XName pPrChange = w + "pPrChange";
        public static readonly XName pPrDefault = w + "pPrDefault";
        public static readonly XName prefixMappings = w + "prefixMappings";
        public static readonly XName printBodyTextBeforeHeader = w + "printBodyTextBeforeHeader";
        public static readonly XName printColBlack = w + "printColBlack";
        public static readonly XName printerSettings = w + "printerSettings";
        public static readonly XName printFormsData = w + "printFormsData";
        public static readonly XName printFractionalCharacterWidth = w + "printFractionalCharacterWidth";
        public static readonly XName printPostScriptOverText = w + "printPostScriptOverText";
        public static readonly XName printTwoOnOne = w + "printTwoOnOne";
        public static readonly XName proofErr = w + "proofErr";
        public static readonly XName proofState = w + "proofState";
        public static readonly XName pStyle = w + "pStyle";
        public static readonly XName ptab = w + "ptab";
        public static readonly XName qFormat = w + "qFormat";
        public static readonly XName query = w + "query";
        public static readonly XName r = w + "r";
        public static readonly XName readModeInkLockDown = w + "readModeInkLockDown";
        public static readonly XName recipientData = w + "recipientData";
        public static readonly XName recommended = w + "recommended";
        public static readonly XName relativeTo = w + "relativeTo";
        public static readonly XName relyOnVML = w + "relyOnVML";
        public static readonly XName removeDateAndTime = w + "removeDateAndTime";
        public static readonly XName removePersonalInformation = w + "removePersonalInformation";
        public static readonly XName restart = w + "restart";
        public static readonly XName result = w + "result";
        public static readonly XName revisionView = w + "revisionView";
        public static readonly XName rFonts = w + "rFonts";
        public static readonly XName richText = w + "richText";
        public static readonly XName right = w + "right";
        public static readonly XName rightChars = w + "rightChars";
        public static readonly XName rightFromText = w + "rightFromText";
        public static readonly XName rPr = w + "rPr";
        public static readonly XName rPrChange = w + "rPrChange";
        public static readonly XName rPrDefault = w + "rPrDefault";
        public static readonly XName rsid = w + "rsid";
        public static readonly XName rsidDel = w + "rsidDel";
        public static readonly XName rsidP = w + "rsidP";
        public static readonly XName rsidR = w + "rsidR";
        public static readonly XName rsidRDefault = w + "rsidRDefault";
        public static readonly XName rsidRoot = w + "rsidRoot";
        public static readonly XName rsidRPr = w + "rsidRPr";
        public static readonly XName rsids = w + "rsids";
        public static readonly XName rsidSect = w + "rsidSect";
        public static readonly XName rsidTr = w + "rsidTr";
        public static readonly XName rStyle = w + "rStyle";
        public static readonly XName rt = w + "rt";
        public static readonly XName rtl = w + "rtl";
        public static readonly XName rtlGutter = w + "rtlGutter";
        public static readonly XName ruby = w + "ruby";
        public static readonly XName rubyAlign = w + "rubyAlign";
        public static readonly XName rubyBase = w + "rubyBase";
        public static readonly XName rubyPr = w + "rubyPr";
        public static readonly XName salt = w + "salt";
        public static readonly XName saveFormsData = w + "saveFormsData";
        public static readonly XName saveInvalidXml = w + "saveInvalidXml";
        public static readonly XName savePreviewPicture = w + "savePreviewPicture";
        public static readonly XName saveSmartTagsAsXml = w + "saveSmartTagsAsXml";
        public static readonly XName saveSubsetFonts = w + "saveSubsetFonts";
        public static readonly XName saveThroughXslt = w + "saveThroughXslt";
        public static readonly XName saveXmlDataOnly = w + "saveXmlDataOnly";
        public static readonly XName scrollbar = w + "scrollbar";
        public static readonly XName sdt = w + "sdt";
        public static readonly XName sdtContent = w + "sdtContent";
        public static readonly XName sdtEndPr = w + "sdtEndPr";
        public static readonly XName sdtPr = w + "sdtPr";
        public static readonly XName sectPr = w + "sectPr";
        public static readonly XName sectPrChange = w + "sectPrChange";
        public static readonly XName selectFldWithFirstOrLastChar = w + "selectFldWithFirstOrLastChar";
        public static readonly XName semiHidden = w + "semiHidden";
        public static readonly XName sep = w + "sep";
        public static readonly XName separator = w + "separator";
        public static readonly XName settings = w + "settings";
        public static readonly XName shadow = w + "shadow";
        public static readonly XName shapeDefaults = w + "shapeDefaults";
        public static readonly XName shapeid = w + "shapeid";
        public static readonly XName shapeLayoutLikeWW8 = w + "shapeLayoutLikeWW8";
        public static readonly XName shd = w + "shd";
        public static readonly XName showBreaksInFrames = w + "showBreaksInFrames";
        public static readonly XName showEnvelope = w + "showEnvelope";
        public static readonly XName showingPlcHdr = w + "showingPlcHdr";
        public static readonly XName showXMLTags = w + "showXMLTags";
        public static readonly XName sig = w + "sig";
        public static readonly XName size = w + "size";
        public static readonly XName sizeAuto = w + "sizeAuto";
        public static readonly XName smallCaps = w + "smallCaps";
        public static readonly XName smartTag = w + "smartTag";
        public static readonly XName smartTagPr = w + "smartTagPr";
        public static readonly XName smartTagType = w + "smartTagType";
        public static readonly XName snapToGrid = w + "snapToGrid";
        public static readonly XName softHyphen = w + "softHyphen";
        public static readonly XName solutionID = w + "solutionID";
        public static readonly XName sourceFileName = w + "sourceFileName";
        public static readonly XName space = w + "space";
        public static readonly XName spaceForUL = w + "spaceForUL";
        public static readonly XName spacing = w + "spacing";
        public static readonly XName spacingInWholePoints = w + "spacingInWholePoints";
        public static readonly XName specVanish = w + "specVanish";
        public static readonly XName spelling = w + "spelling";
        public static readonly XName splitPgBreakAndParaMark = w + "splitPgBreakAndParaMark";
        public static readonly XName src = w + "src";
        public static readonly XName start = w + "start";
        public static readonly XName startOverride = w + "startOverride";
        public static readonly XName statusText = w + "statusText";
        public static readonly XName storeItemID = w + "storeItemID";
        public static readonly XName storeMappedDataAs = w + "storeMappedDataAs";
        public static readonly XName strictFirstAndLastChars = w + "strictFirstAndLastChars";
        public static readonly XName strike = w + "strike";
        public static readonly XName style = w + "style";
        public static readonly XName styleId = w + "styleId";
        public static readonly XName styleLink = w + "styleLink";
        public static readonly XName styleLockQFSet = w + "styleLockQFSet";
        public static readonly XName styleLockTheme = w + "styleLockTheme";
        public static readonly XName stylePaneFormatFilter = w + "stylePaneFormatFilter";
        public static readonly XName stylePaneSortMethod = w + "stylePaneSortMethod";
        public static readonly XName styles = w + "styles";
        public static readonly XName subDoc = w + "subDoc";
        public static readonly XName subFontBySize = w + "subFontBySize";
        public static readonly XName subsetted = w + "subsetted";
        public static readonly XName suff = w + "suff";
        public static readonly XName summaryLength = w + "summaryLength";
        public static readonly XName suppressAutoHyphens = w + "suppressAutoHyphens";
        public static readonly XName suppressBottomSpacing = w + "suppressBottomSpacing";
        public static readonly XName suppressLineNumbers = w + "suppressLineNumbers";
        public static readonly XName suppressOverlap = w + "suppressOverlap";
        public static readonly XName suppressSpacingAtTopOfPage = w + "suppressSpacingAtTopOfPage";
        public static readonly XName suppressSpBfAfterPgBrk = w + "suppressSpBfAfterPgBrk";
        public static readonly XName suppressTopSpacing = w + "suppressTopSpacing";
        public static readonly XName suppressTopSpacingWP = w + "suppressTopSpacingWP";
        public static readonly XName swapBordersFacingPages = w + "swapBordersFacingPages";
        public static readonly XName sym = w + "sym";
        public static readonly XName sz = w + "sz";
        public static readonly XName szCs = w + "szCs";
        public static readonly XName t = w + "t";
        public static readonly XName t1 = w + "t1";
        public static readonly XName t2 = w + "t2";
        public static readonly XName tab = w + "tab";
        public static readonly XName table = w + "table";
        public static readonly XName tabs = w + "tabs";
        public static readonly XName tag = w + "tag";
        public static readonly XName targetScreenSz = w + "targetScreenSz";
        public static readonly XName tbl = w + "tbl";
        public static readonly XName tblBorders = w + "tblBorders";
        public static readonly XName tblCellMar = w + "tblCellMar";
        public static readonly XName tblCellSpacing = w + "tblCellSpacing";
        public static readonly XName tblGrid = w + "tblGrid";
        public static readonly XName tblGridChange = w + "tblGridChange";
        public static readonly XName tblHeader = w + "tblHeader";
        public static readonly XName tblInd = w + "tblInd";
        public static readonly XName tblLayout = w + "tblLayout";
        public static readonly XName tblLook = w + "tblLook";
        public static readonly XName tblOverlap = w + "tblOverlap";
        public static readonly XName tblpPr = w + "tblpPr";
        public static readonly XName tblPr = w + "tblPr";
        public static readonly XName tblPrChange = w + "tblPrChange";
        public static readonly XName tblPrEx = w + "tblPrEx";
        public static readonly XName tblPrExChange = w + "tblPrExChange";
        public static readonly XName tblpX = w + "tblpX";
        public static readonly XName tblpXSpec = w + "tblpXSpec";
        public static readonly XName tblpY = w + "tblpY";
        public static readonly XName tblpYSpec = w + "tblpYSpec";
        public static readonly XName tblStyle = w + "tblStyle";
        public static readonly XName tblStyleColBandSize = w + "tblStyleColBandSize";
        public static readonly XName tblStylePr = w + "tblStylePr";
        public static readonly XName tblStyleRowBandSize = w + "tblStyleRowBandSize";
        public static readonly XName tblW = w + "tblW";
        public static readonly XName tc = w + "tc";
        public static readonly XName tcBorders = w + "tcBorders";
        public static readonly XName tcFitText = w + "tcFitText";
        public static readonly XName tcMar = w + "tcMar";
        public static readonly XName tcPr = w + "tcPr";
        public static readonly XName tcPrChange = w + "tcPrChange";
        public static readonly XName tcW = w + "tcW";
        public static readonly XName temporary = w + "temporary";
        public static readonly XName tentative = w + "tentative";
        public static readonly XName text = w + "text";
        public static readonly XName textAlignment = w + "textAlignment";
        public static readonly XName textboxTightWrap = w + "textboxTightWrap";
        public static readonly XName textDirection = w + "textDirection";
        public static readonly XName textInput = w + "textInput";
        public static readonly XName tgtFrame = w + "tgtFrame";
        public static readonly XName themeColor = w + "themeColor";
        public static readonly XName themeFill = w + "themeFill";
        public static readonly XName themeFillShade = w + "themeFillShade";
        public static readonly XName themeFillTint = w + "themeFillTint";
        public static readonly XName themeFontLang = w + "themeFontLang";
        public static readonly XName themeShade = w + "themeShade";
        public static readonly XName themeTint = w + "themeTint";
        public static readonly XName titlePg = w + "titlePg";
        public static readonly XName tl2br = w + "tl2br";
        public static readonly XName tmpl = w + "tmpl";
        public static readonly XName tooltip = w + "tooltip";
        public static readonly XName top = w + "top";
        public static readonly XName topFromText = w + "topFromText";
        public static readonly XName topLinePunct = w + "topLinePunct";
        public static readonly XName tplc = w + "tplc";
        public static readonly XName tr = w + "tr";
        public static readonly XName tr2bl = w + "tr2bl";
        public static readonly XName trackRevisions = w + "trackRevisions";
        public static readonly XName trHeight = w + "trHeight";
        public static readonly XName trPr = w + "trPr";
        public static readonly XName trPrChange = w + "trPrChange";
        public static readonly XName truncateFontHeightsLikeWP6 = w + "truncateFontHeightsLikeWP6";
        public static readonly XName txbxContent = w + "txbxContent";
        public static readonly XName type = w + "type";
        public static readonly XName types = w + "types";
        public static readonly XName u = w + "u";
        public static readonly XName udl = w + "udl";
        public static readonly XName uiCompat97To2003 = w + "uiCompat97To2003";
        public static readonly XName uiPriority = w + "uiPriority";
        public static readonly XName ulTrailSpace = w + "ulTrailSpace";
        public static readonly XName underlineTabInNumList = w + "underlineTabInNumList";
        public static readonly XName unhideWhenUsed = w + "unhideWhenUsed";
        public static readonly XName updateFields = w + "updateFields";
        public static readonly XName uri = w + "uri";
        public static readonly XName url = w + "url";
        public static readonly XName usb0 = w + "usb0";
        public static readonly XName usb1 = w + "usb1";
        public static readonly XName usb2 = w + "usb2";
        public static readonly XName usb3 = w + "usb3";
        public static readonly XName useAltKinsokuLineBreakRules = w + "useAltKinsokuLineBreakRules";
        public static readonly XName useAnsiKerningPairs = w + "useAnsiKerningPairs";
        public static readonly XName useFELayout = w + "useFELayout";
        public static readonly XName useNormalStyleForList = w + "useNormalStyleForList";
        public static readonly XName usePrinterMetrics = w + "usePrinterMetrics";
        public static readonly XName useSingleBorderforContiguousCells = w + "useSingleBorderforContiguousCells";
        public static readonly XName useWord2002TableStyleRules = w + "useWord2002TableStyleRules";
        public static readonly XName useWord97LineBreakRules = w + "useWord97LineBreakRules";
        public static readonly XName useXSLTWhenSaving = w + "useXSLTWhenSaving";
        public static readonly XName val = w + "val";
        public static readonly XName vAlign = w + "vAlign";
        public static readonly XName value = w + "value";
        public static readonly XName vAnchor = w + "vAnchor";
        public static readonly XName vanish = w + "vanish";
        public static readonly XName vendorID = w + "vendorID";
        public static readonly XName vert = w + "vert";
        public static readonly XName vertAlign = w + "vertAlign";
        public static readonly XName vertAnchor = w + "vertAnchor";
        public static readonly XName vertCompress = w + "vertCompress";
        public static readonly XName view = w + "view";
        public static readonly XName viewMergedData = w + "viewMergedData";
        public static readonly XName vMerge = w + "vMerge";
        public static readonly XName vMergeOrig = w + "vMergeOrig";
        public static readonly XName vSpace = w + "vSpace";
        public static readonly XName _w = w + "w";
        public static readonly XName wAfter = w + "wAfter";
        public static readonly XName wBefore = w + "wBefore";
        public static readonly XName webHidden = w + "webHidden";
        public static readonly XName webSettings = w + "webSettings";
        public static readonly XName widowControl = w + "widowControl";
        public static readonly XName wordWrap = w + "wordWrap";
        public static readonly XName wpJustification = w + "wpJustification";
        public static readonly XName wpSpaceWidth = w + "wpSpaceWidth";
        public static readonly XName wrap = w + "wrap";
        public static readonly XName wrapTrailSpaces = w + "wrapTrailSpaces";
        public static readonly XName writeProtection = w + "writeProtection";
        public static readonly XName x = w + "x";
        public static readonly XName xAlign = w + "xAlign";
        public static readonly XName xpath = w + "xpath";
        public static readonly XName y = w + "y";
        public static readonly XName yAlign = w + "yAlign";
        public static readonly XName yearLong = w + "yearLong";
        public static readonly XName yearShort = w + "yearShort";
        public static readonly XName zoom = w + "zoom";
        public static readonly XName zOrder = w + "zOrder";
        public static readonly XName tblCaption = w + "tblCaption";
        public static readonly XName tblDescription = w + "tblDescription";
        public static readonly XName startChars = w + "startChars";
        public static readonly XName end = w + "end";
        public static readonly XName endChars = w + "endChars";
        public static readonly XName evenHBand = w + "evenHBand";
        public static readonly XName evenVBand = w + "evenVBand";
        public static readonly XName firstRowFirstColumn = w + "firstRowFirstColumn";
        public static readonly XName firstRowLastColumn = w + "firstRowLastColumn";
        public static readonly XName lastRowFirstColumn = w + "lastRowFirstColumn";
        public static readonly XName lastRowLastColumn = w + "lastRowLastColumn";
        public static readonly XName oddHBand = w + "oddHBand";
        public static readonly XName oddVBand = w + "oddVBand";
        public static readonly XName headers = w + "headers";

        public static readonly XName[] BlockLevelContentContainers =
        {
            W.body,
            W.tc,
            W.txbxContent,
            W.hdr,
            W.ftr,
            W.endnote,
            W.footnote
        };

        public static readonly XName[] SubRunLevelContent =
        {
            W.br,
            W.cr,
            W.dayLong,
            W.dayShort,
            W.drawing,
            W.drawing,
            W.monthLong,
            W.monthShort,
            W.noBreakHyphen,
            W.ptab,
            W.pgNum,
            W.pict,
            W.softHyphen,
            W.sym,
            W.t,
            W.tab,
            W.yearLong,
            W.yearShort,
            MC.AlternateContent,
        };
    }

    public static class W10
    {
        public static readonly XNamespace w10 =
            "urn:schemas-microsoft-com:office:word";
        public static readonly XName anchorlock = w10 + "anchorlock";
        public static readonly XName borderbottom = w10 + "borderbottom";
        public static readonly XName borderleft = w10 + "borderleft";
        public static readonly XName borderright = w10 + "borderright";
        public static readonly XName bordertop = w10 + "bordertop";
        public static readonly XName wrap = w10 + "wrap";
    }

    public static class W14
    {
        public static readonly XNamespace w14 =
            "http://schemas.microsoft.com/office/word/2010/wordml";
        public static readonly XName algn = w14 + "algn";
        public static readonly XName alpha = w14 + "alpha";
        public static readonly XName ang = w14 + "ang";
        public static readonly XName b = w14 + "b";
        public static readonly XName bevel = w14 + "bevel";
        public static readonly XName bevelB = w14 + "bevelB";
        public static readonly XName bevelT = w14 + "bevelT";
        public static readonly XName blurRad = w14 + "blurRad";
        public static readonly XName camera = w14 + "camera";
        public static readonly XName cap = w14 + "cap";
        public static readonly XName checkbox = w14 + "checkbox";
        public static readonly XName _checked = w14 + "checked";
        public static readonly XName checkedState = w14 + "checkedState";
        public static readonly XName cmpd = w14 + "cmpd";
        public static readonly XName cntxtAlts = w14 + "cntxtAlts";
        public static readonly XName cNvContentPartPr = w14 + "cNvContentPartPr";
        public static readonly XName conflictMode = w14 + "conflictMode";
        public static readonly XName contentPart = w14 + "contentPart";
        public static readonly XName contourClr = w14 + "contourClr";
        public static readonly XName contourW = w14 + "contourW";
        public static readonly XName defaultImageDpi = w14 + "defaultImageDpi";
        public static readonly XName dir = w14 + "dir";
        public static readonly XName discardImageEditingData = w14 + "discardImageEditingData";
        public static readonly XName dist = w14 + "dist";
        public static readonly XName docId = w14 + "docId";
        public static readonly XName editId = w14 + "editId";
        public static readonly XName enableOpenTypeKerning = w14 + "enableOpenTypeKerning";
        public static readonly XName endA = w14 + "endA";
        public static readonly XName endPos = w14 + "endPos";
        public static readonly XName entityPicker = w14 + "entityPicker";
        public static readonly XName extrusionClr = w14 + "extrusionClr";
        public static readonly XName extrusionH = w14 + "extrusionH";
        public static readonly XName fadeDir = w14 + "fadeDir";
        public static readonly XName fillToRect = w14 + "fillToRect";
        public static readonly XName font = w14 + "font";
        public static readonly XName glow = w14 + "glow";
        public static readonly XName gradFill = w14 + "gradFill";
        public static readonly XName gs = w14 + "gs";
        public static readonly XName gsLst = w14 + "gsLst";
        public static readonly XName h = w14 + "h";
        public static readonly XName hueMod = w14 + "hueMod";
        public static readonly XName id = w14 + "id";
        public static readonly XName kx = w14 + "kx";
        public static readonly XName ky = w14 + "ky";
        public static readonly XName l = w14 + "l";
        public static readonly XName lat = w14 + "lat";
        public static readonly XName ligatures = w14 + "ligatures";
        public static readonly XName lightRig = w14 + "lightRig";
        public static readonly XName lim = w14 + "lim";
        public static readonly XName lin = w14 + "lin";
        public static readonly XName lon = w14 + "lon";
        public static readonly XName lumMod = w14 + "lumMod";
        public static readonly XName lumOff = w14 + "lumOff";
        public static readonly XName miter = w14 + "miter";
        public static readonly XName noFill = w14 + "noFill";
        public static readonly XName numForm = w14 + "numForm";
        public static readonly XName numSpacing = w14 + "numSpacing";
        public static readonly XName nvContentPartPr = w14 + "nvContentPartPr";
        public static readonly XName paraId = w14 + "paraId";
        public static readonly XName path = w14 + "path";
        public static readonly XName pos = w14 + "pos";
        public static readonly XName props3d = w14 + "props3d";
        public static readonly XName prst = w14 + "prst";
        public static readonly XName prstDash = w14 + "prstDash";
        public static readonly XName prstMaterial = w14 + "prstMaterial";
        public static readonly XName r = w14 + "r";
        public static readonly XName rad = w14 + "rad";
        public static readonly XName reflection = w14 + "reflection";
        public static readonly XName rev = w14 + "rev";
        public static readonly XName rig = w14 + "rig";
        public static readonly XName rot = w14 + "rot";
        public static readonly XName round = w14 + "round";
        public static readonly XName sat = w14 + "sat";
        public static readonly XName satMod = w14 + "satMod";
        public static readonly XName satOff = w14 + "satOff";
        public static readonly XName scaled = w14 + "scaled";
        public static readonly XName scene3d = w14 + "scene3d";
        public static readonly XName schemeClr = w14 + "schemeClr";
        public static readonly XName shade = w14 + "shade";
        public static readonly XName shadow = w14 + "shadow";
        public static readonly XName solidFill = w14 + "solidFill";
        public static readonly XName srgbClr = w14 + "srgbClr";
        public static readonly XName stA = w14 + "stA";
        public static readonly XName stPos = w14 + "stPos";
        public static readonly XName styleSet = w14 + "styleSet";
        public static readonly XName stylisticSets = w14 + "stylisticSets";
        public static readonly XName sx = w14 + "sx";
        public static readonly XName sy = w14 + "sy";
        public static readonly XName t = w14 + "t";
        public static readonly XName textFill = w14 + "textFill";
        public static readonly XName textId = w14 + "textId";
        public static readonly XName textOutline = w14 + "textOutline";
        public static readonly XName tint = w14 + "tint";
        public static readonly XName uncheckedState = w14 + "uncheckedState";
        public static readonly XName val = w14 + "val";
        public static readonly XName w = w14 + "w";
        public static readonly XName wProps3d = w14 + "wProps3d";
        public static readonly XName wScene3d = w14 + "wScene3d";
        public static readonly XName wShadow = w14 + "wShadow";
        public static readonly XName wTextFill = w14 + "wTextFill";
        public static readonly XName wTextOutline = w14 + "wTextOutline";
        public static readonly XName xfrm = w14 + "xfrm";
    }

    public static class W15
    {
        public static XNamespace w15 = "http://schemas.microsoft.com/office/word/2012/wordml";
    }

    public static class W16SE
    {
        public static XNamespace w16se = "http://schemas.microsoft.com/office/word/2015/wordml/symex";
    }

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

    public static class WETP
    {
        public static readonly XNamespace wetp = "http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11";
        public static readonly XName extLst = wetp + "extLst";
        public static readonly XName taskpane = wetp + "taskpane";
        public static readonly XName taskpanes = wetp + "taskpanes";
        public static readonly XName web_extension_taskpanes = wetp + "web-extension-taskpanes";
        public static readonly XName webextensionref = wetp + "webextensionref";
    }

    public static class W3DIGSIG
    {
        public static readonly XNamespace w3digsig =
            "http://www.w3.org/2000/09/xmldsig#";
        public static readonly XName CanonicalizationMethod = w3digsig + "CanonicalizationMethod";
        public static readonly XName DigestMethod = w3digsig + "DigestMethod";
        public static readonly XName DigestValue = w3digsig + "DigestValue";
        public static readonly XName Exponent = w3digsig + "Exponent";
        public static readonly XName KeyInfo = w3digsig + "KeyInfo";
        public static readonly XName KeyValue = w3digsig + "KeyValue";
        public static readonly XName Manifest = w3digsig + "Manifest";
        public static readonly XName Modulus = w3digsig + "Modulus";
        public static readonly XName Object = w3digsig + "Object";
        public static readonly XName Reference = w3digsig + "Reference";
        public static readonly XName RSAKeyValue = w3digsig + "RSAKeyValue";
        public static readonly XName Signature = w3digsig + "Signature";
        public static readonly XName SignatureMethod = w3digsig + "SignatureMethod";
        public static readonly XName SignatureProperties = w3digsig + "SignatureProperties";
        public static readonly XName SignatureProperty = w3digsig + "SignatureProperty";
        public static readonly XName SignatureValue = w3digsig + "SignatureValue";
        public static readonly XName SignedInfo = w3digsig + "SignedInfo";
        public static readonly XName Transform = w3digsig + "Transform";
        public static readonly XName Transforms = w3digsig + "Transforms";
        public static readonly XName X509Certificate = w3digsig + "X509Certificate";
        public static readonly XName X509Data = w3digsig + "X509Data";
        public static readonly XName X509IssuerName = w3digsig + "X509IssuerName";
        public static readonly XName X509SerialNumber = w3digsig + "X509SerialNumber";
    }

    public static class WP
    {
        public static readonly XNamespace wp =
            "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        public static readonly XName align = wp + "align";
        public static readonly XName anchor = wp + "anchor";
        public static readonly XName cNvGraphicFramePr = wp + "cNvGraphicFramePr";
        public static readonly XName docPr = wp + "docPr";
        public static readonly XName effectExtent = wp + "effectExtent";
        public static readonly XName extent = wp + "extent";
        public static readonly XName inline = wp + "inline";
        public static readonly XName lineTo = wp + "lineTo";
        public static readonly XName positionH = wp + "positionH";
        public static readonly XName positionV = wp + "positionV";
        public static readonly XName posOffset = wp + "posOffset";
        public static readonly XName simplePos = wp + "simplePos";
        public static readonly XName start = wp + "start";
        public static readonly XName wrapNone = wp + "wrapNone";
        public static readonly XName wrapPolygon = wp + "wrapPolygon";
        public static readonly XName wrapSquare = wp + "wrapSquare";
        public static readonly XName wrapThrough = wp + "wrapThrough";
        public static readonly XName wrapTight = wp + "wrapTight";
        public static readonly XName wrapTopAndBottom = wp + "wrapTopAndBottom";
    }

    public static class WP14
    {
        public static readonly XNamespace wp14 =
            "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";
        public static readonly XName anchorId = wp14 + "anchorId";
        public static readonly XName editId = wp14 + "editId";
        public static readonly XName pctHeight = wp14 + "pctHeight";
        public static readonly XName pctPosVOffset = wp14 + "pctPosVOffset";
        public static readonly XName pctWidth = wp14 + "pctWidth";
        public static readonly XName sizeRelH = wp14 + "sizeRelH";
        public static readonly XName sizeRelV = wp14 + "sizeRelV";
    }

    public static class WPS
    {
        public static readonly XNamespace wps =
            "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
        public static readonly XName altTxbx = wps + "altTxbx";
        public static readonly XName bodyPr = wps + "bodyPr";
        public static readonly XName cNvSpPr = wps + "cNvSpPr";
        public static readonly XName spPr = wps + "spPr";
        public static readonly XName style = wps + "style";
        public static readonly XName textbox = wps + "textbox";
        public static readonly XName txbx = wps + "txbx";
        public static readonly XName wsp = wps + "wsp";
    }

    public static class X
    {
        public static readonly XNamespace x =
            "urn:schemas-microsoft-com:office:excel";
        public static readonly XName Anchor = x + "Anchor";
        public static readonly XName AutoFill = x + "AutoFill";
        public static readonly XName ClientData = x + "ClientData";
        public static readonly XName Column = x + "Column";
        public static readonly XName MoveWithCells = x + "MoveWithCells";
        public static readonly XName Row = x + "Row";
        public static readonly XName SizeWithCells = x + "SizeWithCells";
    }

    public static class XDR
    {
        public static readonly XNamespace xdr =
            "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
        public static readonly XName absoluteAnchor = xdr + "absoluteAnchor";
        public static readonly XName blipFill = xdr + "blipFill";
        public static readonly XName clientData = xdr + "clientData";
        public static readonly XName cNvCxnSpPr = xdr + "cNvCxnSpPr";
        public static readonly XName cNvGraphicFramePr = xdr + "cNvGraphicFramePr";
        public static readonly XName cNvGrpSpPr = xdr + "cNvGrpSpPr";
        public static readonly XName cNvPicPr = xdr + "cNvPicPr";
        public static readonly XName cNvPr = xdr + "cNvPr";
        public static readonly XName cNvSpPr = xdr + "cNvSpPr";
        public static readonly XName col = xdr + "col";
        public static readonly XName colOff = xdr + "colOff";
        public static readonly XName contentPart = xdr + "contentPart";
        public static readonly XName cxnSp = xdr + "cxnSp";
        public static readonly XName ext = xdr + "ext";
        public static readonly XName from = xdr + "from";
        public static readonly XName graphicFrame = xdr + "graphicFrame";
        public static readonly XName grpSp = xdr + "grpSp";
        public static readonly XName grpSpPr = xdr + "grpSpPr";
        public static readonly XName nvCxnSpPr = xdr + "nvCxnSpPr";
        public static readonly XName nvGraphicFramePr = xdr + "nvGraphicFramePr";
        public static readonly XName nvGrpSpPr = xdr + "nvGrpSpPr";
        public static readonly XName nvPicPr = xdr + "nvPicPr";
        public static readonly XName nvSpPr = xdr + "nvSpPr";
        public static readonly XName oneCellAnchor = xdr + "oneCellAnchor";
        public static readonly XName pic = xdr + "pic";
        public static readonly XName pos = xdr + "pos";
        public static readonly XName row = xdr + "row";
        public static readonly XName rowOff = xdr + "rowOff";
        public static readonly XName sp = xdr + "sp";
        public static readonly XName spPr = xdr + "spPr";
        public static readonly XName style = xdr + "style";
        public static readonly XName to = xdr + "to";
        public static readonly XName twoCellAnchor = xdr + "twoCellAnchor";
        public static readonly XName txBody = xdr + "txBody";
        public static readonly XName wsDr = xdr + "wsDr";
        public static readonly XName xfrm = xdr + "xfrm";
    }

    public static class XDR14
    {
        public static readonly XNamespace xdr14 =
            "http://schemas.microsoft.com/office/excel/2010/spreadsheetDrawing";
        public static readonly XName cNvContentPartPr = xdr14 + "cNvContentPartPr";
        public static readonly XName cNvPr = xdr14 + "cNvPr";
        public static readonly XName nvContentPartPr = xdr14 + "nvContentPartPr";
        public static readonly XName nvPr = xdr14 + "nvPr";
        public static readonly XName xfrm = xdr14 + "xfrm";
    }

    public static class XM
    {
        public static readonly XNamespace xm =
            "http://schemas.microsoft.com/office/excel/2006/main";
        public static readonly XName f = xm + "f";
        public static readonly XName _ref = xm + "ref";
        public static readonly XName sqref = xm + "sqref";
    }

    public static class XSI
    {
        public static readonly XNamespace xsi =
            "http://www.w3.org/2001/XMLSchema-instance";
        public static readonly XName type = xsi + "type";
    }


    /************************************* end generated classes *************************************/

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

    public static class Xhtml
    {
        public static readonly XNamespace xhtml = "http://www.w3.org/1999/xhtml";
        public static readonly XName a = xhtml + "a";
        public static readonly XName b = xhtml + "b";
        public static readonly XName body = xhtml + "body";
        public static readonly XName br = xhtml + "br";
        public static readonly XName div = xhtml + "div";
        public static readonly XName h1 = xhtml + "h1";
        public static readonly XName h2 = xhtml + "h2";
        public static readonly XName head = xhtml + "head";
        public static readonly XName html = xhtml + "html";
        public static readonly XName i = xhtml + "i";
        public static readonly XName img = xhtml + "img";
        public static readonly XName meta = xhtml + "meta";
        public static readonly XName p = xhtml + "p";
        public static readonly XName s = xhtml + "s";
        public static readonly XName span = xhtml + "span";
        public static readonly XName style = xhtml + "style";
        public static readonly XName sub = xhtml + "sub";
        public static readonly XName sup = xhtml + "sup";
        public static readonly XName table = xhtml + "table";
        public static readonly XName td = xhtml + "td";
        public static readonly XName title = xhtml + "title";
        public static readonly XName tr = xhtml + "tr";
        public static readonly XName u = xhtml + "u";
    }

    public static class XhtmlNoNamespace
    {
        public static XNamespace xhtml = XNamespace.None;
        public static XName html = xhtml + "html";
        public static XName head = xhtml + "head";
        public static XName title = xhtml + "title";
        public static XName _class = xhtml + "class";
        public static XName colspan = xhtml + "colspan";
        public static XName caption = xhtml + "caption";
        public static XName body = xhtml + "body";
        public static XName div = xhtml + "div";
        public static XName p = xhtml + "p";
        public static XName h1 = xhtml + "h1";
        public static XName h2 = xhtml + "h2";
        public static XName h3 = xhtml + "h3";
        public static XName h4 = xhtml + "h4";
        public static XName h5 = xhtml + "h5";
        public static XName h6 = xhtml + "h6";
        public static XName h7 = xhtml + "h7";
        public static XName h8 = xhtml + "h8";
        public static XName h9 = xhtml + "h9";
        public static XName hr = xhtml + "hr";
        public static XName a = xhtml + "a";
        public static XName b = xhtml + "b";
        public static XName i = xhtml + "i";
        public static XName table = xhtml + "table";
        public static XName th = xhtml + "th";
        public static XName tr = xhtml + "tr";
        public static XName td = xhtml + "td";
        public static XName meta = xhtml + "meta";
        public static XName style = xhtml + "style";
        public static XName br = xhtml + "br";
        public static XName img = xhtml + "img";
        public static XName span = xhtml + "span";
        public static XName href = "href";
        public static XName border = "border";
        public static XName http_equiv = "http-equiv";
        public static XName content = "content";
        public static XName name = "name";
        public static XName width = "width";
        public static XName height = "height";
        public static XName src = "src";
        public static XName alt = "alt";
        public static XName id = "id";
        public static XName descr = "descr";
        public static XName blockquote = "blockquote";
        public static XName type = "type";
        public static XName sub = "sub";
        public static XName sup = "sup";
        public static XName ol = "ol";
        public static XName ul = "ul";
        public static XName li = "li";
        public static XName strong = "Bold";
        public static XName em = "em";
        public static XName tbody = "tbody";
        public static XName valign = "valign";
        public static XName dir = "dir";
        public static XName u = "u";
        public static XName s = "s";
        public static XName rowspan = "rowspan";
        public static XName tt = "tt";
        public static XName code = "code";
        public static XName kbd = "kbd";
        public static XName samp = "samp";
        public static XName pre = "pre";
    }

    public class InvalidOpenXmlDocumentException : Exception
    {
        public InvalidOpenXmlDocumentException(string message) : base(message) { }
    }

    public class OpenXmlPowerToolsException : Exception
    {
        public OpenXmlPowerToolsException(string message) : base(message) { }
    }

    public class ColumnReferenceOutOfRange : Exception
    {
        public ColumnReferenceOutOfRange(string columnReference)
            : base(string.Format("Column reference ({0}) is out of range.", columnReference))
        {
        }
    }

    public class WorksheetAlreadyExistsException : Exception
    {
        public WorksheetAlreadyExistsException(string sheetName)
            : base(string.Format("The worksheet ({0}) already exists.", sheetName))
        {
        }
    }
}
