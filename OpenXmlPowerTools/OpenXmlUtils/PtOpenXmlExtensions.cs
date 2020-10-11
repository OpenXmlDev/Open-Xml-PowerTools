using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class PtOpenXmlExtensions
    {
        public static XDocument GetXDocument(this OpenXmlPart part)
        {
            if (part == null)
            {
                throw new ArgumentNullException("part");
            }

            var partXDocument = part.Annotation<XDocument>();
            if (partXDocument != null)
            {
                return partXDocument;
            }

            using (var partStream = part.GetStream())
            {
                if (partStream.Length == 0)
                {
                    partXDocument = new XDocument
                    {
                        Declaration = new XDeclaration("1.0", "UTF-8", "yes")
                    };
                }
                else
                {
                    using var partXmlReader = XmlReader.Create(partStream);
                    partXDocument = XDocument.Load(partXmlReader);
                }
            }

            part.AddAnnotation(partXDocument);
            return partXDocument;
        }

        public static XDocument GetXDocument(this OpenXmlPart part, out XmlNamespaceManager namespaceManager)
        {
            if (part == null)
            {
                throw new ArgumentNullException("part");
            }

            namespaceManager = part.Annotation<XmlNamespaceManager>();
            var partXDocument = part.Annotation<XDocument>();
            if (partXDocument != null)
            {
                if (namespaceManager != null)
                {
                    return partXDocument;
                }

                namespaceManager = GetManagerFromXDocument(partXDocument);
                part.AddAnnotation(namespaceManager);

                return partXDocument;
            }

            using var partStream = part.GetStream();
            if (partStream.Length == 0)
            {
                partXDocument = new XDocument
                {
                    Declaration = new XDeclaration("1.0", "UTF-8", "yes")
                };

                part.AddAnnotation(partXDocument);

                return partXDocument;
            }
            else
            {
                using var partXmlReader = XmlReader.Create(partStream);
                partXDocument = XDocument.Load(partXmlReader);
                namespaceManager = new XmlNamespaceManager(partXmlReader.NameTable);

                part.AddAnnotation(partXDocument);
                part.AddAnnotation(namespaceManager);

                return partXDocument;
            }
        }

        public static void PutXDocument(this OpenXmlPart part)
        {
            if (part == null)
            {
                throw new ArgumentNullException("part");
            }

            var partXDocument = part.GetXDocument();
            if (partXDocument != null)
            {
                using var partStream = part.GetStream(FileMode.Create, FileAccess.Write);
                using var partXmlWriter = XmlWriter.Create(partStream);
                partXDocument.Save(partXmlWriter);
            }
        }

        public static void PutXDocumentWithFormatting(this OpenXmlPart part)
        {
            if (part == null)
            {
                throw new ArgumentNullException("part");
            }

            var partXDocument = part.GetXDocument();
            if (partXDocument != null)
            {
                using var partStream = part.GetStream(FileMode.Create, FileAccess.Write);
                var settings = new XmlWriterSettings
                {
                    Indent = true,
                    OmitXmlDeclaration = true,
                    NewLineOnAttributes = true
                };
                using var partXmlWriter = XmlWriter.Create(partStream, settings);
                partXDocument.Save(partXmlWriter);
            }
        }

        public static void PutXDocument(this OpenXmlPart part, XDocument document)
        {
            if (part == null)
            {
                throw new ArgumentNullException("part");
            }

            if (document == null)
            {
                throw new ArgumentNullException("document");
            }

            using (var partStream = part.GetStream(FileMode.Create, FileAccess.Write))
            using (var partXmlWriter = XmlWriter.Create(partStream))
            {
                document.Save(partXmlWriter);
            }

            part.RemoveAnnotations<XDocument>();
            part.AddAnnotation(document);
        }

        private static XmlNamespaceManager GetManagerFromXDocument(XDocument xDocument)
        {
            var reader = xDocument.CreateReader();
            var newXDoc = XDocument.Load(reader);

            var rootElement = xDocument.Elements().FirstOrDefault();
            rootElement.ReplaceWith(newXDoc.Root);

            var nameTable = reader.NameTable;
            var namespaceManager = new XmlNamespaceManager(nameTable);
            return namespaceManager;
        }

        public static IEnumerable<XElement> LogicalChildrenContent(this XElement element)
        {
            if (element.Name == W.document)
            {
                return element.Descendants(W.body).Take(1);
            }

            if (element.Name == W.body ||
                element.Name == W.tc ||
                element.Name == W.txbxContent)
            {
                return element
                    .DescendantsTrimmed(e =>
                        e.Name == W.tbl ||
                        e.Name == W.p)
                    .Where(e =>
                        e.Name == W.p ||
                        e.Name == W.tbl);
            }

            if (element.Name == W.tbl)
            {
                return element
                    .DescendantsTrimmed(W.tr)
                    .Where(e => e.Name == W.tr);
            }

            if (element.Name == W.tr)
            {
                return element
                    .DescendantsTrimmed(W.tc)
                    .Where(e => e.Name == W.tc);
            }

            if (element.Name == W.p)
            {
                return element
                    .DescendantsTrimmed(e => e.Name == W.r ||
                        e.Name == W.pict ||
                        e.Name == W.drawing)
                    .Where(e => e.Name == W.r ||
                        e.Name == W.pict ||
                        e.Name == W.drawing);
            }

            if (element.Name == W.r)
            {
                return element
                    .DescendantsTrimmed(e => W.SubRunLevelContent.Contains(e.Name))
                    .Where(e => W.SubRunLevelContent.Contains(e.Name));
            }

            if (element.Name == MC.AlternateContent)
            {
                return element
                    .DescendantsTrimmed(e =>
                        e.Name == W.pict ||
                        e.Name == W.drawing ||
                        e.Name == MC.Fallback)
                    .Where(e =>
                        e.Name == W.pict ||
                        e.Name == W.drawing);
            }

            if (element.Name == W.pict || element.Name == W.drawing)
            {
                return element
                    .DescendantsTrimmed(W.txbxContent)
                    .Where(e => e.Name == W.txbxContent);
            }

            return XElement.EmptySequence;
        }

        public static IEnumerable<XElement> LogicalChildrenContent(this IEnumerable<XElement> source)
        {
            foreach (var e1 in source)
            {
                foreach (var e2 in e1.LogicalChildrenContent())
                {
                    yield return e2;
                }
            }
        }

        public static IEnumerable<XElement> LogicalChildrenContent(this XElement element, XName name)
        {
            return element.LogicalChildrenContent().Where(e => e.Name == name);
        }

        public static IEnumerable<XElement> LogicalChildrenContent(this IEnumerable<XElement> source, XName name)
        {
            foreach (var e1 in source)
            {
                foreach (var e2 in e1.LogicalChildrenContent(name))
                {
                    yield return e2;
                }
            }
        }

        public static IEnumerable<OpenXmlPart> ContentParts(this WordprocessingDocument doc)
        {
            yield return doc.MainDocumentPart;

            foreach (var hdr in doc.MainDocumentPart.HeaderParts)
            {
                yield return hdr;
            }

            foreach (var ftr in doc.MainDocumentPart.FooterParts)
            {
                yield return ftr;
            }

            if (doc.MainDocumentPart.FootnotesPart != null)
            {
                yield return doc.MainDocumentPart.FootnotesPart;
            }

            if (doc.MainDocumentPart.EndnotesPart != null)
            {
                yield return doc.MainDocumentPart.EndnotesPart;
            }
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
            var partList = new HashSet<OpenXmlPart>();

            foreach (var p in container.Parts)
            {
                AddPart(partList, p.OpenXmlPart);
            }

            return partList.OrderBy(p => p.ContentType).ThenBy(p => p.Uri.ToString()).ToList();
        }

        private static void AddPart(HashSet<OpenXmlPart> partList, OpenXmlPart part)
        {
            if (partList.Contains(part))
            {
                return;
            }

            partList.Add(part);
            foreach (var p in part.Parts)
            {
                AddPart(partList, p.OpenXmlPart);
            }
        }
    }
}