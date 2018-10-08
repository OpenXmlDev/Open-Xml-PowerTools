using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    public static class PtOpenXmlExtensions
    {
        public static XDocument GetXDocument(this OpenXmlPart part)
        {
            if (part == null) throw new ArgumentNullException(nameof(part));

            var partXDocument = part.Annotation<XDocument>();
            if (partXDocument != null) return partXDocument;

            using (Stream partStream = part.GetStream())
            {
                if (partStream.Length == 0)
                {
                    partXDocument = new XDocument { Declaration = new XDeclaration("1.0", "UTF-8", "yes") };
                }
                else
                {
                    using (XmlReader partXmlReader = XmlReader.Create(partStream))
                    {
                        partXDocument = XDocument.Load(partXmlReader);
                    }
                }
            }

            part.AddAnnotation(partXDocument);
            return partXDocument;
        }

        public static XDocument GetXDocument(this OpenXmlPart part, out XmlNamespaceManager namespaceManager)
        {
            if (part == null) throw new ArgumentNullException(nameof(part));

            namespaceManager = part.Annotation<XmlNamespaceManager>();
            var partXDocument = part.Annotation<XDocument>();
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
                    partXDocument = new XDocument { Declaration = new XDeclaration("1.0", "UTF-8", "yes") };
                    part.AddAnnotation(partXDocument);

                    return partXDocument;
                }

                using (XmlReader partXmlReader = XmlReader.Create(partStream))
                {
                    partXDocument = XDocument.Load(partXmlReader);
                    XmlNameTable nameTable = partXmlReader.NameTable ?? throw new Exception("NameTable is null.");
                    namespaceManager = new XmlNamespaceManager(nameTable);

                    part.AddAnnotation(partXDocument);
                    part.AddAnnotation(namespaceManager);

                    return partXDocument;
                }
            }
        }

        /// <summary>
        /// Gets the given <see cref="OpenXmlPart" />'s root <see cref="XElement" />.
        /// </summary>
        /// <param name="part">The <see cref="OpenXmlPart" />.</param>
        /// <returns>The root <see cref="XElement" />.</returns>
        public static XElement GetXElement(this OpenXmlPart part)
        {
            if (part == null) throw new ArgumentNullException(nameof(part));

            return part.GetXDocument().Root ?? throw new ArgumentException("Part does not contain a root element.");
        }

        /// <summary>
        /// Saves the cached <see cref="XDocument"/> to the the given <see cref="OpenXmlPart"/>.
        /// </summary>
        /// <param name="part">The <see cref="OpenXmlPart"/>.</param>
        public static void PutXDocument(this OpenXmlPart part)
        {
            if (part == null) throw new ArgumentNullException(nameof(part));

            XDocument partXDocument = part.GetXDocument();
            if (partXDocument != null)
            {
                using (Stream partStream = part.GetStream(FileMode.Create, FileAccess.Write))
                using (XmlWriter partXmlWriter = XmlWriter.Create(partStream))
                {
                    partXDocument.Save(partXmlWriter);
                }
            }
        }

        /// <summary>
        /// Saves the cached <see cref="XDocument"/> to the the given <see cref="OpenXmlPart"/>,
        /// indending the XML markup and creating new lines for attributes.
        /// </summary>
        /// <param name="part">The <see cref="OpenXmlPart"/>.</param>
        public static void PutXDocumentWithFormatting(this OpenXmlPart part)
        {
            if (part == null) throw new ArgumentNullException(nameof(part));

            XDocument partXDocument = part.GetXDocument();
            if (partXDocument != null)
            {
                using (Stream partStream = part.GetStream(FileMode.Create, FileAccess.Write))
                {
                    var settings = new XmlWriterSettings
                    {
                        Indent = true,
                        OmitXmlDeclaration = true,
                        NewLineOnAttributes = true
                    };

                    using (XmlWriter partXmlWriter = XmlWriter.Create(partStream, settings))
                    {
                        partXDocument.Save(partXmlWriter);
                    }
                }
            }
        }

        public static void PutXDocument(this OpenXmlPart part, XDocument document)
        {
            if (part == null) throw new ArgumentNullException(nameof(part));
            if (document == null) throw new ArgumentNullException(nameof(document));

            using (Stream partStream = part.GetStream(FileMode.Create, FileAccess.Write))
            using (XmlWriter partXmlWriter = XmlWriter.Create(partStream))
            {
                document.Save(partXmlWriter);
            }

            part.RemoveAnnotations<XDocument>();
            part.AddAnnotation(document);
        }

        /// <summary>
        /// Writes the cached root <see cref="XElement" /> to the given <see cref="OpenXmlPart" />.
        /// </summary>
        /// <param name="part">The <see cref="OpenXmlPart"/>.</param>
        public static void PutXElement(this OpenXmlPart part)
        {
            if (part == null) throw new ArgumentNullException(nameof(part));

            part.PutXDocument();
        }

        /// <summary>
        /// Writes the given root <see cref="XElement" /> to the given <see cref="OpenXmlPart" />.
        /// </summary>
        /// <param name="part">The <see cref="OpenXmlPart" />.</param>
        /// <param name="root">The root <see cref="XElement" />.</param>
        public static void PutXElement(this OpenXmlPart part, XElement root)
        {
            if (root == null) throw new ArgumentNullException(nameof(root));

            PutXDocument(part, new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root));
        }

        private static XmlNamespaceManager GetManagerFromXDocument(XDocument xDocument)
        {
            XmlReader reader = xDocument.CreateReader();
            XDocument newXDoc = XDocument.Load(reader);

            XElement rootElement = xDocument.Elements().First();
            rootElement.ReplaceWith(newXDoc.Root);

            XmlNameTable nameTable = reader.NameTable ?? throw new Exception("NameTable is null.");
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

            foreach (HeaderPart hdr in doc.MainDocumentPart.HeaderParts)
                yield return hdr;

            foreach (FooterPart ftr in doc.MainDocumentPart.FooterParts)
                yield return ftr;

            if (doc.MainDocumentPart.FootnotesPart != null)
                yield return doc.MainDocumentPart.FootnotesPart;

            if (doc.MainDocumentPart.EndnotesPart != null)
                yield return doc.MainDocumentPart.EndnotesPart;
        }

        /// <summary>
        /// Creates a complete list of all parts contained in the <see cref="OpenXmlPartContainer" />.
        /// </summary>
        /// <param name="container">
        /// A <see cref="WordprocessingDocument" />, <see cref="SpreadsheetDocument" />, or
        /// <see cref="PresentationDocument" />.
        /// </param>
        /// <returns>list of <see cref="OpenXmlPart" />s contained in the <see cref="OpenXmlPartContainer" />.</returns>
        public static List<OpenXmlPart> GetAllParts(this OpenXmlPartContainer container)
        {
            // Use a HashSet so that parts are processed only once.
            var partList = new HashSet<OpenXmlPart>();

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

        public static void IgnoreNamespace(this XElement root, string prefix, XNamespace @namespace)
        {
            // Declare markup compatibility extensions namespace as necessary.
            if (root.Attributes().All(a => a.Value != MC.mc.NamespaceName))
            {
                root.Add(new XAttribute(XNamespace.Xmlns + "mc", MC.mc.NamespaceName));
            }

            // Declare ignored namespace as necessary.
            string namespaceName = @namespace.NamespaceName;
            bool IsIgnoredNamespaceDeclaration(XAttribute a) => a.Name.Namespace == XNamespace.Xmlns && a.Value == namespaceName;
            XAttribute attribute = root.Attributes().FirstOrDefault(IsIgnoredNamespaceDeclaration);
            if (attribute == null)
            {
                attribute = new XAttribute(XNamespace.Xmlns + prefix, namespaceName);
                root.Add(attribute);
            }

            string effectivePrefix = root.GetPrefixOfNamespace(@namespace);

            // Add prefix to mc:Ignorable attribute value.
            var ignorable = (string) root.Attribute(MC.Ignorable);
            if (ignorable != null)
            {
                string[] list = ignorable.Split(' ');
                if (!list.Contains(effectivePrefix))
                {
                    root.SetAttributeValue(MC.Ignorable, ignorable + " " + effectivePrefix);
                }
            }
            else
            {
                root.Add(new XAttribute(MC.Ignorable, effectivePrefix));
            }
        }
    }
}
