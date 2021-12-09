using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    public static class PtOpenXmlExtensions
    {
        public static XDocument GetXDocument(this OpenXmlPart part, out XmlNamespaceManager namespaceManager)
        {
            if (part == null) throw new ArgumentNullException(nameof(part));

            XDocument partXDocument = part.GetXDocument();

            namespaceManager = partXDocument.Annotation<XmlNamespaceManager>();
            if (namespaceManager != null)
            {
                return partXDocument;
            }

            namespaceManager = CreateXmlNamespaceManager(partXDocument.Root);
            partXDocument.AddAnnotation(namespaceManager);

            return partXDocument;
        }

        private static XmlNamespaceManager CreateXmlNamespaceManager(XElement root)
        {
            var namespaceManager = new XmlNamespaceManager(new NameTable());

            if (root == null)
            {
                return namespaceManager;
            }

            foreach (XAttribute declaration in root.Attributes().Where(a => a.IsNamespaceDeclaration))
            {
                namespaceManager.AddNamespace(declaration.Name.LocalName, declaration.Value);
            }

            return namespaceManager;
        }

        /// <summary>
        /// Saves the cached <see cref="XDocument"/> to the the given <see cref="OpenXmlPart"/>.
        /// </summary>
        /// <param name="part">The <see cref="OpenXmlPart"/>.</param>
        [Obsolete("Use SaveXDocument(OpenXmlPart) instead.")]
        public static void PutXDocument(this OpenXmlPart part)
        {
            if (part == null) throw new ArgumentNullException(nameof(part));

            part.SaveXDocument();
        }

        /// <summary>
        /// Saves the given <see cref="XDocument"/> to the the given <see cref="OpenXmlPart"/>.
        /// </summary>
        /// <param name="part">The <see cref="OpenXmlPart"/>.</param>
        /// <param name="document">The <see cref="XDocument"/>.</param>
        [Obsolete("Use SetXDocument(OpenXmlPart, XDocument) instead.")]
        public static void PutXDocument(this OpenXmlPart part, XDocument document)
        {
            if (part == null) throw new ArgumentNullException(nameof(part));

            part.SetXDocument(document);
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
            return source.SelectMany(e1 => e1.LogicalChildrenContent());
        }

        public static IEnumerable<XElement> LogicalChildrenContent(this XElement element, XName name)
        {
            return element.LogicalChildrenContent().Where(e => e.Name == name);
        }

        public static IEnumerable<XElement> LogicalChildrenContent(this IEnumerable<XElement> source, XName name)
        {
            return source.SelectMany(e1 => e1.LogicalChildrenContent(name));
        }

        public static IEnumerable<OpenXmlPart> ContentParts(this WordprocessingDocument doc)
        {
            if (doc.MainDocumentPart == null)
            {
                yield break;
            }

            yield return doc.MainDocumentPart;

            foreach (HeaderPart hdr in doc.MainDocumentPart.HeaderParts)
            {
                yield return hdr;
            }

            foreach (FooterPart ftr in doc.MainDocumentPart.FooterParts)
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
            {
                AddPart(partList, p.OpenXmlPart);
            }

            return partList.OrderBy(p => p.ContentType).ThenBy(p => p.Uri.ToString()).ToList();
        }

        private static void AddPart(HashSet<OpenXmlPart> partList, OpenXmlPart part)
        {
            if (partList.Contains(part)) return;

            partList.Add(part);
            foreach (IdPartPair p in part.Parts)
            {
                AddPart(partList, p.OpenXmlPart);
            }
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
