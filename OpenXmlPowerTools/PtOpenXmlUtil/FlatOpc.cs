using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class FlatOpc
    {
        private static XElement GetContentsAsXml(PackagePart part)
        {
            XNamespace pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";

            if (part.ContentType.EndsWith("xml"))
            {
                using (Stream str = part.GetStream())
                using (var streamReader = new StreamReader(str))
                using (XmlReader xr = XmlReader.Create(streamReader))
                    return new XElement(pkg + "part",
                        new XAttribute(pkg + "name", part.Uri),
                        new XAttribute(pkg + "contentType", part.ContentType),
                        new XElement(pkg + "xmlData",
                            XElement.Load(xr)
                        )
                    );
            }

            using (Stream str = part.GetStream())
            using (var binaryReader = new BinaryReader(str))
            {
                var len = (int) binaryReader.BaseStream.Length;
                byte[] byteArray = binaryReader.ReadBytes(len);

                // the following expression creates the base64String, then chunks
                // it to lines of 76 characters long
                string base64String = Convert.ToBase64String(byteArray)
                    .Select
                    (
                        (c, i) => new FlatOpcTupple
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

                var declaration = new XDeclaration("1.0", "UTF-8", "yes");
                var doc = new XDocument(
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

                var declaration = new XDeclaration("1.0", "UTF-8", "yes");
                var doc = new XDocument(
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

                var declaration = new XDeclaration("1.0", "UTF-8", "yes");
                var doc = new XDocument(
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
                var xmlDoc = new XmlDocument();
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
            var xDoc = new XDocument();
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
                foreach (XElement xmlPart in doc.Root
                    .Elements()
                    .Where(p =>
                        (string) p.Attribute(pkg + "contentType") !=
                        "application/vnd.openxmlformats-package.relationships+xml"))
                {
                    var name = (string) xmlPart.Attribute(pkg + "name");
                    var contentType = (string) xmlPart.Attribute(pkg + "contentType");
                    if (contentType.EndsWith("xml"))
                    {
                        var u = new Uri(name, UriKind.Relative);
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
                        var u = new Uri(name, UriKind.Relative);
                        PackagePart part = package.CreatePart(u, contentType,
                            CompressionOption.SuperFast);
                        using (Stream str = part.GetStream(FileMode.Create))
                        using (var binaryWriter = new BinaryWriter(str))
                        {
                            var base64StringInChunks =
                                (string) xmlPart.Element(pkg + "binaryData");
                            char[] base64CharArray = base64StringInChunks
                                .Where(c => c != '\r' && c != '\n').ToArray();
                            byte[] byteArray =
                                Convert.FromBase64CharArray(base64CharArray,
                                    0, base64CharArray.Length);
                            binaryWriter.Write(byteArray);
                        }
                    }
                }

                foreach (XElement xmlPart in doc.Root.Elements())
                {
                    var name = (string) xmlPart.Attribute(pkg + "name");
                    var contentType = (string) xmlPart.Attribute(pkg + "contentType");
                    if (contentType ==
                        "application/vnd.openxmlformats-package.relationships+xml")
                    {
                        // add the package level relationships
                        if (name == "/_rels/.rels")
                        {
                            foreach (XElement xmlRel in
                                xmlPart.Descendants(rel + "Relationship"))
                            {
                                var id = (string) xmlRel.Attribute("Id");
                                var type = (string) xmlRel.Attribute("Type");
                                var target = (string) xmlRel.Attribute("Target");
                                var targetMode =
                                    (string) xmlRel.Attribute("TargetMode");
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
                                var id = (string) xmlRel.Attribute("Id");
                                var type = (string) xmlRel.Attribute("Type");
                                var target = (string) xmlRel.Attribute("Target");
                                var targetMode =
                                    (string) xmlRel.Attribute("TargetMode");
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

        private class FlatOpcTupple
        {
            public char FoCharacter;
            public int FoChunk;
        }
    }
}
