// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

/*
Here is modification of a WmlDocument:
    public static WmlDocument SimplifyMarkup(WmlDocument doc, SimplifyMarkupSettings settings)
    {
        using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(doc))
        {
            using (WordprocessingDocument document = streamDoc.GetWordprocessingDocument())
            {
                SimplifyMarkup(document, settings);
            }
            return streamDoc.GetModifiedWmlDocument();
        }
    }

Here is read-only of a WmlDocument:

    public static string GetBackgroundColor(WmlDocument doc)
    {
        using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(doc))
        using (WordprocessingDocument document = streamDoc.GetWordprocessingDocument())
        {
            XDocument mainDocument = document.MainDocumentPart.GetXDocument();
            XElement backgroundElement = mainDocument.Descendants(W.background).FirstOrDefault();
            return (backgroundElement == null) ? string.Empty : backgroundElement.Attribute(W.color).Value;
        }
    }

Here is creating a new WmlDocument:

    private OpenXmlPowerToolsDocument CreateSplitDocument(WordprocessingDocument source, List<XElement> contents, string newFileName)
    {
        using (OpenXmlMemoryStreamDocument streamDoc = OpenXmlMemoryStreamDocument.CreateWordprocessingDocument())
        {
            using (WordprocessingDocument document = streamDoc.GetWordprocessingDocument())
            {
                DocumentBuilder.FixRanges(source.MainDocumentPart.GetXDocument(), contents);
                PowerToolsExtensions.SetContent(document, contents);
            }
            OpenXmlPowerToolsDocument newDoc = streamDoc.GetModifiedDocument();
            newDoc.FileName = newFileName;
            return newDoc;
        }
    }
*/

using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    public class PowerToolsDocumentException : Exception
    {
        public PowerToolsDocumentException(string message) : base(message)
        {
        }
    }

    public class PowerToolsInvalidDataException : Exception
    {
        public PowerToolsInvalidDataException(string message) : base(message)
        {
        }
    }

    [SuppressMessage("ReSharper", "MemberCanBeProtected.Global")]
    public class OpenXmlPowerToolsDocument
    {
        public OpenXmlPowerToolsDocument(OpenXmlPowerToolsDocument original)
        {
            DocumentByteArray = new byte[original.DocumentByteArray.Length];
            Array.Copy(original.DocumentByteArray, DocumentByteArray, original.DocumentByteArray.Length);
            FileName = original.FileName;
        }

        public OpenXmlPowerToolsDocument(OpenXmlPowerToolsDocument original, bool convertToTransitional)
        {
            if (convertToTransitional)
            {
                ConvertToTransitional(original.FileName, original.DocumentByteArray);
            }
            else
            {
                DocumentByteArray = new byte[original.DocumentByteArray.Length];
                Array.Copy(original.DocumentByteArray, DocumentByteArray, original.DocumentByteArray.Length);
                FileName = original.FileName;
            }
        }

        public OpenXmlPowerToolsDocument(string fileName)
        {
            FileName = fileName;
            DocumentByteArray = File.ReadAllBytes(fileName);
        }

        public OpenXmlPowerToolsDocument(string fileName, bool convertToTransitional)
        {
            FileName = fileName;

            if (convertToTransitional)
            {
                byte[] tempByteArray = File.ReadAllBytes(fileName);
                ConvertToTransitional(fileName, tempByteArray);
            }
            else
            {
                FileName = fileName;
                DocumentByteArray = File.ReadAllBytes(fileName);
            }
        }

        public OpenXmlPowerToolsDocument(byte[] byteArray)
        {
            DocumentByteArray = new byte[byteArray.Length];
            Array.Copy(byteArray, DocumentByteArray, byteArray.Length);
            FileName = null;
        }

        public OpenXmlPowerToolsDocument(byte[] byteArray, bool convertToTransitional)
        {
            if (convertToTransitional)
            {
                ConvertToTransitional(null, byteArray);
            }
            else
            {
                DocumentByteArray = new byte[byteArray.Length];
                Array.Copy(byteArray, DocumentByteArray, byteArray.Length);
                FileName = null;
            }
        }

        public OpenXmlPowerToolsDocument(string fileName, MemoryStream memStream)
        {
            FileName = fileName;
            DocumentByteArray = new byte[memStream.Length];
            Array.Copy(memStream.GetBuffer(), DocumentByteArray, memStream.Length);
        }

        public OpenXmlPowerToolsDocument(string fileName, MemoryStream memStream, bool convertToTransitional)
        {
            if (convertToTransitional)
            {
                ConvertToTransitional(fileName, memStream.ToArray());
            }
            else
            {
                FileName = fileName;
                DocumentByteArray = new byte[memStream.Length];
                Array.Copy(memStream.GetBuffer(), DocumentByteArray, memStream.Length);
            }
        }

        public string FileName { get; set; }

        public byte[] DocumentByteArray { get; set; }

        public static OpenXmlPowerToolsDocument FromFileName(string fileName)
        {
            byte[] bytes = File.ReadAllBytes(fileName);
            Type type;

            try
            {
                type = GetDocumentType(bytes);
            }
            catch (FileFormatException)
            {
                throw new PowerToolsDocumentException("Not an Open XML document.");
            }

            if (type == typeof(WordprocessingDocument))
            {
                return new WmlDocument(fileName, bytes);
            }

            if (type == typeof(SpreadsheetDocument))
            {
                return new SmlDocument(fileName, bytes);
            }

            if (type == typeof(PresentationDocument))
            {
                return new PmlDocument(fileName, bytes);
            }

            if (type == typeof(Package))
            {
                return new OpenXmlPowerToolsDocument(bytes) { FileName = fileName };
            }

            throw new PowerToolsDocumentException("Not an Open XML document.");
        }

        public static OpenXmlPowerToolsDocument FromDocument(OpenXmlPowerToolsDocument doc)
        {
            Type type = doc.GetDocumentType();
            if (type == typeof(WordprocessingDocument))
            {
                return new WmlDocument(doc);
            }

            if (type == typeof(SpreadsheetDocument))
            {
                return new SmlDocument(doc);
            }

            if (type == typeof(PresentationDocument))
            {
                return new PmlDocument(doc);
            }

            return null; // This should not be possible from a valid OpenXmlPowerToolsDocument object
        }

        private void ConvertToTransitional(string fileName, byte[] tempByteArray)
        {
            Type type;
            try
            {
                type = GetDocumentType(tempByteArray);
            }
            catch (FileFormatException)
            {
                throw new PowerToolsDocumentException("Not an Open XML document.");
            }

            using (var ms = new MemoryStream())
            {
                ms.Write(tempByteArray, 0, tempByteArray.Length);
                if (type == typeof(WordprocessingDocument))
                {
                    using (WordprocessingDocument sDoc = WordprocessingDocument.Open(ms, true))
                    {
                        // following code forces the SDK to serialize
                        foreach (IdPartPair part in sDoc.Parts)
                        {
                            try
                            {
                                OpenXmlPartRootElement unused = part.OpenXmlPart.RootElement;
                            }
                            catch (Exception)
                            {
                                // Ignore
                            }
                        }
                    }
                }
                else if (type == typeof(SpreadsheetDocument))
                {
                    using (SpreadsheetDocument sDoc = SpreadsheetDocument.Open(ms, true))
                    {
                        // following code forces the SDK to serialize
                        foreach (IdPartPair part in sDoc.Parts)
                        {
                            try
                            {
                                OpenXmlPartRootElement unused = part.OpenXmlPart.RootElement;
                            }
                            catch (Exception)
                            {
                                // Ignore
                            }
                        }
                    }
                }
                else if (type == typeof(PresentationDocument))
                {
                    using (PresentationDocument sDoc = PresentationDocument.Open(ms, true))
                    {
                        // following code forces the SDK to serialize
                        foreach (IdPartPair part in sDoc.Parts)
                        {
                            try
                            {
                                OpenXmlPartRootElement unused = part.OpenXmlPart.RootElement;
                            }
                            catch (Exception)
                            {
                                // Ignore
                            }
                        }
                    }
                }

                FileName = fileName;
                DocumentByteArray = ms.ToArray();
            }
        }

        public string GetName()
        {
            if (FileName == null)
                return "Unnamed Document";

            var file = new FileInfo(FileName);
            return file.Name;
        }

        public void SaveAs(string fileName)
        {
            File.WriteAllBytes(fileName, DocumentByteArray);
        }

        public void Save()
        {
            if (FileName == null)
                throw new InvalidOperationException("Attempting to Save a document that has no file name.  Use SaveAs instead.");

            File.WriteAllBytes(FileName, DocumentByteArray);
        }

        public void WriteByteArray(Stream stream)
        {
            stream.Write(DocumentByteArray, 0, DocumentByteArray.Length);
        }

        public Type GetDocumentType()
        {
            return GetDocumentType(DocumentByteArray);
        }

        private static Type GetDocumentType(byte[] bytes)
        {
            // Relationship types:
            const string coreDocument = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
            const string strictCoreDocument = "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument";

            using (var stream = new MemoryStream())
            {
                stream.Write(bytes, 0, bytes.Length);
                using (Package package = Package.Open(stream, FileMode.Open))
                {
                    PackageRelationship relationship =
                        package.GetRelationshipsByType(coreDocument).FirstOrDefault() ??
                        package.GetRelationshipsByType(strictCoreDocument).FirstOrDefault();

                    if (relationship != null)
                    {
                        Uri partUri = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
                        PackagePart part = package.GetPart(partUri);

                        switch (part.ContentType)
                        {
                            case "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml":
                            case "application/vnd.ms-word.document.macroEnabled.main+xml":
                            case "application/vnd.ms-word.template.macroEnabledTemplate.main+xml":
                            case "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml":
                                return typeof(WordprocessingDocument);

                            case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml":
                            case "application/vnd.ms-excel.sheet.macroEnabled.main+xml":
                            case "application/vnd.ms-excel.template.macroEnabled.main+xml":
                            case "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml":
                                return typeof(SpreadsheetDocument);

                            case "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml":
                            case "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml":
                            case "application/vnd.ms-powerpoint.template.macroEnabled.main+xml":
                            case "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml":
                            case "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml":
                            case "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml":
                                return typeof(PresentationDocument);

                            default:
                                return typeof(Package);
                        }
                    }

                    return null;
                }
            }
        }

        public static void SavePartAs(OpenXmlPart part, string filePath)
        {
            Stream partStream = part.GetStream(FileMode.Open, FileAccess.Read);
            var partContent = new byte[partStream.Length];
            partStream.Read(partContent, 0, (int) partStream.Length);

            File.WriteAllBytes(filePath, partContent);
        }
    }

    public partial class WmlDocument : OpenXmlPowerToolsDocument
    {
        public WmlDocument(OpenXmlPowerToolsDocument original)
            : base(original)
        {
            if (GetDocumentType() != typeof(WordprocessingDocument))
                throw new PowerToolsDocumentException("Not a Wordprocessing document.");
        }

        public WmlDocument(OpenXmlPowerToolsDocument original, bool convertToTransitional)
            : base(original, convertToTransitional)
        {
            if (GetDocumentType() != typeof(WordprocessingDocument))
                throw new PowerToolsDocumentException("Not a Wordprocessing document.");
        }

        public WmlDocument(string fileName)
            : base(fileName)
        {
            if (GetDocumentType() != typeof(WordprocessingDocument))
                throw new PowerToolsDocumentException("Not a Wordprocessing document.");
        }

        public WmlDocument(string fileName, bool convertToTransitional)
            : base(fileName, convertToTransitional)
        {
            if (GetDocumentType() != typeof(WordprocessingDocument))
                throw new PowerToolsDocumentException("Not a Wordprocessing document.");
        }

        public WmlDocument(string fileName, byte[] byteArray)
            : base(byteArray)
        {
            FileName = fileName;
            if (GetDocumentType() != typeof(WordprocessingDocument))
                throw new PowerToolsDocumentException("Not a Wordprocessing document.");
        }

        public WmlDocument(string fileName, byte[] byteArray, bool convertToTransitional)
            : base(byteArray, convertToTransitional)
        {
            FileName = fileName;
            if (GetDocumentType() != typeof(WordprocessingDocument))
                throw new PowerToolsDocumentException("Not a Wordprocessing document.");
        }

        public WmlDocument(string fileName, MemoryStream memStream)
            : base(fileName, memStream)
        {
        }

        public WmlDocument(string fileName, MemoryStream memStream, bool convertToTransitional)
            : base(fileName, memStream, convertToTransitional)
        {
        }
    }

    public partial class SmlDocument : OpenXmlPowerToolsDocument
    {
        public SmlDocument(OpenXmlPowerToolsDocument original)
            : base(original)
        {
            if (GetDocumentType() != typeof(SpreadsheetDocument))
                throw new PowerToolsDocumentException("Not a Spreadsheet document.");
        }

        public SmlDocument(OpenXmlPowerToolsDocument original, bool convertToTransitional)
            : base(original, convertToTransitional)
        {
            if (GetDocumentType() != typeof(SpreadsheetDocument))
                throw new PowerToolsDocumentException("Not a Spreadsheet document.");
        }

        public SmlDocument(string fileName)
            : base(fileName)
        {
            if (GetDocumentType() != typeof(SpreadsheetDocument))
                throw new PowerToolsDocumentException("Not a Spreadsheet document.");
        }

        public SmlDocument(string fileName, bool convertToTransitional)
            : base(fileName, convertToTransitional)
        {
            if (GetDocumentType() != typeof(SpreadsheetDocument))
                throw new PowerToolsDocumentException("Not a Spreadsheet document.");
        }

        public SmlDocument(string fileName, byte[] byteArray)
            : base(byteArray)
        {
            FileName = fileName;
            if (GetDocumentType() != typeof(SpreadsheetDocument))
                throw new PowerToolsDocumentException("Not a Spreadsheet document.");
        }

        public SmlDocument(string fileName, byte[] byteArray, bool convertToTransitional)
            : base(byteArray, convertToTransitional)
        {
            FileName = fileName;
            if (GetDocumentType() != typeof(SpreadsheetDocument))
                throw new PowerToolsDocumentException("Not a Spreadsheet document.");
        }

        public SmlDocument(string fileName, MemoryStream memStream)
            : base(fileName, memStream)
        {
        }

        public SmlDocument(string fileName, MemoryStream memStream, bool convertToTransitional)
            : base(fileName, memStream, convertToTransitional)
        {
        }
    }

    public partial class PmlDocument : OpenXmlPowerToolsDocument
    {
        public PmlDocument(OpenXmlPowerToolsDocument original)
            : base(original)
        {
            if (GetDocumentType() != typeof(PresentationDocument))
                throw new PowerToolsDocumentException("Not a Presentation document.");
        }

        public PmlDocument(OpenXmlPowerToolsDocument original, bool convertToTransitional)
            : base(original, convertToTransitional)
        {
            if (GetDocumentType() != typeof(PresentationDocument))
                throw new PowerToolsDocumentException("Not a Presentation document.");
        }

        public PmlDocument(string fileName)
            : base(fileName)
        {
            if (GetDocumentType() != typeof(PresentationDocument))
                throw new PowerToolsDocumentException("Not a Presentation document.");
        }

        public PmlDocument(string fileName, bool convertToTransitional)
            : base(fileName, convertToTransitional)
        {
            if (GetDocumentType() != typeof(PresentationDocument))
                throw new PowerToolsDocumentException("Not a Presentation document.");
        }

        public PmlDocument(string fileName, byte[] byteArray)
            : base(byteArray)
        {
            FileName = fileName;
            if (GetDocumentType() != typeof(PresentationDocument))
                throw new PowerToolsDocumentException("Not a Presentation document.");
        }

        public PmlDocument(string fileName, byte[] byteArray, bool convertToTransitional)
            : base(byteArray, convertToTransitional)
        {
            FileName = fileName;
            if (GetDocumentType() != typeof(PresentationDocument))
                throw new PowerToolsDocumentException("Not a Presentation document.");
        }

        public PmlDocument(string fileName, MemoryStream memStream)
            : base(fileName, memStream)
        {
        }

        public PmlDocument(string fileName, MemoryStream memStream, bool convertToTransitional)
            : base(fileName, memStream, convertToTransitional)
        {
        }
    }

    public class OpenXmlMemoryStreamDocument : IDisposable
    {
        private readonly OpenXmlPowerToolsDocument _document;
        private MemoryStream _docMemoryStream;
        private Package _docPackage;

        public OpenXmlMemoryStreamDocument(OpenXmlPowerToolsDocument doc)
        {
            _document = doc;
            _docMemoryStream = new MemoryStream();
            _docMemoryStream.Write(doc.DocumentByteArray, 0, doc.DocumentByteArray.Length);

            try
            {
                _docPackage = Package.Open(_docMemoryStream, FileMode.Open);
            }
            catch (Exception e)
            {
                throw new PowerToolsDocumentException(e.Message);
            }
        }

        internal OpenXmlMemoryStreamDocument(MemoryStream stream)
        {
            _docMemoryStream = stream;

            try
            {
                _docPackage = Package.Open(_docMemoryStream, FileMode.Open);
            }
            catch (Exception e)
            {
                throw new PowerToolsDocumentException(e.Message);
            }
        }

        public void Dispose()
        {
            Dispose(true);
        }

        public static OpenXmlMemoryStreamDocument CreateWordprocessingDocument()
        {
            var stream = new MemoryStream();
            using (WordprocessingDocument doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                doc.AddMainDocumentPart();
                doc.MainDocumentPart.PutXDocument(new XDocument(
                    new XElement(W.document,
                        new XAttribute(XNamespace.Xmlns + "w", W.w),
                        new XAttribute(XNamespace.Xmlns + "r", R.r),
                        new XElement(W.body))));
                doc.Close();
                return new OpenXmlMemoryStreamDocument(stream);
            }
        }

        public static OpenXmlMemoryStreamDocument CreateSpreadsheetDocument()
        {
            var stream = new MemoryStream();
            using (SpreadsheetDocument doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                doc.AddWorkbookPart();
                XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                XNamespace relationshipsns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                doc.WorkbookPart.PutXDocument(new XDocument(
                    new XElement(ns + "workbook",
                        new XAttribute("xmlns", ns),
                        new XAttribute(XNamespace.Xmlns + "r", relationshipsns),
                        new XElement(ns + "sheets"))));
                doc.Close();
                return new OpenXmlMemoryStreamDocument(stream);
            }
        }

        public static OpenXmlMemoryStreamDocument CreatePresentationDocument()
        {
            var stream = new MemoryStream();
            using (PresentationDocument doc = PresentationDocument.Create(stream, PresentationDocumentType.Presentation))
            {
                doc.AddPresentationPart();
                XNamespace ns = "http://schemas.openxmlformats.org/presentationml/2006/main";
                XNamespace relationshipsns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                XNamespace drawingns = "http://schemas.openxmlformats.org/drawingml/2006/main";
                doc.PresentationPart.PutXDocument(new XDocument(
                    new XElement(ns + "presentation",
                        new XAttribute(XNamespace.Xmlns + "a", drawingns),
                        new XAttribute(XNamespace.Xmlns + "r", relationshipsns),
                        new XAttribute(XNamespace.Xmlns + "p", ns),
                        new XElement(ns + "sldMasterIdLst"),
                        new XElement(ns + "sldIdLst"),
                        new XElement(ns + "notesSz", new XAttribute("cx", "6858000"), new XAttribute("cy", "9144000")))));
                doc.Close();
                return new OpenXmlMemoryStreamDocument(stream);
            }
        }

        public static OpenXmlMemoryStreamDocument CreatePackage()
        {
            var stream = new MemoryStream();
            Package package = Package.Open(stream, FileMode.Create);
            package.Close();
            return new OpenXmlMemoryStreamDocument(stream);
        }

        public Package GetPackage()
        {
            return _docPackage;
        }

        public WordprocessingDocument GetWordprocessingDocument()
        {
            try
            {
                if (GetDocumentType() != typeof(WordprocessingDocument))
                    throw new PowerToolsDocumentException("Not a Wordprocessing document.");

                return WordprocessingDocument.Open(_docPackage);
            }
            catch (Exception e)
            {
                throw new PowerToolsDocumentException(e.Message);
            }
        }

        public SpreadsheetDocument GetSpreadsheetDocument()
        {
            try
            {
                if (GetDocumentType() != typeof(SpreadsheetDocument))
                    throw new PowerToolsDocumentException("Not a Spreadsheet document.");

                return SpreadsheetDocument.Open(_docPackage);
            }
            catch (Exception e)
            {
                throw new PowerToolsDocumentException(e.Message);
            }
        }

        public PresentationDocument GetPresentationDocument()
        {
            try
            {
                if (GetDocumentType() != typeof(PresentationDocument))
                    throw new PowerToolsDocumentException("Not a Presentation document.");

                return PresentationDocument.Open(_docPackage);
            }
            catch (Exception e)
            {
                throw new PowerToolsDocumentException(e.Message);
            }
        }

        public Type GetDocumentType()
        {
            PackageRelationship relationship = _docPackage
                .GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
                .FirstOrDefault();

            if (relationship == null)
            {
                relationship = _docPackage
                    .GetRelationshipsByType("http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument")
                    .FirstOrDefault();
            }

            if (relationship == null)
            {
                throw new PowerToolsDocumentException("Not an Open XML Document.");
            }

            PackagePart part = _docPackage.GetPart(PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri));
            switch (part.ContentType)
            {
                case "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml":
                case "application/vnd.ms-word.document.macroEnabled.main+xml":
                case "application/vnd.ms-word.template.macroEnabledTemplate.main+xml":
                case "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml":
                    return typeof(WordprocessingDocument);

                case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml":
                case "application/vnd.ms-excel.sheet.macroEnabled.main+xml":
                case "application/vnd.ms-excel.template.macroEnabled.main+xml":
                case "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml":
                    return typeof(SpreadsheetDocument);

                case "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml":
                case "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml":
                case "application/vnd.ms-powerpoint.template.macroEnabled.main+xml":
                case "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml":
                case "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml":
                case "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml":
                    return typeof(PresentationDocument);
            }

            return null;
        }

        public OpenXmlPowerToolsDocument GetModifiedDocument()
        {
            _docPackage.Close();
            _docPackage = null;
            return new OpenXmlPowerToolsDocument(_document?.FileName, _docMemoryStream);
        }

        public WmlDocument GetModifiedWmlDocument()
        {
            _docPackage.Close();
            _docPackage = null;
            return new WmlDocument(_document?.FileName, _docMemoryStream);
        }

        public SmlDocument GetModifiedSmlDocument()
        {
            _docPackage.Close();
            _docPackage = null;
            return new SmlDocument(_document?.FileName, _docMemoryStream);
        }

        public PmlDocument GetModifiedPmlDocument()
        {
            _docPackage.Close();
            _docPackage = null;
            return new PmlDocument(_document?.FileName, _docMemoryStream);
        }

        public void Close()
        {
            Dispose(true);
        }

        ~OpenXmlMemoryStreamDocument()
        {
            Dispose(false);
        }

        private void Dispose(bool disposing)
        {
            if (disposing)
            {
                _docPackage?.Close();
                _docMemoryStream?.Dispose();
            }

            if (_docPackage == null && _docMemoryStream == null)
            {
                return;
            }

            _docPackage = null;
            _docMemoryStream = null;

            GC.SuppressFinalize(this);
        }
    }
}
