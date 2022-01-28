using Codeuctivity.Exceptions;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace Codeuctivity
{
    public class OpenXmlMemoryStreamDocument : IDisposable
    {
        private readonly OpenXmlPowerToolsDocument Document;
        private MemoryStream DocMemoryStream;
        private Package DocPackage;

        public OpenXmlMemoryStreamDocument(OpenXmlPowerToolsDocument doc)
        {
            Document = doc;
            DocMemoryStream = new MemoryStream();
            DocMemoryStream.Write(doc.DocumentByteArray, 0, doc.DocumentByteArray.Length);
            try
            {
                DocPackage = Package.Open(DocMemoryStream, FileMode.Open);
            }
            catch (Exception e)
            {
                throw new PowerToolsDocumentException(e.Message);
            }
        }

        internal OpenXmlMemoryStreamDocument(MemoryStream stream)
        {
            DocMemoryStream = stream;
            try
            {
                DocPackage = Package.Open(DocMemoryStream, FileMode.Open);
            }
            catch (Exception e)
            {
                throw new PowerToolsDocumentException(e.Message);
            }
        }

        public static OpenXmlMemoryStreamDocument CreateWordprocessingDocument()
        {
            var stream = new MemoryStream();
            using var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
            doc.AddMainDocumentPart();
            doc.MainDocumentPart.PutXDocument(new XDocument(
                new XElement(W.document,
                    new XAttribute(XNamespace.Xmlns + "w", W.w),
                    new XAttribute(XNamespace.Xmlns + "r", R.r),
                    new XElement(W.body))));
            doc.Close();
            return new OpenXmlMemoryStreamDocument(stream);
        }

        public static OpenXmlMemoryStreamDocument CreateSpreadsheetDocument()
        {
            var stream = new MemoryStream();
            using var doc = SpreadsheetDocument.Create(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
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

        public static OpenXmlMemoryStreamDocument CreatePresentationDocument()
        {
            var stream = new MemoryStream();
            using var doc = PresentationDocument.Create(stream, DocumentFormat.OpenXml.PresentationDocumentType.Presentation);
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

        public static OpenXmlMemoryStreamDocument CreatePackage()
        {
            var stream = new MemoryStream();
            var package = Package.Open(stream, FileMode.Create);
            package.Close();
            return new OpenXmlMemoryStreamDocument(stream);
        }

        public Package GetPackage()
        {
            return DocPackage;
        }

        public WordprocessingDocument GetWordprocessingDocument()
        {
            try
            {
                if (GetDocumentType() != typeof(WordprocessingDocument))
                {
                    throw new PowerToolsDocumentException("Not a Wordprocessing document.");
                }

                return WordprocessingDocument.Open(DocPackage);
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
                {
                    throw new PowerToolsDocumentException("Not a Spreadsheet document.");
                }

                return SpreadsheetDocument.Open(DocPackage);
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
                {
                    throw new PowerToolsDocumentException("Not a Presentation document.");
                }

                return PresentationDocument.Open(DocPackage);
            }
            catch (Exception e)
            {
                throw new PowerToolsDocumentException(e.Message);
            }
        }

        public Type GetDocumentType()
        {
            var relationship = DocPackage.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument").FirstOrDefault();
            if (relationship == null)
            {
                relationship = DocPackage.GetRelationshipsByType("http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument").FirstOrDefault();
            }

            if (relationship == null)
            {
                throw new PowerToolsDocumentException("Not an Open XML Document.");
            }

            var part = DocPackage.GetPart(PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri));
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
            DocPackage.Close();
            DocPackage = null;
            return new OpenXmlPowerToolsDocument(Document?.FileName, DocMemoryStream);
        }

        public WmlDocument GetModifiedWmlDocument()
        {
            DocPackage.Close();
            DocPackage = null;
            return new WmlDocument(Document?.FileName, DocMemoryStream);
        }

        public SmlDocument GetModifiedSmlDocument()
        {
            DocPackage.Close();
            DocPackage = null;
            return new SmlDocument(Document?.FileName, DocMemoryStream);
        }

        public PmlDocument GetModifiedPmlDocument()
        {
            DocPackage.Close();
            DocPackage = null;
            return new PmlDocument(Document?.FileName, DocMemoryStream);
        }

        public void Close()
        {
            Dispose(true);
        }

        public void Dispose()
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
                if (DocPackage != null)
                {
                    DocPackage.Close();
                }
                if (DocMemoryStream != null)
                {
                    DocMemoryStream.Dispose();
                }
            }
            if (DocPackage == null && DocMemoryStream == null)
            {
                return;
            }

            DocPackage = null;
            DocMemoryStream = null;
            GC.SuppressFinalize(this);
        }
    }
}