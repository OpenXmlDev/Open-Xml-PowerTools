using Codeuctivity.Exceptions;
using DocumentFormat.OpenXml.Packaging;
using System.IO;

namespace Codeuctivity
{
    public partial class PmlDocument : OpenXmlPowerToolsDocument
    {
        public PmlDocument(OpenXmlPowerToolsDocument original)
            : base(original)
        {
            if (GetDocumentType() != typeof(PresentationDocument))
            {
                throw new PowerToolsDocumentException("Not a Presentation document.");
            }
        }

        public PmlDocument(OpenXmlPowerToolsDocument original, bool convertToTransitional)
            : base(original, convertToTransitional)
        {
            if (GetDocumentType() != typeof(PresentationDocument))
            {
                throw new PowerToolsDocumentException("Not a Presentation document.");
            }
        }

        public PmlDocument(string fileName)
            : base(fileName)
        {
            if (GetDocumentType() != typeof(PresentationDocument))
            {
                throw new PowerToolsDocumentException("Not a Presentation document.");
            }
        }

        public PmlDocument(string fileName, bool convertToTransitional)
            : base(fileName, convertToTransitional)
        {
            if (GetDocumentType() != typeof(PresentationDocument))
            {
                throw new PowerToolsDocumentException("Not a Presentation document.");
            }
        }

        public PmlDocument(string fileName, byte[] byteArray)
            : base(byteArray)
        {
            FileName = fileName;
            if (GetDocumentType() != typeof(PresentationDocument))
            {
                throw new PowerToolsDocumentException("Not a Presentation document.");
            }
        }

        public PmlDocument(string fileName, byte[] byteArray, bool convertToTransitional)
            : base(byteArray, convertToTransitional)
        {
            FileName = fileName;
            if (GetDocumentType() != typeof(PresentationDocument))
            {
                throw new PowerToolsDocumentException("Not a Presentation document.");
            }
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
}