using System;

namespace OpenXmlPowerTools
{
    public class InvalidOpenXmlDocumentException : Exception
    {
        public InvalidOpenXmlDocumentException(string message) : base(message)
        {
        }

        public InvalidOpenXmlDocumentException()
        {
        }

        public InvalidOpenXmlDocumentException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}