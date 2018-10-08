using System;

namespace OpenXmlPowerTools
{
    public class InvalidOpenXmlDocumentException : Exception
    {
        public InvalidOpenXmlDocumentException(string message) : base(message)
        {
        }
    }
}
