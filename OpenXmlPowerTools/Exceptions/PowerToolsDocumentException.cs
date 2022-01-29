using System;

namespace Codeuctivity.OpenXmlPowerTools.Exceptions
{
    public class PowerToolsDocumentException : Exception
    {
        public PowerToolsDocumentException(string message) : base(message)
        {
        }

        public PowerToolsDocumentException()
        {
        }

        public PowerToolsDocumentException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}