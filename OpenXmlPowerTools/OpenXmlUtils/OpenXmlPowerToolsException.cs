using System;

namespace OpenXmlPowerTools
{
    public class OpenXmlPowerToolsException : Exception
    {
        public OpenXmlPowerToolsException(string message) : base(message)
        {
        }

        public OpenXmlPowerToolsException(string message, Exception innerException) : base(message, innerException)
        {
        }

        public OpenXmlPowerToolsException()
        {
        }
    }
}