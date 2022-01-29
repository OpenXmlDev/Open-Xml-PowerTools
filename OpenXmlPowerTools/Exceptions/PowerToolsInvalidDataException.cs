using System;

namespace Codeuctivity.OpenXmlPowerTools.Exceptions
{
    public class PowerToolsInvalidDataException : Exception
    {
        public PowerToolsInvalidDataException(string message) : base(message)
        {
        }

        public PowerToolsInvalidDataException()
        {
        }

        public PowerToolsInvalidDataException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}