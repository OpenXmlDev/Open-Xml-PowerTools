using System;

namespace OpenXmlPowerTools.HtmlToWml.CSS
{
    public class FatalError : Exception
    {
        public FatalError(string m) : base(m)
        {
        }

        public FatalError(string message, Exception innerException) : base(message, innerException)
        {
        }

        public FatalError()
        {
        }
    }
}