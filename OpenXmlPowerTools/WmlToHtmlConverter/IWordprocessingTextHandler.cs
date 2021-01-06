using System.Collections.Generic;

namespace OpenXmlPowerTools
{
    /// <summary>
    /// Handler that changes every text in w:t
    /// </summary>
    public interface IWordprocessingTextHandler
    {
        /// <summary>
        /// Returns some kind of changed text, that will be used instead of the original
        /// </summary>
        /// <param name="text"></param>
        /// <param name="fontFamily"></param>
        /// <returns></returns>
        public string TransformText(string text, Dictionary<string, string> fontFamily);
    }
}