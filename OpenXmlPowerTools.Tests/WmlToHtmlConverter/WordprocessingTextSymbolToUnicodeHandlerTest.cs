using System.Collections.Generic;
using Xunit;

namespace OpenXmlPowerTools.Tests.WmlToHtmlConverter
{
    public class WordprocessingTextSymbolToUnicodeHandlerTest
    {
        [Theory]
        [InlineData("1", "•1", "Symbol")]
        [InlineData("1", "1", "arial")]
        public void ShouldReplaceWithEquivalent(string original, string expectedEquivalent, string fontFamily)
        {
            var currentStyle = new Dictionary<string, string> { { "font-family", fontFamily } };

            var WordprocessingTextSymbolToUnicodeHandler = new WordprocessingTextSymbolToUnicodeHandler();

            var actual = WordprocessingTextSymbolToUnicodeHandler.TransformText(original, currentStyle);

            Assert.Equal(expectedEquivalent, actual);
        }
    }
}