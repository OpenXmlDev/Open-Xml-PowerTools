using Codeuctivity.BitmapCompare;
using OpenXmlPowerTools;
using OpenXmlPowerTools.OpenXMLWordprocessingMLToHtmlConverter;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace OxPt
{
    public class WmlToHtmlConverterHandlerTests
    {
        private const string minimalPng = "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAALiIAAC4iAari3ZIAAAAMSURBVBhXY0ACDAwAAA4AAXqxuTAAAAAASUVORK5CYII=";
        private const string minimalBmp = "Qk1CAAAAAAAAADoAAAAoAAAAAgAAAAIAAAABAAQAAAAAAAAAAAAiLgAAIi4AAAEAAAABAAAAAAAA/wAAAAAAAAAA";
        private const string minimalJpg = "/9j/4AAQSkZJRgABAQEBLAEsAAD/4QBoRXhpZgAATU0AKgAAAAgABAEaAAUAAAABAAAAPgEbAAUAAAABAAAARgEoAAMAAAABAAIAAAExAAIAAAARAAAATgAAAAAABJPfAAAD6AAEk98AAAPocGFpbnQubmV0IDQuMi4xNAAA/9sAQwACAQEBAQECAQEBAgICAgIEAwICAgIFBAQDBAYFBgYGBQYGBgcJCAYHCQcGBggLCAkKCgoKCgYICwwLCgwJCgoK/9sAQwECAgICAgIFAwMFCgcGBwoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoK/8AAEQgAAgACAwEhAAIRAQMRAf/EAB8AAAEFAQEBAQEBAAAAAAAAAAABAgMEBQYHCAkKC//EALUQAAIBAwMCBAMFBQQEAAABfQECAwAEEQUSITFBBhNRYQcicRQygZGhCCNCscEVUtHwJDNicoIJChYXGBkaJSYnKCkqNDU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6g4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfIycrS09TV1tfY2drh4uPk5ebn6Onq8fLz9PX29/j5+v/EAB8BAAMBAQEBAQEBAQEAAAAAAAABAgMEBQYHCAkKC//EALURAAIBAgQEAwQHBQQEAAECdwABAgMRBAUhMQYSQVEHYXETIjKBCBRCkaGxwQkjM1LwFWJy0QoWJDThJfEXGBkaJicoKSo1Njc4OTpDREVGR0hJSlNUVVZXWFlaY2RlZmdoaWpzdHV2d3h5eoKDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uLj5OXm5+jp6vLz9PX29/j5+v/aAAwDAQACEQMRAD8A/n/ooA//2Q==";

        [Fact]
        public void ShouldTranslateWithWordprocessingTextDummyHandler()
        {
            var expected = "someValue";
            Dictionary<string, string> fontFamily = default!;
            var wordprocessingTextDummyHandler = new TextDummyHandler();

            var actual = wordprocessingTextDummyHandler.TransformText(expected, fontFamily);

            Assert.Equal(expected, actual);
        }

        [Theory]
        [InlineData("png", minimalPng)]
        [InlineData("bmp", minimalBmp)]
        [InlineData("jpeg", minimalJpg)]
        public void ShouldTranslateWithDefaultImageHandler(string imageType, string minimalImage)
        {
            var expectedStart = $"<img src=\"data:image/{imageType};base64,";
            var expectedEnd = "\" xmlns=\"http://www.w3.org/1999/xhtml\" />";
            var binaryBitmap = Convert.FromBase64String(minimalImage);

            using var memeoryStream = new MemoryStream(binaryBitmap);
            var input = new Bitmap(memeoryStream);
            var imageInfo = new ImageInfo
            {
                Bitmap = input
            };

            var defaultImageHandler = new ImageHandler();

            var actual = defaultImageHandler.TransformImage(imageInfo).ToString();

            Assert.StartsWith(expectedStart, actual);
            Assert.EndsWith(expectedEnd, actual);

            var actualBase64Part = actual.Substring(expectedStart.Length, actual.Length - expectedEnd.Length - expectedStart.Length);
            var binaryActualBitmap = Convert.FromBase64String(actualBase64Part);
            using var memeoryStreamBinaryBitmap = new MemoryStream(binaryActualBitmap);
            var actualBitmap = new Bitmap(memeoryStreamBinaryBitmap);

            Assert.True(Compare.ImageAreEqual(input, input));
        }

        [Fact]
        public void ShouldTranslateSymbolsToUnicodeWithDefaultSymbolHandler()
        {
            Dictionary<string, string> fontFamily = default!;
            var defaultSymbolHandler = new SymbolHandler();

            var element = new XElement("symbol", new XAttribute(W._char, "A"));

            var actual = defaultSymbolHandler.TransformSymbol(element, fontFamily);

            Assert.Equal("<span xmlns=\"http://www.w3.org/1999/xhtml\">&#10;</span>", actual.ToString());
        }

        [Fact]
        public void ShouldTranslatePageBreaksWithBreakHandler()
        {
            var breakHandler = new BreakHandler();

            var element = new XElement("br");

            var actual = breakHandler.TransformBreak(element);

            Assert.Equal(3, actual.Count());
            Assert.Equal("<br xmlns=\"http://www.w3.org/1999/xhtml\" />", actual.ElementAt(0).ToString());
            Assert.Equal("&#x200e;", actual.ElementAt(1).ToString());
            Assert.Null(actual.ElementAt(2));
        }

        [Fact]
        public void ShouldTranslateFontInRunSymbolWithFontHandler()
        {
            var fontHandler = new FontHandler();

            var element = new XElement("run", new XElement(W.sym, new XAttribute(W.font, "SomeSymbolFont")), new XAttribute(PtOpenXml.FontName, "SomeRunFont"));

            var actual = fontHandler.TranslateRunStyleFont(element);

            Assert.Equal("SomeSymbolFont", actual);
        }

        [Fact]
        public void ShouldTranslateFontInRunWithFontHandler()
        {
            var fontHandler = new FontHandler();

            var element = new XElement("run", new XAttribute(PtOpenXml.FontName, "SomeRunFont"));

            var actual = fontHandler.TranslateRunStyleFont(element);

            Assert.Equal("SomeRunFont", actual);
        }

        [Fact]
        public void ShouldTranslateFontInParagraphWithFontHandler()
        {
            var fontHandler = new FontHandler();

            var element = new XElement("run", new XAttribute(PtOpenXml.FontName, "SomeRunFont"));

            var actual = fontHandler.TranslateParagraphStyleFont(element);

            Assert.Equal("SomeRunFont", actual);
        }
    }
}