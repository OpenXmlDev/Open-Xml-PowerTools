// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

#if !ELIDE_XUNIT_TESTS

namespace OpenXmlPowerTools.Tests
{
    public class PowerToolsBlockTests : TestsBase
    {
        [Fact]
        public void CanUsePowerToolsBlockToDemarcateApis()
        {
            using (var stream = new MemoryStream())
            {
                CreateEmptyWordprocessingDocument(stream);

                using (var wordDocument = WordprocessingDocument.Open(stream, true))
                {
                    var part = wordDocument.MainDocumentPart;

                    // Add a paragraph through the SDK.
                    var body = part.Document.Body;
                    body.AppendChild(new Paragraph(new Run(new Text("Added through SDK"))));

                    // This demonstrates the use of the PowerToolsBlock in a using statement to
                    // demarcate the intermittent use of the PowerTools.
                    using (new PowerToolsBlock(wordDocument))
                    {
                        // Assert that we can see the paragraph added through the strongly typed classes.
                        var content = part.GetXDocument();
                        var paragraphElements = content.Descendants(W.p).ToList();
                        Assert.Single(paragraphElements);
                        Assert.Equal("Added through SDK", paragraphElements[0].Value);

                        // Add a paragraph through the PowerTools.
                        var bodyElement = content.Descendants(W.body).First();
                        bodyElement.Add(new XElement(W.p, new XElement(W.r, new XElement(W.t, "Added through PowerTools"))));
                        part.PutXDocument();
                    }

                    // Get the part's content through the SDK. Having used the PowerToolsBlock,
                    // we should see both paragraphs.
                    body = part.Document.Body;
                    var paragraphs = body.Elements<Paragraph>().ToList();
                    Assert.Equal(2, paragraphs.Count);
                    Assert.Equal("Added through SDK", paragraphs[0].InnerText);
                    Assert.Equal("Added through PowerTools", paragraphs[1].InnerText);
                }
            }
        }

        [Fact]
        public void ConstructorThrowsWhenPassingNull()
        {
            Assert.Throws<ArgumentNullException>(() => new PowerToolsBlock(null));
        }
    }
}

#endif