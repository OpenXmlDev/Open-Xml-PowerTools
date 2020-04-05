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
    public class StronglyTypedBlockTests : TestsBase
    {
        [Fact]
        public void CanUseStronglyTypedBlockToDemarcateApis()
        {
            using (var stream = new MemoryStream())
            {
                CreateEmptyWordprocessingDocument(stream);

                using (var wordDocument = WordprocessingDocument.Open(stream, true))
                {
                    var part = wordDocument.MainDocumentPart;

                    // Add a paragraph through the PowerTools.
                    var content = part.GetXDocument();
                    var bodyElement = content.Descendants(W.body).First();
                    bodyElement.Add(new XElement(W.p, new XElement(W.r, new XElement(W.t, "Added through PowerTools"))));
                    part.PutXDocument();

                    // This demonstrates the use of the StronglyTypedBlock in a using statement to
                    // demarcate the intermittent use of the strongly typed classes.
                    using (new StronglyTypedBlock(wordDocument))
                    {
                        // Assert that we can see the paragraph added through the PowerTools.
                        var body = part.Document.Body;
                        var paragraphs = body.Elements<Paragraph>().ToList();
                        Assert.Single(paragraphs);
                        Assert.Equal("Added through PowerTools", paragraphs[0].InnerText);

                        // Add a paragraph through the SDK.
                        body.AppendChild(new Paragraph(new Run(new Text("Added through SDK"))));
                    }

                    // Assert that we can see the paragraphs added through the PowerTools and the SDK.
                    content = part.GetXDocument();
                    var paragraphElements = content.Descendants(W.p).ToList();
                    Assert.Equal(2, paragraphElements.Count);
                    Assert.Equal("Added through PowerTools", paragraphElements[0].Value);
                    Assert.Equal("Added through SDK", paragraphElements[1].Value);
                }
            }
        }

        [Fact]
        public void ConstructorThrowsWhenPassingNull()
        {
            Assert.Throws<ArgumentNullException>(() => new StronglyTypedBlock(null));
        }
    }
}

#endif