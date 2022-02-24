// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;
using Xunit;

#if !ELIDE_XUNIT_TESTS

namespace OxPt
{
    public class FaTests
    {
        [Theory]
        [InlineData("FA001-00010", "FA/RevTracking/001-DeletedRun.docx")]
        [InlineData("FA001-00020", "FA/RevTracking/002-DeletedNumberedParagraphs.docx")]
        [InlineData("FA001-00030", "FA/RevTracking/003-DeletedFieldCode.docx")]
        [InlineData("FA001-00040", "FA/RevTracking/004-InsertedNumberingProperties.docx")]
        [InlineData("FA001-00050", "FA/RevTracking/005-InsertedNumberedParagraph.docx")]
        [InlineData("FA001-00060", "FA/RevTracking/006-DeletedTableRow.docx")]
        [InlineData("FA001-00070", "FA/RevTracking/007-InsertedTableRow.docx")]
        [InlineData("FA001-00080", "FA/RevTracking/008-InsertedFieldCode.docx")]
        [InlineData("FA001-00090", "FA/RevTracking/009-InsertedParagraph.docx")]
        [InlineData("FA001-00100", "FA/RevTracking/010-InsertedRun.docx")]
        [InlineData("FA001-00110", "FA/RevTracking/011-InsertedMathChar.docx")]
        [InlineData("FA001-00120", "FA/RevTracking/012-DeletedMathChar.docx")]
        [InlineData("FA001-00130", "FA/RevTracking/013-DeletedParagraph.docx")]
        [InlineData("FA001-00140", "FA/RevTracking/014-MovedParagraph.docx")]
        [InlineData("FA001-00150", "FA/RevTracking/015-InsertedContentControl.docx")]
        [InlineData("FA001-00160", "FA/RevTracking/016-DeletedContentControl.docx")]
        [InlineData("FA001-00170", "FA/RevTracking/017-NumberingChange.docx")]
        [InlineData("FA001-00180", "FA/RevTracking/018-ParagraphPropertiesChange.docx")]
        [InlineData("FA001-00190", "FA/RevTracking/019-RunPropertiesChange.docx")]
        [InlineData("FA001-00200", "FA/RevTracking/020-SectionPropertiesChange.docx")]
        [InlineData("FA001-00210", "FA/RevTracking/021-TableGridChange.docx")]
        [InlineData("FA001-00220", "FA/RevTracking/022-TablePropertiesChange.docx")]
        [InlineData("FA001-00230", "FA/RevTracking/023-CellPropertiesChange.docx")]
        [InlineData("FA001-00240", "FA/RevTracking/024-RowPropertiesChange.docx")]

        public void FA001_DocumentsWithRevTracking(string testId, string src)
        {
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Load the source document
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/");
            FileInfo sourceDocxFi = new FileInfo(Path.Combine(sourceDir.FullName, src));
            WmlDocument wmlSourceDocument = new WmlDocument(sourceDocxFi.FullName);

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Create the dir for the test
            var rootTempDir = TestUtil.TempDir;
            var thisTestTempDir = new DirectoryInfo(Path.Combine(rootTempDir.FullName, testId));
            if (thisTestTempDir.Exists)
                Assert.True(false, "Duplicate test id: " + testId);
            else
                thisTestTempDir.Create();
            var tempDirFullName = thisTestTempDir.FullName;

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Copy src DOCX to temp directory, for ease of review

            var sourceDocxCopiedToDestFileName = new FileInfo(Path.Combine(tempDirFullName, sourceDocxFi.Name));
            if (!sourceDocxCopiedToDestFileName.Exists)
                wmlSourceDocument.SaveAs(sourceDocxCopiedToDestFileName.FullName);

            var sourceDocxAcceptedCopiedToDestFileName = new FileInfo(Path.Combine(tempDirFullName, sourceDocxFi.Name.ToLower().Replace(".docx", "-accepted.docx")));
            var wmlSourceAccepted = RevisionProcessor.AcceptRevisions(wmlSourceDocument);
            wmlSourceAccepted.SaveAs(sourceDocxAcceptedCopiedToDestFileName.FullName);

            var outFi = new FileInfo(Path.Combine(tempDirFullName, "Output.docx"));
            FormattingAssemblerSettings settings = new FormattingAssemblerSettings();
            var assembledWml = FormattingAssembler.AssembleFormatting(wmlSourceDocument, settings);
            assembledWml.SaveAs(outFi.FullName);

            var outAcceptedFi = new FileInfo(Path.Combine(tempDirFullName, "Output-accepted.docx"));
            var assembledAcceptedWml = RevisionProcessor.AcceptRevisions(assembledWml);
            assembledAcceptedWml.SaveAs(outAcceptedFi.FullName);

            Validate(outFi);
        }

        private void Validate(FileInfo fi)
        {
            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(fi.FullName, true))
            {
                OpenXmlValidator v = new OpenXmlValidator();
                var errors = v.Validate(wDoc).Where(ve =>
                {
                    var found = s_ExpectedErrors.Any(xe => ve.Description.Contains(xe));
                    return !found;
                });

                if (errors.Count() != 0)
                {
                    StringBuilder sb = new StringBuilder();
                    foreach (var item in errors)
                    {
                        sb.Append(item.Description).Append(Environment.NewLine);
                    }
                    var s = sb.ToString();
                    Assert.True(false, s);
                }
            }
        }

        private static List<string> s_ExpectedErrors = new List<string>()
        {
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:evenHBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:evenVBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRow' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRowFirstColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRowLastColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRow' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRowFirstColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRowLastColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:noHBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:noVBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:oddHBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:oddVBand' attribute is not declared.",
            "The element has unexpected child element 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:updateFields'.",
            "The attribute 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:name' has invalid value 'useWord2013TrackBottomHyphenation'. The Enumeration constraint failed.",
            "The 'http://schemas.microsoft.com/office/word/2012/wordml:restartNumberingAfterBreak' attribute is not declared.",
            "Attribute 'id' should have unique value. Its current value '",
            "The 'urn:schemas-microsoft-com:mac:vml:blur' attribute is not declared.",
            "Attribute 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:id' should have unique value. Its current value '",
            "The element has unexpected child element 'http://schemas.microsoft.com/office/word/2012/wordml:",
            "The element has invalid child element 'http://schemas.microsoft.com/office/word/2012/wordml:",
            "The 'urn:schemas-microsoft-com:mac:vml:complextextbox' attribute is not declared.",
            "http://schemas.microsoft.com/office/word/2010/wordml:",
            "http://schemas.microsoft.com/office/word/2008/9/12/wordml:",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:allStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:customStyles' attribute is not declared.",
            "The element has invalid child element 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:ins'.",
            "The element has invalid child element 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:del'.",
        };
    }
}
#endif
