using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OpenXmlPowerTools;
using OpenXmlPowerTools.Tests;
using System;
using System.IO;
using System.Linq;
using System.Text;
using Xunit;

namespace OxPt
{
    public class WcTests2
    {
        [Theory]
        [InlineData("CZ-1000", "CZ/CZ001-Plain.docx", "CZ/CZ001-Plain-Mod.docx", 1)]
        [InlineData("CZ-1010", "CZ/CZ002-Multi-Paragraphs.docx", "CZ/CZ002-Multi-Paragraphs-Mod.docx", 1)]
        [InlineData("CZ-1020", "CZ/CZ003-Multi-Paragraphs.docx", "CZ/CZ003-Multi-Paragraphs-Mod.docx", 1)]
        [InlineData("CZ-1030", "CZ/CZ004-Multi-Paragraphs-in-Cell.docx", "CZ/CZ004-Multi-Paragraphs-in-Cell-Mod.docx", 1)]
        public void CZ001_CompareTrackedInPrev(string testId, string name1, string name2, int revisionCount)
        {
            // TODO: Do we need to keep the revision count parameter?
            Assert.Equal(1, revisionCount);

            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var source1Docx = new FileInfo(Path.Combine(sourceDir.FullName, name1));
            var source2Docx = new FileInfo(Path.Combine(sourceDir.FullName, name2));

            var rootTempDir = TestUtil.TempDir;
            var thisTestTempDir = new DirectoryInfo(Path.Combine(rootTempDir.FullName, testId));
            Assert.False(thisTestTempDir.Exists, "Duplicate test id???");
            thisTestTempDir.Create();

            var source1CopiedToDestDocx = new FileInfo(Path.Combine(thisTestTempDir.FullName, source1Docx.Name));
            var source2CopiedToDestDocx = new FileInfo(Path.Combine(thisTestTempDir.FullName, source2Docx.Name));
            File.Copy(source1Docx.FullName, source1CopiedToDestDocx.FullName);
            File.Copy(source2Docx.FullName, source2CopiedToDestDocx.FullName);

            var before = source1CopiedToDestDocx.Name.Replace(".docx", "");
            var after = source2CopiedToDestDocx.Name.Replace(".docx", "");
            var docxWithRevisionsFi = new FileInfo(Path.Combine(thisTestTempDir.FullName, before + "-COMPARE-" + after + ".docx"));

            var source1Wml = new WmlDocument(source1CopiedToDestDocx.FullName);
            var source2Wml = new WmlDocument(source2CopiedToDestDocx.FullName);
            var settings = new WmlComparerSettings
            {
                DebugTempFileDi = thisTestTempDir
            };
            var comparedWml = WmlComparer.Compare(source1Wml, source2Wml, settings);

            comparedWml.SaveAs(docxWithRevisionsFi.FullName);
            using var ms = new MemoryStream();
            ms.Write(comparedWml.DocumentByteArray, 0, comparedWml.DocumentByteArray.Length);
            using var wDoc = WordprocessingDocument.Open(ms, true);
            var validator = new OpenXmlValidator();
            var errors = validator.Validate(wDoc).Where(e => !ExpectedErrors.Contains(e.Description));
            if (errors.Any())
            {
                var ind = "  ";
                var sb = new StringBuilder();
                foreach (var err in errors)
                {
                    sb.Append("Error" + Environment.NewLine);
                    sb.Append(ind + "ErrorType: " + err.ErrorType.ToString() + Environment.NewLine);
                    sb.Append(ind + "Description: " + err.Description + Environment.NewLine);
                    sb.Append(ind + "Part: " + err.Part.Uri.ToString() + Environment.NewLine);
                    sb.Append(ind + "XPath: " + err.Path.XPath + Environment.NewLine);
                }
                var sbs = sb.ToString();

                Assert.True(sbs.Length == 0, sbs);
            }
        }

        public static readonly string[] ExpectedErrors = new string[] {
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRow' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRow' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:noHBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:noVBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:allStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:customStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:latentStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:stylesInUse' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:headingStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:numberingStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:tableStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:directFormattingOnRuns' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:directFormattingOnParagraphs' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:directFormattingOnNumbering' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:directFormattingOnTables' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:clearFormatting' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:top3HeadingStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:visibleStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:alternateStyleNames' attribute is not declared.",
            "The attribute 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:val' has invalid value '0'. The MinInclusive constraint failed. The value must be greater than or equal to 1.",
            "The attribute 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:val' has invalid value '0'. The MinInclusive constraint failed. The value must be greater than or equal to 2.",
        };
    }
}