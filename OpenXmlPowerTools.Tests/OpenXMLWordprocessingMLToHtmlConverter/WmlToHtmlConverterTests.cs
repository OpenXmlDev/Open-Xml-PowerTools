using Codeuctivity.HtmlRenderer;
using Codeuctivity.OpenXmlPowerTools.OpenXMLWordprocessingMLToHtmlConverter;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Xunit;

namespace Codeuctivity.Tests.OpenXMLWordprocessingMLToHtmlConverter
{
    public class WmlToHtmlConverterTests
    {
        // PowerShell one liner that generates InlineData for all files in a directory
        // dir | % { '[InlineData("' + $_.Name + '")]' } | clip

        [Theory]
        [InlineData("HC001-5DayTourPlanTemplate.docx", 10000, false)]
        [InlineData("HC002-Hebrew-01.docx", 1593, false)]
        [InlineData("HC003-Hebrew-02.docx", 0, false)]
        [InlineData("HC004-ResumeTemplate.docx", 0, false)]
        [InlineData("HC005-TaskPlanTemplate.docx", 0, false)]
        [InlineData("HC006-Test-01.docx", 80000, false)]
        [InlineData("HC007-Test-02.docx", 0, true)]
        [InlineData("HC008-Test-03.docx", 0, false)]
        [InlineData("HC009-Test-04.docx", 0, false)]
        [InlineData("HC010-Test-05.docx", 0, true)]
        [InlineData("HC011-Test-06.docx", 80000, false)]
        [InlineData("HC012-Test-07.docx", 10, false)]
        [InlineData("HC013-Test-08.docx", 0, false)]
        [InlineData("HC014-RTL-Table-01.docx", 20000, false)]
        [InlineData("HC015-Vertical-Spacing-atLeast.docx", 0, false)]
        [InlineData("HC016-Horizontal-Spacing-firstLine.docx", 0, false)]
        [InlineData("HC017-Vertical-Alignment-Cell-01.docx", 25000, false)]
        [InlineData("HC018-Vertical-Alignment-Para-01.docx", 0, false)]
        [InlineData("HC019-Hidden-Run.docx", 0, false)]
        [InlineData("HC020-Small-Caps.docx", 0, false)]
        [InlineData("HC021-Symbols.docx", 0, false)]
        [InlineData("HC022-Table-Of-Contents.docx", 10, false)]
        [InlineData("HC023-Hyperlink.docx", 0, false)]
        [InlineData("HC024-Tabs-01.docx", 0, false)]
        [InlineData("HC025-Tabs-02.docx", 0, false)]
        [InlineData("HC026-Tabs-03.docx", 0, false)]
        [InlineData("HC027-Tabs-04.docx", 0, false)]
        [InlineData("HC028-No-Break-Hyphen.docx", 0, false)]
        [InlineData("HC029-Table-Merged-Cells.docx", 30000, false)]
        [InlineData("HC030-Content-Controls.docx", 0, false)]
        [InlineData("HC031-Complicated-Document.docx", 5000, false)]
        [InlineData("HC032-Named-Color.docx", 0, false)]
        [InlineData("HC033-Run-With-Border.docx", 5000, false)]
        [InlineData("HC034-Run-With-Position.docx", 0, false)]
        [InlineData("HC035-Strike-Through.docx", 0, false)]
        [InlineData("HC036-Super-Script.docx", 0, false)]
        [InlineData("HC037-Sub-Script.docx", 0, false)]
        [InlineData("HC038-Conflicting-Border-Weight.docx", 10000, false)]
        [InlineData("HC039-Bold.docx", 0, false)]
        [InlineData("HC040-Hyperlink-Fieldcode-01.docx", 0, false)]
        [InlineData("HC041-Hyperlink-Fieldcode-02.docx", 0, false)]
        [InlineData("HC042-Image-Png.docx", 0, false)]
        [InlineData("HC043-Chart.docx", 0, false)]
        [InlineData("HC044-Embedded-Workbook.docx", 0, false)]
        [InlineData("HC045-Italic.docx", 0, false)]
        [InlineData("HC046-BoldAndItalic.docx", 0, false)]
        [InlineData("HC047-No-Section.docx", 0, false)]
        [InlineData("HC048-Excerpt.docx", 0, false)]
        [InlineData("HC049-Borders.docx", 0, false)]
        [InlineData("HC050-Shaded-Text-01.docx", 15000, false)]
        [InlineData("HC051-Shaded-Text-02.docx", 15000, false)]
        [InlineData("HC060-Image-with-Hyperlink.docx", 0, false)]
        [InlineData("HC061-Hyperlink-in-Field.docx", 0, false)]
        [InlineData("Tabs.docx", 0, false)]
        public async Task HC001(string name, int expectedPixeNoise, bool imageSizeMayDiffer)
        {
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));
            var settings = new WmlToHtmlConverterSettings(sourceDocx.FullName);

            var oxPtConvertedDestHtml = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-3-OxPt.html")));
            await ConvertToHtml(sourceDocx, oxPtConvertedDestHtml, settings, expectedPixeNoise, imageSizeMayDiffer);
        }

        [Theory]
        [InlineData("HC006-Test-01.docx", 80000)]
        public async Task HC002_NoCssClasses(string name, int expectedPixeNoise)
        {
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));
            var settings = new WmlToHtmlConverterSettings(sourceDocx.FullName, new ImageHandler(), new TextDummyHandler(), new SymbolHandler(), new BreakHandler(), new FontHandler(), false);

            var oxPtConvertedDestHtml = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-5-OxPt-No-CSS-Classes.html")));
            await ConvertToHtml(sourceDocx, oxPtConvertedDestHtml, settings, expectedPixeNoise, false);
        }

        private static async Task ConvertToHtml(FileInfo sourceDocx, FileInfo destFileName, WmlToHtmlConverterSettings settings, int expectedPixeNoise, bool imageSizeMayDiffer)
        {
            var byteArray = File.ReadAllBytes(sourceDocx.FullName);
            var expectedRenderdResult = Path.Combine(sourceDocx.Directory.FullName, sourceDocx.Name + "Expectation.png");
            using var memoryStream = new MemoryStream();
            memoryStream.Write(byteArray, 0, byteArray.Length);
            using var wDoc = WordprocessingDocument.Open(memoryStream, true);
            var outputDirectory = destFileName.Directory;
            destFileName = new FileInfo(Path.Combine(outputDirectory.FullName, destFileName.Name));

            var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);

            // Note: the XHTML returned by ConvertToHtmlTransform contains objects of type XEntity. PtOpenXmlUtil.cs define the XEntity class. See http://blogs.msdn.com/ericwhite/archive/2010/01/21/writing-entity-references-using-linq-to-xml.aspx for detailed explanation. If you further transform the XML tree returned by ConvertToHtmlTransform, you must do it correctly, or entities will not be serialized properly.

            var htmlString = html.ToString(SaveOptions.DisableFormatting);
            File.WriteAllText(destFileName.FullName, htmlString, Encoding.UTF8);

            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                return;
            }

            await AssertRenderedHtmlIsEqual(destFileName.FullName, expectedRenderdResult, expectedPixeNoise, imageSizeMayDiffer);
        }

        internal static async Task AssertRenderedHtmlIsEqual(string actualFilePath, string expectReferenceFilePath, int allowedPixelErrorCount, bool imageSizeMayDiffer)
        {
            var actualFullPath = Path.GetFullPath(actualFilePath);

            Assert.True(File.Exists(actualFullPath), $"actualFilePath not found {actualFullPath}");
            await using var chromiumRenderer = await Renderer.CreateAsync();
            var pathRasterizedHtml = actualFilePath + ".png";
            await chromiumRenderer.ConvertHtmlToPng(actualFilePath, pathRasterizedHtml);

            await AssertImageIsEqual(pathRasterizedHtml, expectReferenceFilePath, allowedPixelErrorCount, imageSizeMayDiffer);
        }

        internal static async Task AssertImageIsEqual(string actualImagePath, string expectImageFilePath, int allowedPixelErrorCount, bool imageSizeMayDiffer)
        {
            var actualFullPath = Path.GetFullPath(actualImagePath);
            var expectFullPath = Path.GetFullPath(expectImageFilePath);

            Assert.True(File.Exists(actualFullPath), $"actualImagePath not found {actualFullPath}");

            //Uncomment following line to create or update expectation for new test cases
            //File.Copy(actualFullPath, expectFullPath, true);

            Assert.True(File.Exists(expectFullPath), $"ExpectReferenceImagePath not found \n{expectFullPath}\n copy over \n{actualFullPath}\n if this is a new test case.");

            if (ImageSharpCompare.ImageSharpCompare.ImagesAreEqual(actualFullPath, expectFullPath))
            {
                return;
            }

            var osSpecificDiffFileSuffix = RuntimeInformation.IsOSPlatform(OSPlatform.Linux) ? "linux" : "win";

            var allowedDiffImage = $"{expectFullPath}.diff.{osSpecificDiffFileSuffix}.png";
            var newDiffImage = $"{actualFullPath}.diff.png";

            if (!ImageSharpCompare.ImageSharpCompare.ImagesHaveEqualSize(actualFullPath, expectFullPath))
            {
                if (imageSizeMayDiffer)
                {
                    return;
                }

                // Uncomment following line to create or update a allowed diff file
                // File.Copy(actualFullPath, expectFullPath, true);

                SaveToGithubActionsPickupTestresultsDirectory(actualFullPath, expectFullPath);
                Assert.Fail($"Actual Dimension differs from expected \nExpected {expectFullPath}\ndiffers to actual {actualFullPath} \nReplace {expectFullPath} with the new value.");
            }

            using (var maskImage = ImageSharpCompare.ImageSharpCompare.CalcDiffMaskImage(actualFullPath, expectFullPath))
            {
                using var fs = new FileStream(newDiffImage, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                await maskImage.SaveAsync(fs, new SixLabors.ImageSharp.Formats.Png.PngEncoder());
            }

            // Uncomment following line to create or update a allowed diff file
            //File.Copy(actualFullPath, allowedDiffImage, true);

            if (File.Exists(allowedDiffImage))
            {
                if (!ImageSharpCompare.ImageSharpCompare.ImagesHaveEqualSize(actualFullPath, allowedDiffImage))
                {
                    Assert.True(false, $"AllowedDiffImage Dimension differs from allowed \nReplace {allowedDiffImage} with {actualFullPath}.");
                }

                var resultWithAllowedDiff = ImageSharpCompare.ImageSharpCompare.CalcDiff(actualFullPath, expectFullPath, allowedDiffImage);

                var pixelErrorCountAboveExpectedWithDiff = resultWithAllowedDiff.PixelErrorCount > allowedPixelErrorCount;
                if (pixelErrorCountAboveExpectedWithDiff)
                {
                    SaveToGithubActionsPickupTestresultsDirectory(actualFullPath, expectFullPath);
                    Assert.Fail($"Expected PixelErrorCount beyond {allowedPixelErrorCount} but was {resultWithAllowedDiff.PixelErrorCount}\nExpected {expectFullPath}\ndiffers to actual {actualFullPath}\n diff is {newDiffImage}\n");
                }
                return;
            }

            var result = ImageSharpCompare.ImageSharpCompare.CalcDiff(actualFullPath, expectFullPath);

            var pixelErrorCountAboveExpected = result.PixelErrorCount > allowedPixelErrorCount;
            if (pixelErrorCountAboveExpected)
            {
                SaveToGithubActionsPickupTestresultsDirectory(actualFullPath, expectFullPath);

                Assert.Fail($"Expected PixelErrorCount beyond {allowedPixelErrorCount} but was {result.PixelErrorCount}\nExpected {expectFullPath}\ndiffers to actual {actualFullPath}\n Diff is {newDiffImage} \nReplace {actualFullPath} with the new value or store the diff as {allowedDiffImage}.");
            }
        }

        private static void SaveToGithubActionsPickupTestresultsDirectory(string actualFullPath, string expectFullPath)
        {
            var fileName = Path.GetFileName(actualFullPath);
            var expectFullDirectory = Path.GetDirectoryName(expectFullPath);
            var expectFullDirectoryFullPath = Path.GetFullPath(expectFullDirectory);

            var testResultDirectoryActual = Path.Combine(expectFullDirectoryFullPath, "../TestResult/Actual");
            var testResultDirectoryExpected = Path.Combine(expectFullDirectoryFullPath, "../TestResult/Expected");
            CreateDirectory(Path.Combine(expectFullDirectoryFullPath, "../TestResult"));
            CreateDirectory(testResultDirectoryActual);
            CreateDirectory(testResultDirectoryExpected);

            File.Copy(actualFullPath, Path.Combine(testResultDirectoryActual, fileName), true);
            File.Copy(expectFullPath, Path.Combine(testResultDirectoryExpected, fileName), true);

            static void CreateDirectory(string testResultDirectory)
            {
                if (!Directory.Exists(testResultDirectory))
                {
                    Directory.CreateDirectory(testResultDirectory);
                }
            }
        }
    }
}
