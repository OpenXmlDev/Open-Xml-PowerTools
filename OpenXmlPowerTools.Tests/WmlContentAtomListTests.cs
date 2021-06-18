#define COPY_FILES_FOR_DEBUGGING

using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using OpenXmlPowerTools.Tests;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using Xunit;

namespace OxPt
{
    public class CaTests
    {
        [Theory]
        [InlineData("HC009-Test-04.docx")]
        public void CA002_Annotations(string name)
        {
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));

#if COPY_FILES_FOR_DEBUGGING
            var sourceCopiedToDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-1-Source.docx")));
            if (!sourceCopiedToDestDocx.Exists)
            {
                File.Copy(sourceDocx.FullName, sourceCopiedToDestDocx.FullName);
            }

            var annotatedDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-2-Annotated.docx")));
            if (!annotatedDocx.Exists)
            {
                File.Copy(sourceDocx.FullName, annotatedDocx.FullName);
            }

            using var wDoc = WordprocessingDocument.Open(annotatedDocx.FullName, true);
            var contentParent = wDoc.MainDocumentPart.GetXDocument().Root.Element(W.body);
            var settings = new WmlComparerSettings();

            var type = typeof(WmlComparer);
            var method = type.GetMethods(BindingFlags.NonPublic | BindingFlags.Static).Where(x => x.Name == "CreateComparisonUnitAtomList" && x.IsStatic).Single();

            //Act
            method.Invoke(null, new object[] { wDoc.MainDocumentPart, contentParent, settings });
#endif
        }

        [Theory]
        [InlineData("CA/CA009-altChunk.docx")]
        public void CA003_ContentAtoms_Throws(string name)
        {
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));
            var thisGuid = Guid.NewGuid().ToString().Replace("-", "");
            var sourceCopiedToDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", string.Format(CultureInfo.InvariantCulture, "-{0}-1-Source.docx", thisGuid))));
            if (!sourceCopiedToDestDocx.Exists)
            {
                File.Copy(sourceDocx.FullName, sourceCopiedToDestDocx.FullName);
            }

            var coalescedDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", string.Format(CultureInfo.InvariantCulture, "-{0}-2-Coalesced.docx", thisGuid))));
            if (!coalescedDocx.Exists)
            {
                File.Copy(sourceDocx.FullName, coalescedDocx.FullName);
            }

            var contentAtomDataFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", string.Format(CultureInfo.InvariantCulture, "-{0}-3-ContentAtomData.txt", thisGuid))));

            using var wDoc = WordprocessingDocument.Open(coalescedDocx.FullName, true);
            var exception = Assert.Throws<TargetInvocationException>(() =>
              {
                  var contentParent = wDoc.MainDocumentPart.GetXDocument().Root.Element(W.body);
                  var settings = new WmlComparerSettings();

                  var type = typeof(WmlComparer);
                  var method = type.GetMethods(BindingFlags.NonPublic | BindingFlags.Static).Where(x => x.Name == "CreateComparisonUnitAtomList" && x.IsStatic).Single();

                  //Act
                  method.Invoke(null, new object[] { wDoc.MainDocumentPart, contentParent, settings });
              });

            Assert.IsType<NotSupportedException>(exception.InnerException);
        }
    }
}