// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#define COPY_FILES_FOR_DEBUGGING

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
using OpenXmlPowerTools;
using Xunit;

#if !ELIDE_XUNIT_TESTS

namespace OxPt
{
    public class CaTests
    {
        /*
         * This test was removed because it depends on the Coalesce method, which is only ever used
         * by this test.
         *
        [Theory]
        [InlineData("CA/CA001-Plain.docx", 60)]
        [InlineData("CA/CA002-Bookmark.docx", 7)]
        [InlineData("CA/CA003-Numbered-List.docx", 8)]
        [InlineData("CA/CA004-TwoParas.docx", 88)]
        [InlineData("CA/CA005-Table.docx", 27)]
        [InlineData("CA/CA006-ContentControl.docx", 60)]
        [InlineData("CA/CA007-DayLong.docx", 10)]
        [InlineData("CA/CA008-Footnote-Reference.docx", 23)]
        [InlineData("CA/CA010-Delete-Run.docx", 16)]
        [InlineData("CA/CA011-Insert-Run.docx", 16)]
        [InlineData("CA/CA012-fldSimple.docx", 10)]
        [InlineData("CA/CA013-Lots-of-Stuff.docx", 168)]
        [InlineData("CA/CA014-Complex-Table.docx", 193)]
        [InlineData("WC/WC024-Table-Before.docx", 24)]
        [InlineData("WC/WC024-Table-After2.docx", 18)]
        //[InlineData("", 0)]
        //[InlineData("", 0)]
        //[InlineData("", 0)]
        //[InlineData("", 0)]

        public void CA001_ContentAtoms(string name, int contentAtomCount)
        {
            FileInfo sourceDocx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name));

            var thisGuid = Guid.NewGuid().ToString().Replace("-", "");
            var sourceCopiedToDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", string.Format("-{0}-1-Source.docx", thisGuid))));
            if (!sourceCopiedToDestDocx.Exists)
                File.Copy(sourceDocx.FullName, sourceCopiedToDestDocx.FullName);

            var coalescedDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", string.Format("-{0}-2-Coalesced.docx", thisGuid))));
            if (!coalescedDocx.Exists)
                File.Copy(sourceDocx.FullName, coalescedDocx.FullName);

            var contentAtomDataFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", string.Format("-{0}-3-ContentAtomData.txt", thisGuid))));

            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(coalescedDocx.FullName, true))
            {
                var contentParent = wDoc.MainDocumentPart.GetXDocument().Root.Element(W.body);
                var settings = new WmlComparerSettings();
                ComparisonUnitAtom[] contentAtomList = WmlComparer.CreateComparisonUnitAtomList(wDoc.MainDocumentPart, contentParent, settings);
                StringBuilder sb = new StringBuilder();
                var part = wDoc.MainDocumentPart;

                sb.AppendFormat("Part: {0}", part.Uri.ToString());
                sb.Append(Environment.NewLine);
                sb.Append(ComparisonUnit.ComparisonUnitListToString(contentAtomList.ToArray()) + Environment.NewLine);
                sb.Append(Environment.NewLine);

                XDocument newMainXDoc = WmlComparer.Coalesce(contentAtomList);
                var partXDoc = wDoc.MainDocumentPart.GetXDocument();
                partXDoc.Root.ReplaceWith(newMainXDoc.Root);
                wDoc.MainDocumentPart.PutXDocument();

                File.WriteAllText(contentAtomDataFi.FullName, sb.ToString());

                Assert.Equal(contentAtomCount, contentAtomList.Count());
            }
        }
        */

        [Theory]
        [InlineData("HC009-Test-04.docx")]
        public void CA002_Annotations(string name)
        {
            FileInfo sourceDocx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name));

#if COPY_FILES_FOR_DEBUGGING
            var sourceCopiedToDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-1-Source.docx")));
            if (!sourceCopiedToDestDocx.Exists)
                File.Copy(sourceDocx.FullName, sourceCopiedToDestDocx.FullName);

            var annotatedDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-2-Annotated.docx")));
            if (!annotatedDocx.Exists)
                File.Copy(sourceDocx.FullName, annotatedDocx.FullName);

            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(annotatedDocx.FullName, true))
            {
                var contentParent = wDoc.MainDocumentPart.GetXDocument().Root.Element(W.body);
                var settings = new WmlComparerSettings();
                WmlComparer.CreateComparisonUnitAtomList(wDoc.MainDocumentPart, contentParent, settings);
            }
#endif
        }

        [Theory]
        [InlineData("CA/CA009-altChunk.docx")]
        //[InlineData("")]
        //[InlineData("")]
        //[InlineData("")]

        public void CA003_ContentAtoms_Throws(string name)
        {
            FileInfo sourceDocx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name));
            var thisGuid = Guid.NewGuid().ToString().Replace("-", "");
            var sourceCopiedToDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", string.Format("-{0}-1-Source.docx", thisGuid))));
            if (!sourceCopiedToDestDocx.Exists)
                File.Copy(sourceDocx.FullName, sourceCopiedToDestDocx.FullName);

            var coalescedDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", string.Format("-{0}-2-Coalesced.docx", thisGuid))));
            if (!coalescedDocx.Exists)
                File.Copy(sourceDocx.FullName, coalescedDocx.FullName);

            var contentAtomDataFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", string.Format("-{0}-3-ContentAtomData.txt", thisGuid))));

            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(coalescedDocx.FullName, true))
            {
                Assert.Throws<NotSupportedException>(() =>
                {
                    var contentParent = wDoc.MainDocumentPart.GetXDocument().Root.Element(W.body);
                    var settings = new WmlComparerSettings();
                    WmlComparer.CreateComparisonUnitAtomList(wDoc.MainDocumentPart, contentParent, settings);
                });
            }
        }
    }
}

#endif
