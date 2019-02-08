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
using OpenXmlPowerTools;
using Xunit;

#if !ELIDE_XUNIT_TESTS

namespace OxPt
{
    public class RpTests
    {
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        // perf settings
        public static bool m_CopySourceFilesToTempDir = true;
        public static bool m_OpenTempDirInExplorer = false;
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        [Theory]
        //[InlineData("RP/RP001-Tracked-Revisions-01.docx")]
        //[InlineData("RP/RP001-Tracked-Revisions-02.docx")]
        [InlineData("RP/RP002-Deleted-Text.docx")]
        [InlineData("RP/RP003-Inserted-Text.docx")]
        [InlineData("RP/RP004-Deleted-Text-in-CC.docx")]
        [InlineData("RP/RP005-Deleted-Paragraph-Mark.docx")]
        [InlineData("RP/RP006-Inserted-Paragraph-Mark.docx")]
        [InlineData("RP/RP007-Multiple-Deleted-Para-Mark.docx")]
        [InlineData("RP/RP008-Multiple-Inserted-Para-Mark.docx")]
        [InlineData("RP/RP009-Deleted-Table-Row.docx")]
        [InlineData("RP/RP010-Inserted-Table-Row.docx")]
        [InlineData("RP/RP011-Multiple-Deleted-Rows.docx")]
        [InlineData("RP/RP012-Multiple-Inserted-Rows.docx")]
        [InlineData("RP/RP013-Deleted-Math-Control-Char.docx")]
        [InlineData("RP/RP014-Inserted-Math-Control-Char.docx")]
        [InlineData("RP/RP015-MoveFrom-MoveTo.docx")]
        [InlineData("RP/RP016-Deleted-CC.docx")]
        [InlineData("RP/RP017-Inserted-CC.docx")]
        [InlineData("RP/RP018-MoveFrom-MoveTo-CC.docx")]
        [InlineData("RP/RP019-Deleted-Field-Code.docx")]
        [InlineData("RP/RP020-Inserted-Field-Code.docx")]
        [InlineData("RP/RP021-Inserted-Numbering-Properties.docx")]
        [InlineData("RP/RP022-NumberingChange.docx")]
        [InlineData("RP/RP023-NumberingChange.docx")]
        [InlineData("RP/RP024-ParagraphMark-rPr-Change.docx")]
        [InlineData("RP/RP025-Paragraph-Props-Change.docx")]
        [InlineData("RP/RP026-NumberingChange.docx")]
        [InlineData("RP/RP027-Change-Section.docx")]
        [InlineData("RP/RP028-Table-Grid-Change.docx")]
        [InlineData("RP/RP029-Table-Row-Props-Change.docx")]
        [InlineData("RP/RP030-Table-Row-Props-Change.docx")]
        [InlineData("RP/RP031-Table-Prop-Change.docx")]
        [InlineData("RP/RP032-Table-Prop-Change.docx")]
        [InlineData("RP/RP033-Table-Prop-Ex-Change.docx")]
        [InlineData("RP/RP034-Deleted-Cells.docx")]
        [InlineData("RP/RP035-Inserted-Cells.docx")]
        [InlineData("RP/RP036-Vert-Merged-Cells.docx")]
        [InlineData("RP/RP037-Changed-Style-Para-Props.docx")]
        [InlineData("RP/RP038-Inserted-Paras-at-End.docx")]
        [InlineData("RP/RP039-Inserted-Paras-at-End.docx")]
        [InlineData("RP/RP040-Deleted-Paras-at-End.docx")]
        [InlineData("RP/RP041-Cell-With-Empty-Paras-at-End.docx")]
        [InlineData("RP/RP042-Deleted-Para-Mark-at-End.docx")]
        [InlineData("RP/RP043-MERGEFORMAT-Field-Code.docx")]
        [InlineData("RP/RP044-MERGEFORMAT-Field-Code.docx")]
        [InlineData("RP/RP045-One-and-Half-Deleted-Lines-at-End.docx")]
        [InlineData("RP/RP046-Consecutive-Deleted-Ranges.docx")]
        [InlineData("RP/RP047-Inserted-and-Deleted-Paragraph-Mark.docx")]
        [InlineData("RP/RP048-Deleted-Inserted-Para-Mark.docx")]
        [InlineData("RP/RP049-Deleted-Para-Before-Table.docx")]
        [InlineData("RP/RP050-Deleted-Footnote.docx")]
        [InlineData("RP/RP052-Deleted-Para-Mark.docx")]
        public void RP001(string name)
        {
            var sourceFi = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name));
            var baselineAcceptedFi = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name.Replace(".docx", "-Accepted.docx")));
            var baselineRejectedFi = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name.Replace(".docx", "-Rejected.docx")));

            WmlDocument sourceWml = new WmlDocument(sourceFi.FullName);
            WmlDocument afterRejectingWml = RevisionProcessor.RejectRevisions(sourceWml);
            WmlDocument afterAcceptingWml = RevisionProcessor.AcceptRevisions(sourceWml);

            var processedAcceptedFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceFi.Name.Replace(".docx", "-Accepted.docx")));
            afterAcceptingWml.SaveAs(processedAcceptedFi.FullName);

            var processedRejectedFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceFi.Name.Replace(".docx", "-Rejected.docx")));
            afterRejectingWml.SaveAs(processedRejectedFi.FullName);

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Copy source files to temp dir
            if (m_CopySourceFilesToTempDir)
            {
                while (true)
                {
                    try
                    {
                        ////////// CODE TO REPEAT UNTIL SUCCESS //////////
                        var sourceDocxCopiedToDestFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceFi.Name));
                        if (!sourceDocxCopiedToDestFi.Exists)
                            sourceWml.SaveAs(sourceDocxCopiedToDestFi.FullName);
                        //////////////////////////////////////////////////
                        break;
                    }
                    catch (IOException)
                    {
                        System.Threading.Thread.Sleep(50);
                    }
                }
            }

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // create batch file to copy properly processed documents to the TestFiles directory.
            while (true)
            {
                try
                {
                    ////////// CODE TO REPEAT UNTIL SUCCESS //////////
                    var batchFileName = "Copy-Gen-Files-To-TestFiles.bat";
                    var batchFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, batchFileName));
                    var batch = "";
                    batch += "copy " + processedAcceptedFi.FullName + " " + baselineAcceptedFi.FullName + Environment.NewLine;
                    batch += "copy " + processedRejectedFi.FullName + " " + baselineRejectedFi.FullName + Environment.NewLine;
                    if (batchFi.Exists)
                        File.AppendAllText(batchFi.FullName, batch);
                    else
                        File.WriteAllText(batchFi.FullName, batch);
                    //////////////////////////////////////////////////
                    break;
                }
                catch (IOException)
                {
                    System.Threading.Thread.Sleep(50);
                }
            }

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Open Windows Explorer
            if (m_OpenTempDirInExplorer)
            {
                while (true)
                {
                    try
                    {
                        ////////// CODE TO REPEAT UNTIL SUCCESS //////////
                        var semaphorFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "z_ExplorerOpenedSemaphore.txt"));
                        if (!semaphorFi.Exists)
                        {
                            File.WriteAllText(semaphorFi.FullName, "");
                            TestUtil.Explorer(TestUtil.TempDir);
                        }
                        //////////////////////////////////////////////////
                        break;
                    }
                    catch (IOException)
                    {
                        System.Threading.Thread.Sleep(50);
                    }
                }
            }

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Use WmlComparer to see if accepted baseline is same as processed
            if (baselineAcceptedFi.Exists)
            {
                var baselineAcceptedWml = new WmlDocument(baselineAcceptedFi.FullName);
                WmlComparerSettings wmlComparerSettings = new WmlComparerSettings();
                WmlDocument result = WmlComparer.Compare(baselineAcceptedWml, afterAcceptingWml, wmlComparerSettings);
                var revisions = WmlComparer.GetRevisions(result, wmlComparerSettings);
                if (revisions.Any())
                {
                    Assert.True(false, "Regression Error: Accepted baseline document did not match processed document");
                }
            }
            else
            {
                Assert.True(false, "No Accepted baseline document");
            }

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Use WmlComparer to see if rejected baseline is same as processed
            if (baselineRejectedFi.Exists)
            {
                var baselineRejectedWml = new WmlDocument(baselineRejectedFi.FullName);
                WmlComparerSettings wmlComparerSettings = new WmlComparerSettings();
                WmlDocument result = WmlComparer.Compare(baselineRejectedWml, afterRejectingWml, wmlComparerSettings);
                var revisions = WmlComparer.GetRevisions(result, wmlComparerSettings);
                if (revisions.Any())
                {
                    Assert.True(false, "Regression Error: Rejected baseline document did not match processed document");
                }
            }
            else
            {
                Assert.True(false, "No Rejected baseline document");
            }
        }
    }
}

#endif
