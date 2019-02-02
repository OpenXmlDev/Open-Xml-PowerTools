// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static partial class WmlComparer
    {
        private static List<CorrelatedSequence> DoLcsAlgorithm(CorrelatedSequence unknown, WmlComparerSettings settings)
        {
            var newListOfCorrelatedSequence = new List<CorrelatedSequence>();

            ComparisonUnit[] cul1 = unknown.ComparisonUnitArray1;
            ComparisonUnit[] cul2 = unknown.ComparisonUnitArray2;

            // first thing to do - if we have an unknown with zero length on left or right side, create appropriate
            // this is a code optimization that enables easier processing of cases elsewhere.
            if (cul1.Length > 0 && cul2.Length == 0)
            {
                var deletedCorrelatedSequence = new CorrelatedSequence
                {
                    CorrelationStatus = CorrelationStatus.Deleted,
                    ComparisonUnitArray1 = cul1,
                    ComparisonUnitArray2 = null
                };
                newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);
                return newListOfCorrelatedSequence;
            }

            if (cul1.Length == 0 && cul2.Length > 0)
            {
                var insertedCorrelatedSequence = new CorrelatedSequence
                {
                    CorrelationStatus = CorrelationStatus.Inserted,
                    ComparisonUnitArray1 = null,
                    ComparisonUnitArray2 = cul2
                };
                newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
                return newListOfCorrelatedSequence;
            }

            if (cul1.Length == 0 && cul2.Length == 0)
            {
                // this will effectively remove the unknown with no data on either side from the current data model.
                return newListOfCorrelatedSequence;
            }

            var currentLongestCommonSequenceLength = 0;
            int currentI1 = -1;
            int currentI2 = -1;
            for (var i1 = 0; i1 < cul1.Length - currentLongestCommonSequenceLength; i1++)
            {
                for (var i2 = 0; i2 < cul2.Length - currentLongestCommonSequenceLength; i2++)
                {
                    var thisSequenceLength = 0;
                    int thisI1 = i1;
                    int thisI2 = i2;

                    while (true)
                    {
                        if (cul1[thisI1].SHA1Hash == cul2[thisI2].SHA1Hash)
                        {
                            thisI1++;
                            thisI2++;
                            thisSequenceLength++;

                            if (thisI1 == cul1.Length || thisI2 == cul2.Length)
                            {
                                if (thisSequenceLength > currentLongestCommonSequenceLength)
                                {
                                    currentLongestCommonSequenceLength = thisSequenceLength;
                                    currentI1 = i1;
                                    currentI2 = i2;
                                }

                                break;
                            }
                        }
                        else
                        {
                            if (thisSequenceLength > currentLongestCommonSequenceLength)
                            {
                                currentLongestCommonSequenceLength = thisSequenceLength;
                                currentI1 = i1;
                                currentI2 = i2;
                            }

                            break;
                        }
                    }
                }
            }

            // never start a common section with a paragraph mark.
            while (true)
            {
                if (currentLongestCommonSequenceLength <= 1)
                    break;

                ComparisonUnit firstCommon = cul1[currentI1];

                if (!(firstCommon is ComparisonUnitWord firstCommonWord))
                    break;

                // if the word contains more than one atom, then not a paragraph mark
                if (firstCommonWord.Contents.Count != 1)
                    break;

                if (!(firstCommonWord.Contents.First() is ComparisonUnitAtom firstCommonAtom))
                    break;

                if (firstCommonAtom.ContentElement.Name != W.pPr)
                    break;

                --currentLongestCommonSequenceLength;
                if (currentLongestCommonSequenceLength == 0)
                {
                    currentI1 = -1;
                    currentI2 = -1;
                }
                else
                {
                    ++currentI1;
                    ++currentI2;
                }
            }

            var isOnlyParagraphMark = false;
            if (currentLongestCommonSequenceLength == 1)
            {
                ComparisonUnit firstCommon = cul1[currentI1];

                if (firstCommon is ComparisonUnitWord firstCommonWord)
                {
                    // if the word contains more than one atom, then not a paragraph mark
                    if (firstCommonWord.Contents.Count == 1)
                    {
                        if (firstCommonWord.Contents.First() is ComparisonUnitAtom firstCommonAtom)
                        {
                            if (firstCommonAtom.ContentElement.Name == W.pPr)
                                isOnlyParagraphMark = true;
                        }
                    }
                }
            }

            // don't match just a single character
            if (currentLongestCommonSequenceLength == 1)
            {
                if (cul2[currentI2] is ComparisonUnitAtom cuw2)
                {
                    if (cuw2.ContentElement.Name == W.t && cuw2.ContentElement.Value == " ")
                    {
                        currentI1 = -1;
                        currentI2 = -1;
                        currentLongestCommonSequenceLength = 0;
                    }
                }
            }

            // don't match only word break characters
            if (currentLongestCommonSequenceLength > 0 && currentLongestCommonSequenceLength <= 3)
            {
                ComparisonUnit[] commonSequence = cul1.Skip(currentI1).Take(currentLongestCommonSequenceLength).ToArray();

                // if they are all ComparisonUnitWord objects
                bool oneIsNotWord = commonSequence.Any(cs => !(cs is ComparisonUnitWord));
                bool allAreWords = !oneIsNotWord;
                if (allAreWords)
                {
                    bool contentOtherThanWordSplitChars = commonSequence
                        .Cast<ComparisonUnitWord>()
                        .Any(cs =>
                        {
                            bool otherThanText = cs.DescendantContentAtoms().Any(dca => dca.ContentElement.Name != W.t);
                            if (otherThanText) return true;

                            bool otherThanWordSplit = cs
                                .DescendantContentAtoms()
                                .Any(dca =>
                                {
                                    string charValue = dca.ContentElement.Value;
                                    bool isWordSplit = settings.WordSeparators.Contains(charValue[0]);
                                    return !isWordSplit;
                                });

                            return otherThanWordSplit;
                        });
                    if (!contentOtherThanWordSplitChars)
                    {
                        currentI1 = -1;
                        currentI2 = -1;
                        currentLongestCommonSequenceLength = 0;
                    }
                }
            }

            // if we are only looking at text, and if the longest common subsequence is less than 15% of the whole, then forget it,
            // don't find that LCS.
            if (!isOnlyParagraphMark && currentLongestCommonSequenceLength > 0)
            {
                bool anyButWord1 = cul1.Any(cu => !(cu is ComparisonUnitWord));
                bool anyButWord2 = cul2.Any(cu => !(cu is ComparisonUnitWord));

                if (!anyButWord1 && !anyButWord2)
                {
                    int maxLen = Math.Max(cul1.Length, cul2.Length);
                    if (currentLongestCommonSequenceLength / (double) maxLen < settings.DetailThreshold)
                    {
                        currentI1 = -1;
                        currentI2 = -1;
                        currentLongestCommonSequenceLength = 0;
                    }
                }
            }

            if (currentI1 == -1 && currentI2 == -1)
            {
                int leftLength = unknown.ComparisonUnitArray1.Length;

                int leftTables = unknown
                    .ComparisonUnitArray1
                    .OfType<ComparisonUnitGroup>()
                    .Count(l => l.ComparisonUnitGroupType == ComparisonUnitGroupType.Table);

                int leftRows = unknown
                    .ComparisonUnitArray1
                    .OfType<ComparisonUnitGroup>()
                    .Count(l => l.ComparisonUnitGroupType == ComparisonUnitGroupType.Row);

                int leftParagraphs = unknown
                    .ComparisonUnitArray1
                    .OfType<ComparisonUnitGroup>()
                    .Count(l => l.ComparisonUnitGroupType == ComparisonUnitGroupType.Paragraph);

                int leftTextboxes = unknown
                    .ComparisonUnitArray1
                    .OfType<ComparisonUnitGroup>()
                    .Count(l => l.ComparisonUnitGroupType == ComparisonUnitGroupType.Textbox);

                int leftWords = unknown
                    .ComparisonUnitArray1
                    .OfType<ComparisonUnitWord>()
                    .Count();

                int rightLength = unknown.ComparisonUnitArray2.Length;

                int rightTables = unknown
                    .ComparisonUnitArray2
                    .OfType<ComparisonUnitGroup>()
                    .Count(l => l.ComparisonUnitGroupType == ComparisonUnitGroupType.Table);

                int rightRows = unknown
                    .ComparisonUnitArray2
                    .OfType<ComparisonUnitGroup>()
                    .Count(l => l.ComparisonUnitGroupType == ComparisonUnitGroupType.Row);

                int rightParagraphs = unknown
                    .ComparisonUnitArray2
                    .OfType<ComparisonUnitGroup>()
                    .Count(l => l.ComparisonUnitGroupType == ComparisonUnitGroupType.Paragraph);

                int rightTextboxes = unknown
                    .ComparisonUnitArray2
                    .OfType<ComparisonUnitGroup>()
                    .Count(l => l.ComparisonUnitGroupType == ComparisonUnitGroupType.Textbox);

                int rightWords = unknown
                    .ComparisonUnitArray2
                    .OfType<ComparisonUnitWord>()
                    .Count();

                // if either side has both words, rows and text boxes, then we need to separate out into separate
                // unknown correlated sequences
                // group adjacent based on whether word, row, or textbox
                // in most cases, the count of groups will be the same, but they may differ
                // if the first group on either side is word, then create a deleted or inserted corr sequ for it.
                // then have counter on both sides pointing to the first matched pairs of rows
                // create an unknown corr sequ for it.
                // increment both counters
                // if one is at end but the other is not, then tag the remaining content as inserted or deleted, and done.
                // if both are at the end, then done
                // return the new list of corr sequ

                bool leftOnlyWordsRowsTextboxes = leftLength == leftWords + leftRows + leftTextboxes;
                bool rightOnlyWordsRowsTextboxes = rightLength == rightWords + rightRows + rightTextboxes;
                if ((leftWords > 0 || rightWords > 0) &&
                    (leftRows > 0 || rightRows > 0 || leftTextboxes > 0 || rightTextboxes > 0) &&
                    leftOnlyWordsRowsTextboxes &&
                    rightOnlyWordsRowsTextboxes)
                {
                    IGrouping<string, ComparisonUnit>[] leftGrouped = unknown
                        .ComparisonUnitArray1
                        .GroupAdjacent(cu =>
                        {
                            if (cu is ComparisonUnitWord)
                            {
                                return "Word";
                            }

                            var cug = cu as ComparisonUnitGroup;
                            if (cug?.ComparisonUnitGroupType == ComparisonUnitGroupType.Row)
                            {
                                return "Row";
                            }

                            if (cug?.ComparisonUnitGroupType == ComparisonUnitGroupType.Textbox)
                            {
                                return "Textbox";
                            }

                            throw new OpenXmlPowerToolsException("Internal error");
                        })
                        .ToArray();

                    IGrouping<string, ComparisonUnit>[] rightGrouped = unknown
                        .ComparisonUnitArray2
                        .GroupAdjacent(cu =>
                        {
                            if (cu is ComparisonUnitWord)
                            {
                                return "Word";
                            }

                            var cug = cu as ComparisonUnitGroup;
                            if (cug?.ComparisonUnitGroupType == ComparisonUnitGroupType.Row)
                            {
                                return "Row";
                            }

                            if (cug?.ComparisonUnitGroupType == ComparisonUnitGroupType.Textbox)
                            {
                                return "Textbox";
                            }

                            throw new OpenXmlPowerToolsException("Internal error");
                        })
                        .ToArray();

                    var iLeft = 0;
                    var iRight = 0;

                    // create an unknown corr sequ for it.
                    // increment both counters
                    // if one is at end but the other is not, then tag the remaining content as inserted or deleted, and done.
                    // if both are at the end, then done
                    // return the new list of corr sequ

                    while (true)
                    {
                        if (leftGrouped[iLeft].Key == rightGrouped[iRight].Key)
                        {
                            var unknownCorrelatedSequence = new CorrelatedSequence
                            {
                                ComparisonUnitArray1 = leftGrouped[iLeft].ToArray(),
                                ComparisonUnitArray2 = rightGrouped[iRight].ToArray(),
                                CorrelationStatus = CorrelationStatus.Unknown
                            };
                            newListOfCorrelatedSequence.Add(unknownCorrelatedSequence);
                            ++iLeft;
                            ++iRight;
                        }

                        // have to decide which of the following two branches to do first based on whether the left contains a paragraph mark
                        // i.e. cant insert a string of deleted text right before a table.

                        else if (leftGrouped[iLeft].Key == "Word" &&
                                 leftGrouped[iLeft]
                                     .Select(lg => lg.DescendantContentAtoms())
                                     .SelectMany(m => m).Last()
                                     .ContentElement
                                     .Name != W.pPr &&
                                 rightGrouped[iRight].Key == "Row")
                        {
                            var insertedCorrelatedSequence = new CorrelatedSequence
                            {
                                ComparisonUnitArray1 = null,
                                ComparisonUnitArray2 = rightGrouped[iRight].ToArray(),
                                CorrelationStatus = CorrelationStatus.Inserted
                            };
                            newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
                            ++iRight;
                        }
                        else if (rightGrouped[iRight].Key == "Word" &&
                                 rightGrouped[iRight]
                                     .Select(lg => lg.DescendantContentAtoms())
                                     .SelectMany(m => m)
                                     .Last()
                                     .ContentElement
                                     .Name != W.pPr &&
                                 leftGrouped[iLeft].Key == "Row")
                        {
                            var insertedCorrelatedSequence = new CorrelatedSequence
                            {
                                ComparisonUnitArray1 = null,
                                ComparisonUnitArray2 = leftGrouped[iLeft].ToArray(),
                                CorrelationStatus = CorrelationStatus.Inserted
                            };
                            newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
                            ++iLeft;
                        }

                        else if (leftGrouped[iLeft].Key == "Word" && rightGrouped[iRight].Key != "Word")
                        {
                            var deletedCorrelatedSequence = new CorrelatedSequence
                            {
                                ComparisonUnitArray1 = leftGrouped[iLeft].ToArray(),
                                ComparisonUnitArray2 = null,
                                CorrelationStatus = CorrelationStatus.Deleted
                            };
                            newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);
                            ++iLeft;
                        }

                        else if (leftGrouped[iLeft].Key != "Word" && rightGrouped[iRight].Key == "Word")
                        {
                            var insertedCorrelatedSequence = new CorrelatedSequence
                            {
                                ComparisonUnitArray1 = null,
                                ComparisonUnitArray2 = rightGrouped[iRight].ToArray(),
                                CorrelationStatus = CorrelationStatus.Inserted
                            };
                            newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
                            ++iRight;
                        }

                        if (iLeft == leftGrouped.Length && iRight == rightGrouped.Length)
                            return newListOfCorrelatedSequence;

                        // if there is content on the left, but not content on the right
                        if (iRight == rightGrouped.Length)
                        {
                            for (int j = iLeft; j < leftGrouped.Length; j++)
                            {
                                var deletedCorrelatedSequence = new CorrelatedSequence
                                {
                                    ComparisonUnitArray1 = leftGrouped[j].ToArray(),
                                    ComparisonUnitArray2 = null,
                                    CorrelationStatus = CorrelationStatus.Deleted
                                };
                                newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);
                            }

                            return newListOfCorrelatedSequence;
                        }

                        // there is content on the right but not on the left

                        if (iLeft == leftGrouped.Length)
                        {
                            for (int j = iRight; j < rightGrouped.Length; j++)
                            {
                                var insertedCorrelatedSequence = new CorrelatedSequence
                                {
                                    ComparisonUnitArray1 = null,
                                    ComparisonUnitArray2 = rightGrouped[j].ToArray(),
                                    CorrelationStatus = CorrelationStatus.Inserted
                                };
                                newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
                            }

                            return newListOfCorrelatedSequence;
                        }

                        // else continue on next round.
                    }
                }

                // if both sides contain tables and paragraphs, then split into multiple unknown corr sequ
                if (leftTables > 0 && rightTables > 0 &&
                    leftParagraphs > 0 && rightParagraphs > 0 &&
                    (leftLength > 1 || rightLength > 1))
                {
                    IGrouping<string, ComparisonUnit>[] leftGrouped = unknown
                        .ComparisonUnitArray1
                        .GroupAdjacent(cu =>
                        {
                            var cug = (ComparisonUnitGroup) cu;
                            return cug.ComparisonUnitGroupType == ComparisonUnitGroupType.Table ? "Table" : "Para";
                        })
                        .ToArray();

                    IGrouping<string, ComparisonUnit>[] rightGrouped = unknown
                        .ComparisonUnitArray2
                        .GroupAdjacent(cu =>
                        {
                            var cug = (ComparisonUnitGroup) cu;
                            return cug.ComparisonUnitGroupType == ComparisonUnitGroupType.Table ? "Table" : "Para";
                        })
                        .ToArray();

                    var iLeft = 0;
                    var iRight = 0;

                    // create an unknown corr sequ for it.
                    // increment both counters
                    // if one is at end but the other is not, then tag the remaining content as inserted or deleted, and done.
                    // if both are at the end, then done
                    // return the new list of corr sequ

                    while (true)
                    {
                        if (leftGrouped[iLeft].Key == "Table" && rightGrouped[iRight].Key == "Table" ||
                            leftGrouped[iLeft].Key == "Para" && rightGrouped[iRight].Key == "Para")
                        {
                            var unknownCorrelatedSequence = new CorrelatedSequence
                            {
                                ComparisonUnitArray1 = leftGrouped[iLeft].ToArray(),
                                ComparisonUnitArray2 = rightGrouped[iRight].ToArray(),
                                CorrelationStatus = CorrelationStatus.Unknown
                            };
                            newListOfCorrelatedSequence.Add(unknownCorrelatedSequence);
                            ++iLeft;
                            ++iRight;
                        }
                        else if (leftGrouped[iLeft].Key == "Para" && rightGrouped[iRight].Key == "Table")
                        {
                            var deletedCorrelatedSequence = new CorrelatedSequence
                            {
                                ComparisonUnitArray1 = leftGrouped[iLeft].ToArray(),
                                ComparisonUnitArray2 = null,
                                CorrelationStatus = CorrelationStatus.Deleted
                            };
                            newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);
                            ++iLeft;
                        }
                        else if (leftGrouped[iLeft].Key == "Table" && rightGrouped[iRight].Key == "Para")
                        {
                            var insertedCorrelatedSequence = new CorrelatedSequence
                            {
                                ComparisonUnitArray1 = null,
                                ComparisonUnitArray2 = rightGrouped[iRight].ToArray(),
                                CorrelationStatus = CorrelationStatus.Inserted
                            };
                            newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
                            ++iRight;
                        }

                        if (iLeft == leftGrouped.Length && iRight == rightGrouped.Length)
                            return newListOfCorrelatedSequence;

                        // if there is content on the left, but not content on the right
                        if (iRight == rightGrouped.Length)
                        {
                            for (int j = iLeft; j < leftGrouped.Length; j++)
                            {
                                var deletedCorrelatedSequence = new CorrelatedSequence
                                {
                                    ComparisonUnitArray1 = leftGrouped[j].ToArray(),
                                    ComparisonUnitArray2 = null,
                                    CorrelationStatus = CorrelationStatus.Deleted
                                };
                                newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);
                            }

                            return newListOfCorrelatedSequence;
                        }

                        // there is content on the right but not on the left

                        if (iLeft == leftGrouped.Length)
                        {
                            for (int j = iRight; j < rightGrouped.Length; j++)
                            {
                                var insertedCorrelatedSequence = new CorrelatedSequence
                                {
                                    ComparisonUnitArray1 = null,
                                    ComparisonUnitArray2 = rightGrouped[j].ToArray(),
                                    CorrelationStatus = CorrelationStatus.Inserted
                                };
                                newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
                            }

                            return newListOfCorrelatedSequence;
                        }

                        // else continue on next round.
                    }
                }

                // If both sides consists of a single table, and if the table contains merged cells, then mark as deleted/inserted
                if (leftTables == 1 && leftLength == 1 &&
                    rightTables == 1 && rightLength == 1)
                {
                    List<CorrelatedSequence> result = DoLcsAlgorithmForTable(unknown);
                    if (result != null)
                        return result;
                }

                // If either side contains only paras or tables, then flatten and iterate.
                bool leftOnlyParasTablesTextboxes = leftLength == leftTables + leftParagraphs + leftTextboxes;
                bool rightOnlyParasTablesTextboxes = rightLength == rightTables + rightParagraphs + rightTextboxes;
                if (leftOnlyParasTablesTextboxes && rightOnlyParasTablesTextboxes)
                {
                    // flatten paras and tables, and iterate
                    ComparisonUnit[] left = unknown
                        .ComparisonUnitArray1
                        .Select(cu => cu.Contents)
                        .SelectMany(m => m)
                        .ToArray();

                    ComparisonUnit[] right = unknown
                        .ComparisonUnitArray2
                        .Select(cu => cu.Contents)
                        .SelectMany(m => m)
                        .ToArray();

                    var unknownCorrelatedSequence = new CorrelatedSequence
                    {
                        CorrelationStatus = CorrelationStatus.Unknown,
                        ComparisonUnitArray1 = left,
                        ComparisonUnitArray2 = right
                    };
                    newListOfCorrelatedSequence.Add(unknownCorrelatedSequence);

                    return newListOfCorrelatedSequence;
                }

                // if first of left is a row and first of right is a row
                // then flatten the row to cells and iterate.

                if (unknown.ComparisonUnitArray1.FirstOrDefault() is ComparisonUnitGroup firstLeft &&
                    unknown.ComparisonUnitArray2.FirstOrDefault() is ComparisonUnitGroup firstRight)
                {
                    if (firstLeft.ComparisonUnitGroupType == ComparisonUnitGroupType.Row &&
                        firstRight.ComparisonUnitGroupType == ComparisonUnitGroupType.Row)
                    {
                        ComparisonUnit[] leftContent = firstLeft.Contents.ToArray();
                        ComparisonUnit[] rightContent = firstRight.Contents.ToArray();

                        int lenLeft = leftContent.Length;
                        int lenRight = rightContent.Length;

                        if (lenLeft < lenRight)
                        {
                            leftContent = leftContent
                                .Concat(Enumerable.Repeat<ComparisonUnit>(null, lenRight - lenLeft))
                                .ToArray();
                        }
                        else if (lenRight < lenLeft)
                        {
                            rightContent = rightContent
                                .Concat(Enumerable.Repeat<ComparisonUnit>(null, lenLeft - lenRight))
                                .ToArray();
                        }

                        List<CorrelatedSequence> newCs = leftContent.Zip(rightContent, (l, r) =>
                            {
                                if (l != null && r != null)
                                {
                                    var unknownCorrelatedSequence = new CorrelatedSequence
                                    {
                                        ComparisonUnitArray1 = new[] { l },
                                        ComparisonUnitArray2 = new[] { r },
                                        CorrelationStatus = CorrelationStatus.Unknown
                                    };
                                    return new[] { unknownCorrelatedSequence };
                                }

                                if (l == null)
                                {
                                    var insertedCorrelatedSequence = new CorrelatedSequence
                                    {
                                        ComparisonUnitArray1 = null,
                                        ComparisonUnitArray2 = r.Contents.ToArray(),
                                        CorrelationStatus = CorrelationStatus.Inserted
                                    };
                                    return new[] { insertedCorrelatedSequence };
                                }

                                var deletedCorrelatedSequence = new CorrelatedSequence
                                {
                                    ComparisonUnitArray1 = l.Contents.ToArray(),
                                    ComparisonUnitArray2 = null,
                                    CorrelationStatus = CorrelationStatus.Deleted
                                };
                                return new[] { deletedCorrelatedSequence };
                            })
                            .SelectMany(m => m)
                            .ToList();

                        foreach (CorrelatedSequence cs in newCs)
                        {
                            newListOfCorrelatedSequence.Add(cs);
                        }

                        ComparisonUnit[] remainderLeft = unknown
                            .ComparisonUnitArray1
                            .Skip(1)
                            .ToArray();

                        ComparisonUnit[] remainderRight = unknown
                            .ComparisonUnitArray2
                            .Skip(1)
                            .ToArray();

                        if (remainderLeft.Length > 0 && remainderRight.Length == 0)
                        {
                            var deletedCorrelatedSequence = new CorrelatedSequence
                            {
                                CorrelationStatus = CorrelationStatus.Deleted,
                                ComparisonUnitArray1 = remainderLeft,
                                ComparisonUnitArray2 = null
                            };
                            newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);
                        }
                        else if (remainderRight.Length > 0 && remainderLeft.Length == 0)
                        {
                            var insertedCorrelatedSequence = new CorrelatedSequence
                            {
                                CorrelationStatus = CorrelationStatus.Inserted,
                                ComparisonUnitArray1 = null,
                                ComparisonUnitArray2 = remainderRight
                            };
                            newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
                        }
                        else if (remainderLeft.Length > 0 && remainderRight.Length > 0)
                        {
                            var unknownCorrelatedSequence2 = new CorrelatedSequence
                            {
                                CorrelationStatus = CorrelationStatus.Unknown,
                                ComparisonUnitArray1 = remainderLeft,
                                ComparisonUnitArray2 = remainderRight
                            };
                            newListOfCorrelatedSequence.Add(unknownCorrelatedSequence2);
                        }

                        if (False)
                        {
                            var sb = new StringBuilder();
                            foreach (CorrelatedSequence item in newListOfCorrelatedSequence)
                            {
                                sb.Append(item).Append(Environment.NewLine);
                            }

                            string sbs = sb.ToString();
                            TestUtil.NotePad(sbs);
                        }

                        return newListOfCorrelatedSequence;
                    }

                    if (firstLeft.ComparisonUnitGroupType == ComparisonUnitGroupType.Cell &&
                        firstRight.ComparisonUnitGroupType == ComparisonUnitGroupType.Cell)
                    {
                        ComparisonUnit[] left = firstLeft
                            .Contents
                            .ToArray();

                        ComparisonUnit[] right = firstRight
                            .Contents
                            .ToArray();

                        var unknownCorrelatedSequence = new CorrelatedSequence
                        {
                            CorrelationStatus = CorrelationStatus.Unknown,
                            ComparisonUnitArray1 = left,
                            ComparisonUnitArray2 = right
                        };
                        newListOfCorrelatedSequence.Add(unknownCorrelatedSequence);

                        ComparisonUnit[] remainderLeft = unknown
                            .ComparisonUnitArray1
                            .Skip(1)
                            .ToArray();

                        ComparisonUnit[] remainderRight = unknown
                            .ComparisonUnitArray2
                            .Skip(1)
                            .ToArray();

                        if (remainderLeft.Length > 0 && remainderRight.Length == 0)
                        {
                            var deletedCorrelatedSequence = new CorrelatedSequence
                            {
                                CorrelationStatus = CorrelationStatus.Deleted,
                                ComparisonUnitArray1 = remainderLeft,
                                ComparisonUnitArray2 = null
                            };
                            newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);
                        }
                        else if (remainderRight.Length > 0 && remainderLeft.Length == 0)
                        {
                            var insertedCorrelatedSequence = new CorrelatedSequence
                            {
                                CorrelationStatus = CorrelationStatus.Inserted,
                                ComparisonUnitArray1 = null,
                                ComparisonUnitArray2 = remainderRight
                            };
                            newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
                        }
                        else if (remainderLeft.Length > 0 && remainderRight.Length > 0)
                        {
                            var unknownCorrelatedSequence2 = new CorrelatedSequence
                            {
                                CorrelationStatus = CorrelationStatus.Unknown,
                                ComparisonUnitArray1 = remainderLeft,
                                ComparisonUnitArray2 = remainderRight
                            };
                            newListOfCorrelatedSequence.Add(unknownCorrelatedSequence2);
                        }

                        return newListOfCorrelatedSequence;
                    }
                }

                if (unknown.ComparisonUnitArray1.Any() && unknown.ComparisonUnitArray2.Any())
                {
                    if (unknown.ComparisonUnitArray1.First() is ComparisonUnitWord &&
                        unknown.ComparisonUnitArray2.First() is ComparisonUnitGroup right &&
                        right.ComparisonUnitGroupType == ComparisonUnitGroupType.Row)
                    {
                        var insertedCorrelatedSequence3 = new CorrelatedSequence
                        {
                            CorrelationStatus = CorrelationStatus.Inserted,
                            ComparisonUnitArray1 = null,
                            ComparisonUnitArray2 = unknown.ComparisonUnitArray2
                        };
                        newListOfCorrelatedSequence.Add(insertedCorrelatedSequence3);

                        var deletedCorrelatedSequence3 = new CorrelatedSequence
                        {
                            CorrelationStatus = CorrelationStatus.Deleted,
                            ComparisonUnitArray1 = unknown.ComparisonUnitArray1,
                            ComparisonUnitArray2 = null
                        };
                        newListOfCorrelatedSequence.Add(deletedCorrelatedSequence3);

                        return newListOfCorrelatedSequence;
                    }

                    if (unknown.ComparisonUnitArray2.First() is ComparisonUnitWord &&
                        unknown.ComparisonUnitArray1.First() is ComparisonUnitGroup left2 &&
                        left2.ComparisonUnitGroupType == ComparisonUnitGroupType.Row)
                    {
                        var deletedCorrelatedSequence3 = new CorrelatedSequence
                        {
                            CorrelationStatus = CorrelationStatus.Deleted,
                            ComparisonUnitArray1 = unknown.ComparisonUnitArray1,
                            ComparisonUnitArray2 = null
                        };
                        newListOfCorrelatedSequence.Add(deletedCorrelatedSequence3);

                        var insertedCorrelatedSequence3 = new CorrelatedSequence
                        {
                            CorrelationStatus = CorrelationStatus.Inserted,
                            ComparisonUnitArray1 = null,
                            ComparisonUnitArray2 = unknown.ComparisonUnitArray2
                        };
                        newListOfCorrelatedSequence.Add(insertedCorrelatedSequence3);

                        return newListOfCorrelatedSequence;
                    }

                    ComparisonUnitAtom lastContentAtomLeft = unknown
                        .ComparisonUnitArray1
                        .Select(cu => cu.DescendantContentAtoms().Last())
                        .LastOrDefault();

                    ComparisonUnitAtom lastContentAtomRight = unknown
                        .ComparisonUnitArray2
                        .Select(cu => cu.DescendantContentAtoms().Last())
                        .LastOrDefault();

                    if (lastContentAtomLeft != null && lastContentAtomRight != null)
                    {
                        if (lastContentAtomLeft.ContentElement.Name == W.pPr &&
                            lastContentAtomRight.ContentElement.Name != W.pPr)
                        {
                            var insertedCorrelatedSequence5 = new CorrelatedSequence
                            {
                                CorrelationStatus = CorrelationStatus.Inserted,
                                ComparisonUnitArray1 = null,
                                ComparisonUnitArray2 = unknown.ComparisonUnitArray2
                            };
                            newListOfCorrelatedSequence.Add(insertedCorrelatedSequence5);

                            var deletedCorrelatedSequence5 = new CorrelatedSequence
                            {
                                CorrelationStatus = CorrelationStatus.Deleted,
                                ComparisonUnitArray1 = unknown.ComparisonUnitArray1,
                                ComparisonUnitArray2 = null
                            };
                            newListOfCorrelatedSequence.Add(deletedCorrelatedSequence5);

                            return newListOfCorrelatedSequence;
                        }

                        if (lastContentAtomLeft.ContentElement.Name != W.pPr &&
                            lastContentAtomRight.ContentElement.Name == W.pPr)
                        {
                            var deletedCorrelatedSequence5 = new CorrelatedSequence
                            {
                                CorrelationStatus = CorrelationStatus.Deleted,
                                ComparisonUnitArray1 = unknown.ComparisonUnitArray1,
                                ComparisonUnitArray2 = null
                            };
                            newListOfCorrelatedSequence.Add(deletedCorrelatedSequence5);

                            var insertedCorrelatedSequence5 = new CorrelatedSequence
                            {
                                CorrelationStatus = CorrelationStatus.Inserted,
                                ComparisonUnitArray1 = null,
                                ComparisonUnitArray2 = unknown.ComparisonUnitArray2
                            };
                            newListOfCorrelatedSequence.Add(insertedCorrelatedSequence5);

                            return newListOfCorrelatedSequence;
                        }
                    }
                }

                var deletedCorrelatedSequence4 = new CorrelatedSequence
                {
                    CorrelationStatus = CorrelationStatus.Deleted,
                    ComparisonUnitArray1 = unknown.ComparisonUnitArray1,
                    ComparisonUnitArray2 = null
                };
                newListOfCorrelatedSequence.Add(deletedCorrelatedSequence4);

                var insertedCorrelatedSequence4 = new CorrelatedSequence
                {
                    CorrelationStatus = CorrelationStatus.Inserted,
                    ComparisonUnitArray1 = null,
                    ComparisonUnitArray2 = unknown.ComparisonUnitArray2
                };
                newListOfCorrelatedSequence.Add(insertedCorrelatedSequence4);

                return newListOfCorrelatedSequence;
            }

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////
            // here we have the longest common subsequence.
            // but it may start in the middle of a paragraph.
            // therefore need to dispose of the content from the beginning of the longest common subsequence to the
            // beginning of the paragraph.
            // this should be in a separate unknown region
            // if countCommonAtEnd != 0, and if it contains a paragraph mark, then if there are comparison units in
            // the same paragraph before the common at end (in either version)
            // then we want to put all of those comparison units into a single unknown, where they must be resolved
            // against each other.  We don't want those comparison units to go into the middle unknown comparison unit.

            var remainingInLeftParagraph = 0;
            var remainingInRightParagraph = 0;
            if (currentLongestCommonSequenceLength != 0)
            {
                List<ComparisonUnit> commonSeq = unknown
                    .ComparisonUnitArray1
                    .Skip(currentI1)
                    .Take(currentLongestCommonSequenceLength)
                    .ToList();

                ComparisonUnit firstOfCommonSeq = commonSeq.First();
                if (firstOfCommonSeq is ComparisonUnitWord)
                {
                    // are there any paragraph marks in the common seq at end?
                    if (commonSeq.Any(cu =>
                    {
                        ComparisonUnitAtom firstComparisonUnitAtom = cu.Contents.OfType<ComparisonUnitAtom>().FirstOrDefault();
                        if (firstComparisonUnitAtom == null)
                            return false;

                        return firstComparisonUnitAtom.ContentElement.Name == W.pPr;
                    }))
                    {
                        remainingInLeftParagraph = unknown
                            .ComparisonUnitArray1
                            .Take(currentI1)
                            .Reverse()
                            .TakeWhile(cu =>
                            {
                                if (!(cu is ComparisonUnitWord))
                                    return false;

                                ComparisonUnitAtom firstComparisonUnitAtom =
                                    cu.Contents.OfType<ComparisonUnitAtom>().FirstOrDefault();
                                if (firstComparisonUnitAtom == null)
                                    return true;

                                return firstComparisonUnitAtom.ContentElement.Name != W.pPr;
                            })
                            .Count();
                        remainingInRightParagraph = unknown
                            .ComparisonUnitArray2
                            .Take(currentI2)
                            .Reverse()
                            .TakeWhile(cu =>
                            {
                                if (!(cu is ComparisonUnitWord))
                                    return false;

                                ComparisonUnitAtom firstComparisonUnitAtom =
                                    cu.Contents.OfType<ComparisonUnitAtom>().FirstOrDefault();
                                if (firstComparisonUnitAtom == null)
                                    return true;

                                return firstComparisonUnitAtom.ContentElement.Name != W.pPr;
                            })
                            .Count();
                    }
                }
            }

            int countBeforeCurrentParagraphLeft = currentI1 - remainingInLeftParagraph;
            int countBeforeCurrentParagraphRight = currentI2 - remainingInRightParagraph;

            if (countBeforeCurrentParagraphLeft > 0 && countBeforeCurrentParagraphRight == 0)
            {
                var deletedCorrelatedSequence = new CorrelatedSequence
                {
                    CorrelationStatus = CorrelationStatus.Deleted,
                    ComparisonUnitArray1 = cul1
                        .Take(countBeforeCurrentParagraphLeft)
                        .ToArray(),
                    ComparisonUnitArray2 = null
                };
                newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);
            }
            else if (countBeforeCurrentParagraphLeft == 0 && countBeforeCurrentParagraphRight > 0)
            {
                var insertedCorrelatedSequence = new CorrelatedSequence
                {
                    CorrelationStatus = CorrelationStatus.Inserted,
                    ComparisonUnitArray1 = null,
                    ComparisonUnitArray2 = cul2
                        .Take(countBeforeCurrentParagraphRight)
                        .ToArray()
                };
                newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
            }
            else if (countBeforeCurrentParagraphLeft > 0 && countBeforeCurrentParagraphRight > 0)
            {
                var unknownCorrelatedSequence = new CorrelatedSequence
                {
                    CorrelationStatus = CorrelationStatus.Unknown,
                    ComparisonUnitArray1 = cul1
                        .Take(countBeforeCurrentParagraphLeft)
                        .ToArray(),
                    ComparisonUnitArray2 = cul2
                        .Take(countBeforeCurrentParagraphRight)
                        .ToArray()
                };

                newListOfCorrelatedSequence.Add(unknownCorrelatedSequence);
            }
            else if (countBeforeCurrentParagraphLeft == 0 && countBeforeCurrentParagraphRight == 0)
            {
                // nothing to do
            }

            if (remainingInLeftParagraph > 0 && remainingInRightParagraph == 0)
            {
                var deletedCorrelatedSequence = new CorrelatedSequence
                {
                    CorrelationStatus = CorrelationStatus.Deleted,
                    ComparisonUnitArray1 = cul1
                        .Skip(countBeforeCurrentParagraphLeft)
                        .Take(remainingInLeftParagraph)
                        .ToArray(),
                    ComparisonUnitArray2 = null
                };
                newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);
            }
            else if (remainingInLeftParagraph == 0 && remainingInRightParagraph > 0)
            {
                var insertedCorrelatedSequence = new CorrelatedSequence
                {
                    CorrelationStatus = CorrelationStatus.Inserted,
                    ComparisonUnitArray1 = null,
                    ComparisonUnitArray2 = cul2
                        .Skip(countBeforeCurrentParagraphRight)
                        .Take(remainingInRightParagraph)
                        .ToArray()
                };
                newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
            }
            else if (remainingInLeftParagraph > 0 && remainingInRightParagraph > 0)
            {
                var unknownCorrelatedSequence = new CorrelatedSequence
                {
                    CorrelationStatus = CorrelationStatus.Unknown,
                    ComparisonUnitArray1 = cul1
                        .Skip(countBeforeCurrentParagraphLeft)
                        .Take(remainingInLeftParagraph)
                        .ToArray(),
                    ComparisonUnitArray2 = cul2
                        .Skip(countBeforeCurrentParagraphRight)
                        .Take(remainingInRightParagraph)
                        .ToArray()
                };
                newListOfCorrelatedSequence.Add(unknownCorrelatedSequence);
            }
            else if (remainingInLeftParagraph == 0 && remainingInRightParagraph == 0)
            {
                // nothing to do
            }

            var middleEqual = new CorrelatedSequence
            {
                CorrelationStatus = CorrelationStatus.Equal,
                ComparisonUnitArray1 = cul1
                    .Skip(currentI1)
                    .Take(currentLongestCommonSequenceLength)
                    .ToArray(),
                ComparisonUnitArray2 = cul2
                    .Skip(currentI2)
                    .Take(currentLongestCommonSequenceLength)
                    .ToArray()
            };
            newListOfCorrelatedSequence.Add(middleEqual);


            int endI1 = currentI1 + currentLongestCommonSequenceLength;
            int endI2 = currentI2 + currentLongestCommonSequenceLength;

            ComparisonUnit[] remaining1 = cul1
                .Skip(endI1)
                .ToArray();

            ComparisonUnit[] remaining2 = cul2
                .Skip(endI2)
                .ToArray();

            // here is the point that we want to make a new unknown from this point to the end of the paragraph that
            // contains the equal parts.
            // this will never hurt anything, and will in many cases result in a better difference.

            if (middleEqual.ComparisonUnitArray1[middleEqual.ComparisonUnitArray1.Length - 1] is ComparisonUnitWord leftCuw)
            {
                ComparisonUnitAtom lastContentAtom = leftCuw.DescendantContentAtoms().LastOrDefault();

                // if the middleEqual did not end with a paragraph mark
                if (lastContentAtom != null && lastContentAtom.ContentElement.Name != W.pPr)
                {
                    int idx1 = FindIndexOfNextParaMark(remaining1);
                    int idx2 = FindIndexOfNextParaMark(remaining2);

                    var unknownCorrelatedSequenceRemaining = new CorrelatedSequence
                    {
                        CorrelationStatus = CorrelationStatus.Unknown,
                        ComparisonUnitArray1 = remaining1.Take(idx1).ToArray(),
                        ComparisonUnitArray2 = remaining2.Take(idx2).ToArray()
                    };
                    newListOfCorrelatedSequence.Add(unknownCorrelatedSequenceRemaining);

                    var unknownCorrelatedSequenceAfter = new CorrelatedSequence
                    {
                        CorrelationStatus = CorrelationStatus.Unknown,
                        ComparisonUnitArray1 = remaining1.Skip(idx1).ToArray(),
                        ComparisonUnitArray2 = remaining2.Skip(idx2).ToArray()
                    };
                    newListOfCorrelatedSequence.Add(unknownCorrelatedSequenceAfter);

                    return newListOfCorrelatedSequence;
                }
            }

            var unknownCorrelatedSequence20 = new CorrelatedSequence
            {
                CorrelationStatus = CorrelationStatus.Unknown,
                ComparisonUnitArray1 = remaining1,
                ComparisonUnitArray2 = remaining2
            };
            newListOfCorrelatedSequence.Add(unknownCorrelatedSequence20);

            return newListOfCorrelatedSequence;
        }

        private static List<CorrelatedSequence> DoLcsAlgorithmForTable(CorrelatedSequence unknown)
        {
            var newListOfCorrelatedSequence = new List<CorrelatedSequence>();

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////
            // if we have a table with the same number of rows, and all rows have equal CorrelatedSHA1Hash, then we can
            // flatten and compare every corresponding row.
            // This is true regardless of whether there are horizontally or vertically merged cells, since that
            // characteristic is incorporated into the CorrespondingSHA1Hash. This is probably not very common, but it
            // will never do any harm.
            var tblGroup1 = (ComparisonUnitGroup) unknown.ComparisonUnitArray1.First();
            var tblGroup2 = (ComparisonUnitGroup) unknown.ComparisonUnitArray2.First();

            if (tblGroup1.Contents.Count == tblGroup2.Contents.Count) // if there are the same number of rows
            {
                var zipped = tblGroup1
                    .Contents
                    .Zip(
                        tblGroup2.Contents,
                        (r1, r2) => new
                        {
                            Row1 = r1 as ComparisonUnitGroup,
                            Row2 = r2 as ComparisonUnitGroup
                        })
                    .ToList();

                bool canCollapse = zipped.All(z => z.Row1.CorrelatedSHA1Hash == z.Row2.CorrelatedSHA1Hash);

                if (canCollapse)
                {
                    newListOfCorrelatedSequence = zipped
                        .Select(z =>
                        {
                            var unknownCorrelatedSequence = new CorrelatedSequence
                            {
                                ComparisonUnitArray1 = new ComparisonUnit[] { z.Row1 },
                                ComparisonUnitArray2 = new ComparisonUnit[] { z.Row2 },
                                CorrelationStatus = CorrelationStatus.Unknown
                            };
                            return unknownCorrelatedSequence;
                        })
                        .ToList();
                    return newListOfCorrelatedSequence;
                }
            }

            ComparisonUnitAtom firstContentAtom1 = tblGroup1.DescendantContentAtoms().FirstOrDefault();
            if (firstContentAtom1 == null)
            {
                throw new OpenXmlPowerToolsException("Internal error");
            }

            XElement tblElement1 = firstContentAtom1
                .AncestorElements
                .Reverse()
                .First(a => a.Name == W.tbl);

            ComparisonUnitAtom firstContentAtom2 = tblGroup2.DescendantContentAtoms().FirstOrDefault();
            if (firstContentAtom2 == null)
            {
                throw new OpenXmlPowerToolsException("Internal error");
            }

            XElement tblElement2 = firstContentAtom2
                .AncestorElements
                .Reverse()
                .First(a => a.Name == W.tbl);

            bool leftContainsMerged = tblElement1
                .Descendants()
                .Any(d => d.Name == W.vMerge || d.Name == W.gridSpan);

            bool rightContainsMerged = tblElement2
                .Descendants()
                .Any(d => d.Name == W.vMerge || d.Name == W.gridSpan);

            if (leftContainsMerged || rightContainsMerged)
            {
                // If StructureSha1Hash is the same for both tables, then we know that the structure of the tables is
                // identical, so we can break into correlated sequences for rows.
                if (tblGroup1.StructureSHA1Hash != null &&
                    tblGroup2.StructureSHA1Hash != null &&
                    tblGroup1.StructureSHA1Hash == tblGroup2.StructureSHA1Hash)
                {
                    var zipped = tblGroup1.Contents.Zip(tblGroup2.Contents, (r1, r2) => new
                    {
                        Row1 = r1 as ComparisonUnitGroup,
                        Row2 = r2 as ComparisonUnitGroup
                    });
                    newListOfCorrelatedSequence = zipped
                        .Select(z =>
                        {
                            var unknownCorrelatedSequence = new CorrelatedSequence
                            {
                                ComparisonUnitArray1 = new ComparisonUnit[] { z.Row1 },
                                ComparisonUnitArray2 = new ComparisonUnit[] { z.Row2 },
                                CorrelationStatus = CorrelationStatus.Unknown
                            };
                            return unknownCorrelatedSequence;
                        })
                        .ToList();
                    return newListOfCorrelatedSequence;
                }

                // otherwise flatten to rows
                var deletedCorrelatedSequence = new CorrelatedSequence
                {
                    ComparisonUnitArray1 = unknown
                        .ComparisonUnitArray1
                        .Select(z => z.Contents)
                        .SelectMany(m => m)
                        .ToArray(),
                    ComparisonUnitArray2 = null,
                    CorrelationStatus = CorrelationStatus.Deleted
                };

                newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);

                var insertedCorrelatedSequence = new CorrelatedSequence
                {
                    ComparisonUnitArray1 = null,
                    ComparisonUnitArray2 = unknown
                        .ComparisonUnitArray2
                        .Select(z => z.Contents)
                        .SelectMany(m => m)
                        .ToArray(),
                    CorrelationStatus = CorrelationStatus.Inserted
                };

                newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);

                return newListOfCorrelatedSequence;
            }

            return null;
        }

        private static int FindIndexOfNextParaMark(ComparisonUnit[] cul)
        {
            for (var i = 0; i < cul.Length; i++)
            {
                var cuw = (ComparisonUnitWord) cul[i];
                ComparisonUnitAtom lastAtom = cuw.DescendantContentAtoms().LastOrDefault();
                if (lastAtom?.ContentElement.Name == W.pPr)
                {
                    return i;
                }
            }

            return cul.Length;
        }
    }
}
