

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

// It is possible to optimize DescendantContentAtoms

/// Currently, the unid is set at the beginning of the algorithm.  It is used by the code that establishes correlation based on first rejecting
/// tracked revisions, then correlating paragraphs/tables.  It is requred for this algorithm - after finding a correlated sequence in the document with rejected
/// revisions, it uses the unid to find the same paragraph in the document without rejected revisions, then sets the correlated sha1 hash in that document.
///
/// But then when accepting tracked revisions, for certain paragraphs (where there are deleted paragraph marks) it is going to lose the unids.  But this isn't a
/// problem because when paragraph marks are deleted, the correlation is definitely no longer possible.  Any paragraphs that are in a range of paragraphs that
/// are coalesced can't be correlated to paragraphs in the other document via their hash.  At that point we no longer care what their unids are.
///
/// But after that it is only used to reconstruct the tree.  It is also used in the debugging code that
/// prints the various correlated sequences and comparison units - this is display for debugging purposes only.

/// The key idea here is that a given paragraph will always have the same ancestors, and it doesn't matter whether the content was deleted from the old document,
/// inserted into the new document, or set as equal.  At this point, we identify a paragraph as a sequential list of content atoms, terminated by a paragraph mark.
/// This entire list will for a single paragraph, regardless of whether the paragraph is a child of the body, or if the paragraph is in a cell in a table, or if
/// the paragraph is in a text box.  The list of ancestors, from the paragraph to the root of the XML tree will be the same for all content atoms in the paragraph.
///
/// Therefore:
///
/// Iterate through the list of content atoms backwards.  When the loop sees a paragraph mark, it gets the ancestor unids from the paragraph mark to the top of the
/// tree, and sets this as the same for all content atoms in the paragraph.  For descendants of the paragraph mark, it doesn't really matter if content is put into
/// separate runs or what not.  We don't need to be concerned about what the unids are for descendants of the paragraph.

namespace OpenXmlPowerTools
{
    internal class ComparisonUnitGroup : ComparisonUnit
    {
        public ComparisonUnitGroupType ComparisonUnitGroupType;
        public string CorrelatedSHA1Hash;
        public string StructureSHA1Hash;

        public ComparisonUnitGroup(IEnumerable<ComparisonUnit> comparisonUnitList, ComparisonUnitGroupType groupType, int level)
        {
            Contents = comparisonUnitList.ToList();
            ComparisonUnitGroupType = groupType;
            var first = comparisonUnitList.First();
            var comparisonUnitAtom = GetFirstComparisonUnitAtomOfGroup(first);
            XName ancestorName = null;
            if (groupType == OpenXmlPowerTools.ComparisonUnitGroupType.Table)
            {
                ancestorName = W.tbl;
            }
            else if (groupType == OpenXmlPowerTools.ComparisonUnitGroupType.Row)
            {
                ancestorName = W.tr;
            }
            else if (groupType == OpenXmlPowerTools.ComparisonUnitGroupType.Cell)
            {
                ancestorName = W.tc;
            }
            else if (groupType == OpenXmlPowerTools.ComparisonUnitGroupType.Paragraph)
            {
                ancestorName = W.p;
            }
            else if (groupType == OpenXmlPowerTools.ComparisonUnitGroupType.Textbox)
            {
                ancestorName = W.txbxContent;
            }

            var ancestorsToLookAt = comparisonUnitAtom.AncestorElements.Where(ae => ae.Name == W.tbl || ae.Name == W.tr || ae.Name == W.tc || ae.Name == W.p || ae.Name == W.txbxContent).ToArray(); ;
            var ancestor = ancestorsToLookAt[level];

            if (ancestor == null)
            {
                throw new OpenXmlPowerToolsException("Internal error: ComparisonUnitGroup");
            }

            SHA1Hash = (string)ancestor.Attribute(PtOpenXml.SHA1Hash);
            CorrelatedSHA1Hash = (string)ancestor.Attribute(PtOpenXml.CorrelatedSHA1Hash);
            StructureSHA1Hash = (string)ancestor.Attribute(PtOpenXml.StructureSHA1Hash);
        }

        public static ComparisonUnitAtom GetFirstComparisonUnitAtomOfGroup(ComparisonUnit group)
        {
            var thisGroup = group;
            while (true)
            {
                var tg = thisGroup as ComparisonUnitGroup;
                if (tg != null)
                {
                    thisGroup = tg.Contents.First();
                    continue;
                }
                var tw = thisGroup as ComparisonUnitWord;
                if (tw == null)
                {
                    throw new OpenXmlPowerToolsException("Internal error: GetFirstComparisonUnitAtomOfGroup");
                }

                var ca = (ComparisonUnitAtom)tw.Contents.First();
                return ca;
            }
        }

        public override string ToString(int indent)
        {
            var sb = new StringBuilder();
            sb.Append("".PadRight(indent) + "Group Type: " + ComparisonUnitGroupType.ToString() + " SHA1:" + SHA1Hash + Environment.NewLine);
            foreach (var comparisonUnitAtom in Contents)
            {
                sb.Append(comparisonUnitAtom.ToString(indent + 2));
            }

            return sb.ToString();
        }
    }
}