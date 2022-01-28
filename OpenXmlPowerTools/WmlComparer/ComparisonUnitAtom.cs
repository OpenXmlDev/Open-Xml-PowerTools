using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

// It is possible to optimize DescendantContentAtoms

// Currently, the unid is set at the beginning of the algorithm.  It is used by the code that establishes correlation based on first rejecting
// tracked revisions, then correlating paragraphs/tables.  It is requred for this algorithm - after finding a correlated sequence in the document with rejected
// revisions, it uses the unid to find the same paragraph in the document without rejected revisions, then sets the correlated sha1 hash in that document.
//
// But then when accepting tracked revisions, for certain paragraphs (where there are deleted paragraph marks) it is going to lose the unids.  But this isn't a
// problem because when paragraph marks are deleted, the correlation is definitely no longer possible.  Any paragraphs that are in a range of paragraphs that
// are coalesced can't be correlated to paragraphs in the other document via their hash.  At that point we no longer care what their unids are.
//
// But after that it is only used to reconstruct the tree.  It is also used in the debugging code that
// prints the various correlated sequences and comparison units - this is display for debugging purposes only.

// The key idea here is that a given paragraph will always have the same ancestors, and it doesn't matter whether the content was deleted from the old document,
// inserted into the new document, or set as equal.  At this point, we identify a paragraph as a sequential list of content atoms, terminated by a paragraph mark.
// This entire list will for a single paragraph, regardless of whether the paragraph is a child of the body, or if the paragraph is in a cell in a table, or if
// the paragraph is in a text box.  The list of ancestors, from the paragraph to the root of the XML tree will be the same for all content atoms in the paragraph.
//
// Therefore:
//
// Iterate through the list of content atoms backwards.  When the loop sees a paragraph mark, it gets the ancestor unids from the paragraph mark to the top of the
// tree, and sets this as the same for all content atoms in the paragraph.  For descendants of the paragraph mark, it doesn't really matter if content is put into
// separate runs or what not.  We don't need to be concerned about what the unids are for descendants of the paragraph.

namespace Codeuctivity.WmlComparer
{
    public class ComparisonUnitAtom : ComparisonUnit
    {
        // AncestorElements are kept in order from the body to the leaf, because this is the order in which we need to access in order to reassemble the document.  However, in many places in the code, it is necessary to find the nearest ancestor, i.e. cell so it is necessary to reverse the order when looking for it, i.e. look from the leaf back to the body element.

        public XElement[] AncestorElements { get; set; }
        public string[] AncestorUnids { get; set; }
        public XElement ContentElement { get; set; }
        public XElement ContentElementBefore { get; set; }
        public ComparisonUnitAtom ComparisonUnitAtomBefore { get; set; }
        public OpenXmlPart Part { get; set; }
        public XElement RevTrackElement { get; set; }

        public ComparisonUnitAtom(XElement contentElement, XElement[] ancestorElements, OpenXmlPart part, WmlComparerSettings settings)
        {
            ContentElement = contentElement;
            AncestorElements = ancestorElements;
            Part = part;
            RevTrackElement = GetRevisionTrackingElementFromAncestors(contentElement, AncestorElements);
            if (RevTrackElement == null)
            {
                CorrelationStatus = CorrelationStatus.Equal;
            }
            else
            {
                if (RevTrackElement.Name == W.del)
                {
                    CorrelationStatus = CorrelationStatus.Deleted;
                }
                else if (RevTrackElement.Name == W.ins)
                {
                    CorrelationStatus = CorrelationStatus.Inserted;
                }
            }
            var sha1Hash = (string)contentElement.Attribute(PtOpenXml.SHA1Hash);
            if (sha1Hash != null)
            {
                SHA1Hash = sha1Hash;
            }
            else
            {
                var shaHashString = GetSha1HashStringForElement(ContentElement, settings);
                SHA1Hash = PtUtils.SHA1HashStringForUTF8String(shaHashString);
            }
        }

        private string GetSha1HashStringForElement(XElement contentElement, WmlComparerSettings settings)
        {
            var text = contentElement.Value;
            if (settings.CaseInsensitive)
            {
                text = text.ToUpper(settings.CultureInfo);
            }

            if (settings.ConflateBreakingAndNonbreakingSpaces)
            {
                text = text.Replace(' ', '\x00a0');
            }

            return contentElement.Name.LocalName + text;
        }

        private static XElement GetRevisionTrackingElementFromAncestors(XElement contentElement, XElement[] ancestors)
        {
            XElement revTrackElement = null;

            if (contentElement.Name == W.pPr)
            {
                revTrackElement = contentElement
                    .Elements(W.rPr)
                    .Elements()
                    .FirstOrDefault(e => e.Name == W.del || e.Name == W.ins);
                return revTrackElement;
            }

            revTrackElement = ancestors.FirstOrDefault(a => a.Name == W.del || a.Name == W.ins);
            return revTrackElement;
        }

        public override string ToString(int indent)
        {
            var xNamePad = 16;
            var indentString = "".PadRight(indent);

            var sb = new StringBuilder();
            sb.Append(indentString);
            var correlationStatus = "";
            if (CorrelationStatus != CorrelationStatus.Nil)
            {
                correlationStatus = string.Format("[{0}] ", CorrelationStatus.ToString().PadRight(8));
            }

            if (ContentElement.Name == W.t || ContentElement.Name == W.delText)
            {
                sb.AppendFormat("Atom {0}: {1} {2} SHA1:{3} ", PadLocalName(xNamePad, this), ContentElement.Value, correlationStatus, SHA1Hash.Substring(0, 8));
                AppendAncestorsDump(sb, this);
            }
            else
            {
                sb.AppendFormat("Atom {0}:   {1} SHA1:{2} ", PadLocalName(xNamePad, this), correlationStatus, SHA1Hash.Substring(0, 8));
                AppendAncestorsDump(sb, this);
            }
            return sb.ToString();
        }

        public string ToStringAncestorUnids(int indent)
        {
            var xNamePad = 16;
            var indentString = "".PadRight(indent);

            var sb = new StringBuilder();
            sb.Append(indentString);
            var correlationStatus = "";
            if (CorrelationStatus != CorrelationStatus.Nil)
            {
                correlationStatus = string.Format("[{0}] ", CorrelationStatus.ToString().PadRight(8));
            }

            if (ContentElement.Name == W.t || ContentElement.Name == W.delText)
            {
                sb.AppendFormat("Atom {0}: {1} {2} SHA1:{3} ", PadLocalName(xNamePad, this), ContentElement.Value, correlationStatus, SHA1Hash.Substring(0, 8));
                AppendAncestorsUnidsDump(sb, this);
            }
            else
            {
                sb.AppendFormat("Atom {0}:   {1} SHA1:{2} ", PadLocalName(xNamePad, this), correlationStatus, SHA1Hash.Substring(0, 8));
                AppendAncestorsUnidsDump(sb, this);
            }
            return sb.ToString();
        }

        public override string ToString()
        {
            return ToString(0);
        }

        public string ToStringAncestorUnids()
        {
            return ToStringAncestorUnids(0);
        }

        private static string PadLocalName(int xNamePad, ComparisonUnitAtom item)
        {
            return (item.ContentElement.Name.LocalName + " ").PadRight(xNamePad, '-') + " ";
        }

        private void AppendAncestorsDump(StringBuilder sb, ComparisonUnitAtom sr)
        {
            var s = sr.AncestorElements.Select(p => p.Name.LocalName + GetUnid(p) + "/").StringConcatenate().TrimEnd('/');
            sb.Append("Ancestors:" + s);
        }

        private void AppendAncestorsUnidsDump(StringBuilder sb, ComparisonUnitAtom sr)
        {
            var zipped = sr.AncestorElements.Zip(sr.AncestorUnids, (a, u) => new
            {
                AncestorElement = a,
                AncestorUnid = u,
            });
            var s = zipped.Select(p => p.AncestorElement.Name.LocalName + "[" + p.AncestorUnid.Substring(0, 8) + "]/").StringConcatenate().TrimEnd('/');
            sb.Append("Ancestors:" + s);
        }

        private string GetUnid(XElement p)
        {
            var unid = (string)p.Attribute(PtOpenXml.Unid);
            if (unid == null)
            {
                return "";
            }

            return "[" + unid.Substring(0, 8) + "]";
        }

        public static string ComparisonUnitAtomListToString(List<ComparisonUnitAtom> comparisonUnitAtomList, int indent)
        {
            var sb = new StringBuilder();
            var cal = comparisonUnitAtomList
                .Select((ca, i) => new
                {
                    ComparisonUnitAtom = ca,
                    Index = i,
                });
            foreach (var item in cal)
            {
                sb.Append("".PadRight(indent))
                  .AppendFormat("[{0:000000}] ", item.Index + 1)
                  .Append(item.ComparisonUnitAtom.ToString(0) + Environment.NewLine);
            }

            return sb.ToString();
        }
    }
}