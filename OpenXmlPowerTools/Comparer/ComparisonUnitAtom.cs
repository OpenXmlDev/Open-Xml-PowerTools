// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    public class ComparisonUnitAtom : ComparisonUnit
    {
        public ComparisonUnitAtom(
            XElement contentElement,
            XElement[] ancestorElements,
            OpenXmlPart part,
            WmlComparerSettings settings)
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

            var sha1Hash = (string) contentElement.Attribute(PtOpenXml.SHA1Hash);
            if (sha1Hash != null)
            {
                SHA1Hash = sha1Hash;
            }
            else
            {
                string shaHashString = GetSha1HashStringForElement(ContentElement, settings);
                SHA1Hash = WmlComparerUtil.SHA1HashStringForUTF8String(shaHashString);
            }
        }

        // AncestorElements are kept in order from the body to the leaf, because this is the order in which we need to access in order
        // to reassemble the document.  However, in many places in the code, it is necessary to find the nearest ancestor, i.e. cell
        // so it is necessary to reverse the order when looking for it, i.e. look from the leaf back to the body element.

        public XElement[] AncestorElements { get; }

        public XElement ContentElement { get; }

        public XElement RevTrackElement { get; }

        public string[] AncestorUnids { get; set; }

        public ComparisonUnitAtom ComparisonUnitAtomBefore { get; set; }

        public XElement ContentElementBefore { get; set; }

        public OpenXmlPart Part { get; }

        private static string GetSha1HashStringForElement(XElement contentElement, WmlComparerSettings settings)
        {
            string text = contentElement.Value;
            if (settings.CaseInsensitive)
            {
                text = text.ToUpper(settings.CultureInfo);
            }

            return contentElement.Name.LocalName + text;
        }

        private static XElement GetRevisionTrackingElementFromAncestors(
            XElement contentElement,
            IEnumerable<XElement> ancestors)
        {
            return contentElement.Name == W.pPr
                ? contentElement.Elements(W.rPr).Elements().FirstOrDefault(e => e.Name == W.del || e.Name == W.ins)
                : ancestors.FirstOrDefault(a => a.Name == W.del || a.Name == W.ins);
        }

        public override string ToString()
        {
            return ToString(0);
        }

        public override string ToString(int indent)
        {
            const int xNamePad = 16;
            string indentString = "".PadRight(indent);

            var sb = new StringBuilder();
            sb.Append(indentString);

            var correlationStatus = "";
            if (CorrelationStatus != CorrelationStatus.Nil)
            {
                correlationStatus = $"[{CorrelationStatus.ToString().PadRight(8)}] ";
            }

            if (ContentElement.Name == W.t || ContentElement.Name == W.delText)
            {
                sb.AppendFormat(
                    "Atom {0}: {1} {2} SHA1:{3} ",
                    PadLocalName(xNamePad, this),
                    ContentElement.Value,
                    correlationStatus,
                    SHA1Hash.Substring(0, 8));

                AppendAncestorsDump(sb, this);
            }
            else
            {
                sb.AppendFormat(
                    "Atom {0}:   {1} SHA1:{2} ",
                    PadLocalName(xNamePad, this),
                    correlationStatus,
                    SHA1Hash.Substring(0, 8));

                AppendAncestorsDump(sb, this);
            }

            return sb.ToString();
        }

        public string ToStringAncestorUnids()
        {
            return ToStringAncestorUnids(0);
        }

        private string ToStringAncestorUnids(int indent)
        {
            const int xNamePad = 16;
            string indentString = "".PadRight(indent);

            var sb = new StringBuilder();
            sb.Append(indentString);

            var correlationStatus = "";
            if (CorrelationStatus != CorrelationStatus.Nil)
            {
                correlationStatus = $"[{CorrelationStatus.ToString().PadRight(8)}] ";
            }

            if (ContentElement.Name == W.t || ContentElement.Name == W.delText)
            {
                sb.AppendFormat(
                    "Atom {0}: {1} {2} SHA1:{3} ",
                    PadLocalName(xNamePad, this),
                    ContentElement.Value,
                    correlationStatus,
                    SHA1Hash.Substring(0, 8));

                AppendAncestorsUnidsDump(sb, this);
            }
            else
            {
                sb.AppendFormat(
                    "Atom {0}:   {1} SHA1:{2} ",
                    PadLocalName(xNamePad, this),
                    correlationStatus,
                    SHA1Hash.Substring(0, 8));

                AppendAncestorsUnidsDump(sb, this);
            }

            return sb.ToString();
        }

        private static string PadLocalName(int xNamePad, ComparisonUnitAtom item)
        {
            return (item.ContentElement.Name.LocalName + " ").PadRight(xNamePad, '-') + " ";
        }

        private static void AppendAncestorsDump(StringBuilder sb, ComparisonUnitAtom sr)
        {
            string s = sr
                .AncestorElements.Select(p => p.Name.LocalName + GetUnid(p) + "/")
                .StringConcatenate()
                .TrimEnd('/');

            sb.Append("Ancestors:" + s);
        }

        private static void AppendAncestorsUnidsDump(StringBuilder sb, ComparisonUnitAtom sr)
        {
            var zipped = sr.AncestorElements.Zip(sr.AncestorUnids, (a, u) => new
            {
                AncestorElement = a,
                AncestorUnid = u
            });

            string s = zipped
                .Select(p => p.AncestorElement.Name.LocalName + "[" + p.AncestorUnid.Substring(0, 8) + "]/")
                .StringConcatenate().TrimEnd('/');

            sb.Append("Ancestors:" + s);
        }

        private static string GetUnid(XElement p)
        {
            var unid = (string) p.Attribute(PtOpenXml.Unid);
            return unid == null ? "" : "[" + unid.Substring(0, 8) + "]";
        }
    }
}
