// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Drawing;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static partial class WmlComparer
    {
        private class Atgbw
        {
            public int? Key;
            public ComparisonUnitAtom ComparisonUnitAtomMember;
            public int NextIndex;
        }

        private class ConsolidationInfo
        {
            public string Revisor;
            public Color Color;
            public XElement RevisionElement;
            public bool InsertBefore;
            public string RevisionHash;
            public XElement[] Footnotes;
            public XElement[] Endnotes;
            public string RevisionString; // for debugging purposes only
        }

        private class RecursionInfo
        {
            public XName ElementName;
            public XName[] ChildElementPropertyNames;
        }
    }
}
