// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static partial class WmlComparer
    {
        public class WmlComparerRevision
        {
            public WmlComparerRevisionType RevisionType { get; set; }
            public string Text { get; set; }
            public string Author { get; set; }
            public string Date { get; set; }
            public XElement ContentXElement { get; set; }
            public XElement RevisionXElement { get; set; }
            public Uri PartUri { get; set; }
            public string PartContentType { get; set; }
        }

        public enum WmlComparerRevisionType
        {
            Inserted,
            Deleted
        }
    }
}