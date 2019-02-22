// Copyright (c) Lowell Stewart. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public interface IMetadataParser
    {
        string DelimiterOpen { get; }
        string DelimiterClose { get; }
        string EmbedOpen { get; }
        string EmbedClose { get; }
        XElement TransformContentToMetadata(string content);
    }

    public class MetadataParseException : Exception
    {
        public MetadataParseException() { }
        public MetadataParseException(string message) : base(message) { }
        public MetadataParseException(string message, Exception inner) : base(message, inner) { }
    }

}