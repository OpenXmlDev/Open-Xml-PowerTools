// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Security.Cryptography;
using System.Text;

namespace OpenXmlPowerTools
{
    internal static class WmlComparerUtil
    {
        public static string SHA1HashStringForUTF8String(string s)
        {
            var bytes = Encoding.UTF8.GetBytes(s);
            var sha1 = SHA1.Create();
            var hashBytes = sha1.ComputeHash(bytes);
            return HexStringFromBytes(hashBytes);
        }

        public static string SHA1HashStringForByteArray(byte[] bytes)
        {
            var sha1 = SHA1.Create();
            var hashBytes = sha1.ComputeHash(bytes);
            return HexStringFromBytes(hashBytes);
        }

        public static string HexStringFromBytes(byte[] bytes)
        {
            var sb = new StringBuilder();
            foreach (var b in bytes)
            {
                var hex = b.ToString("x2");
                sb.Append(hex);
            }

            return sb.ToString();
        }

        public static ComparisonUnitGroupType ComparisonUnitGroupTypeFromLocalName(string localName)
        {
            switch (localName)
            {
                case "p":
                    return ComparisonUnitGroupType.Paragraph;
                case "tbl":
                    return ComparisonUnitGroupType.Table;
                case "tr":
                    return ComparisonUnitGroupType.Row;
                case "tc":
                    return ComparisonUnitGroupType.Cell;
                case "txbxContent":
                    return ComparisonUnitGroupType.Textbox;
                default:
                    throw new ArgumentOutOfRangeException(nameof(localName),
                        $@"Unsupported localName: '{localName}'.");
            }
        }
    }
}
