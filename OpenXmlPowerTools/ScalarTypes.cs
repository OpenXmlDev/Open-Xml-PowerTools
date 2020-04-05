// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections;
using System.Collections.ObjectModel;

namespace OpenXmlPowerTools
{
    internal static class DefaultScalarTypes
    {
        private static readonly Hashtable defaultScalarTypesHash = new Hashtable(StringComparer.OrdinalIgnoreCase)
        {
            { "System.String", null },
                { "System.SByte", null },
                { "System.Byte", null },
                { "System.Int16", null },
                { "System.UInt16", null },
                { "System.Int32", 10 },
                { "System.UInt32", 10 },
                { "System.Int64", null },
                { "System.UInt64", null },
                { "System.Char", 1 },
                { "System.Single", null },
                { "System.Double", null },
                { "System.Boolean", 5 },
                { "System.Decimal", null },
                { "System.IntPtr", null },
                { "System.Security.SecureString", null }
        };

        internal static bool IsTypeInList(Collection<string> typeNames)
        {
            var text = PSObjectIsOfExactType(typeNames);
            return !string.IsNullOrEmpty(text) && (PSObjectIsEnum(typeNames) || defaultScalarTypesHash.ContainsKey(text));
        }

        internal static string PSObjectIsOfExactType(Collection<string> typeNames)
        {
            if (typeNames.Count != 0)
            {
                return typeNames[0];
            }
            return null;
        }

        internal static bool PSObjectIsEnum(Collection<string> typeNames)
        {
            return typeNames.Count >= 2 && !string.IsNullOrEmpty(typeNames[1]) && string.Equals(typeNames[1], "System.Enum", StringComparison.Ordinal);
        }
    }
}