/***************************************************************************

Copyright (c) Microsoft Corporation 2012-2015.

This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license can be found here:

http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx

Published at http://OpenXmlDeveloper.org
Resource Center and Documentation: http://openxmldeveloper.org/wiki/w/wiki/powertools-for-open-xml.aspx

Developer: Eric White
Blog: http://www.ericwhite.com
Twitter: @EricWhiteDev
Email: eric@ericwhite.com

***************************************************************************/

using System;
using System.Collections;
using System.Collections.ObjectModel;

namespace OpenXmlPowerTools
{
    internal static class DefaultScalarTypes
    {
        private static readonly Hashtable defaultScalarTypesHash;
        internal static bool IsTypeInList(Collection<string> typeNames)
        {
            string text = PSObjectIsOfExactType(typeNames);
            return !string.IsNullOrEmpty(text) && (PSObjectIsEnum(typeNames) || DefaultScalarTypes.defaultScalarTypesHash.ContainsKey(text));
        }

        static DefaultScalarTypes()
        {
            DefaultScalarTypes.defaultScalarTypesHash = new Hashtable(StringComparer.OrdinalIgnoreCase);
            DefaultScalarTypes.defaultScalarTypesHash.Add("System.String", null);
            DefaultScalarTypes.defaultScalarTypesHash.Add("System.SByte", null);
            DefaultScalarTypes.defaultScalarTypesHash.Add("System.Byte", null);
            DefaultScalarTypes.defaultScalarTypesHash.Add("System.Int16", null);
            DefaultScalarTypes.defaultScalarTypesHash.Add("System.UInt16", null);
            DefaultScalarTypes.defaultScalarTypesHash.Add("System.Int32", 10);
            DefaultScalarTypes.defaultScalarTypesHash.Add("System.UInt32", 10);
            DefaultScalarTypes.defaultScalarTypesHash.Add("System.Int64", null);
            DefaultScalarTypes.defaultScalarTypesHash.Add("System.UInt64", null);
            DefaultScalarTypes.defaultScalarTypesHash.Add("System.Char", 1);
            DefaultScalarTypes.defaultScalarTypesHash.Add("System.Single", null);
            DefaultScalarTypes.defaultScalarTypesHash.Add("System.Double", null);
            DefaultScalarTypes.defaultScalarTypesHash.Add("System.Boolean", 5);
            DefaultScalarTypes.defaultScalarTypesHash.Add("System.Decimal", null);
            DefaultScalarTypes.defaultScalarTypesHash.Add("System.IntPtr", null);
            DefaultScalarTypes.defaultScalarTypesHash.Add("System.Security.SecureString", null);
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
