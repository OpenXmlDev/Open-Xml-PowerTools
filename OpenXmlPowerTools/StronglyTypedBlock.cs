// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Packaging;
using JetBrains.Annotations;

#pragma warning disable IDE0060 // Remove unused parameter

namespace OpenXmlPowerTools
{
    /// <summary>
    /// Does nothing.
    /// </summary>
    [Obsolete("No longer required with DocumentFormat.OpenXml.Linq")]
    [PublicAPI]
    public class StronglyTypedBlock : IDisposable
    {
        public StronglyTypedBlock(OpenXmlPackage package)
        {
        }

        public void Dispose()
        {
        }
    }
}
