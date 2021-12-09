// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Packaging;
using JetBrains.Annotations;

namespace OpenXmlPowerTools
{
    [PublicAPI]
    public static class PowerToolsBlockExtensions
    {
        /// <summary>
        /// Does nothing.
        /// </summary>
        [Obsolete("No longer required with DocumentFormat.OpenXml.Linq")]
        public static void BeginPowerToolsBlock(this OpenXmlPackage package)
        {
        }

        /// <summary>
        /// Does nothing.
        /// </summary>
        [Obsolete("No longer required with DocumentFormat.OpenXml.Linq")]
        public static void EndPowerToolsBlock(this OpenXmlPackage package)
        {
        }
    }
}
