// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    /// <summary>
    /// Provides an elegant way of wrapping a set of invocations of the PowerTools in a using
    /// statement that demarcates those invokations as one "block" before and after which the
    /// strongly typed classes provided by the Open XML SDK can be used safely.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This class lends itself to scenarios where the PowerTools and Linq-to-XML are used as
    /// a secondary API for working with Open XML elements, next to the strongly typed classes
    /// provided by the Open XML SDK. In these scenarios, the class would be
    /// used as follows:
    /// </para>
    /// <code>
    ///     [Your code using the strongly typed classes]
    /// 
    ///     using (new PowerToolsBlock(wordprocessingDocument))
    ///     {
    ///         [Your code using the PowerTools]
    ///     }
    /// 
    ///    [Your code using the strongly typed classes]
    /// </code>
    /// <para>
    /// Upon creation, instances of this class will invoke the
    /// <see cref="PowerToolsBlockExtensions.BeginPowerToolsBlock"/> method on the package
    /// to begin the transaction.  Upon disposal, instances of this class will call the
    /// <see cref="PowerToolsBlockExtensions.EndPowerToolsBlock"/> method on the package
    /// to end the transaction.
    /// </para>
    /// </remarks>
    /// <seealso cref="StronglyTypedBlock" />
    /// <seealso cref="PowerToolsBlockExtensions.BeginPowerToolsBlock"/>
    /// <seealso cref="PowerToolsBlockExtensions.EndPowerToolsBlock"/>
    public class PowerToolsBlock : IDisposable
    {
        private OpenXmlPackage _package;

        public PowerToolsBlock(OpenXmlPackage package)
        {
            if (package == null) throw new ArgumentNullException("package");

            _package = package;
            _package.BeginPowerToolsBlock();
        }

        public void Dispose()
        {
            if (_package == null) return;

            _package.EndPowerToolsBlock();
            _package = null;
        }
    }
}
