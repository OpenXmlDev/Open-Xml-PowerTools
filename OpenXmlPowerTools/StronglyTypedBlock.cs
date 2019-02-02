// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    /// <summary>
    /// Provides an elegant way of wrapping a set of invocations of the strongly typed 
    /// classes provided by the Open XML SDK) in a using statement that demarcates those
    /// invokations as one "block" before and after which the PowerTools can be used safely.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This class lends itself to scenarios where the PowerTools and Linq-to-XML are used as
    /// the primary API for working with Open XML elements, next to the strongly typed classes
    /// provided by the Open XML SDK. In these scenarios, the class would be used as follows:
    /// </para>
    /// <code>
    ///     [Your code using the PowerTools]
    /// 
    ///     using (new NonPowerToolsBlock(wordprocessingDocument))
    ///     {
    ///         [Your code using the strongly typed classes]
    ///     }
    /// 
    ///    [Your code using the PowerTools]
    /// </code>
    /// <para>
    /// Upon creation, instances of this class will invoke the
    /// <see cref="PowerToolsBlockExtensions.EndPowerToolsBlock"/> method on the package
    /// to begin the block. Upon disposal, instances of this class will call the
    /// <see cref="PowerToolsBlockExtensions.BeginPowerToolsBlock"/> method on the package
    /// to end the block.
    /// </para>
    /// </remarks>
    /// <seealso cref="PowerToolsBlock"/>
    /// <seealso cref="PowerToolsBlockExtensions.BeginPowerToolsBlock"/>
    /// <seealso cref="PowerToolsBlockExtensions.EndPowerToolsBlock"/>
    public class StronglyTypedBlock : IDisposable
    {
        private OpenXmlPackage _package;

        public StronglyTypedBlock(OpenXmlPackage package)
        {
            if (package == null) throw new ArgumentNullException("package");

            _package = package;
            _package.EndPowerToolsBlock();
        }

        public void Dispose()
        {
            if (_package == null) return;

            _package.BeginPowerToolsBlock();
            _package = null;
        }
    }
}
