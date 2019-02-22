// Copyright (c) Lowell Stewart. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Threading.Tasks;

namespace OpenXmlPowerTools
{
    public interface IDataContext : IMetadataParser
    {
        string EvaluateText(string selector, bool optional);
        bool EvaluateBool(string selector, string match, string notMatch);
        IDataContext[] EvaluateList(string selector, bool optional);
        void Release();
    }

    public interface IAsyncDataContext : IMetadataParser
    {
        Task<string> EvaluateTextAsync(string selector, bool optional);
        Task<bool> EvaluateBoolAsync(string selector, string match, string notMatch);
        Task<IAsyncDataContext[]> EvaluateListAsync(string selector, bool optional);
        Task ReleaseAsync();
    }

    public class EvaluationException : Exception
    {
        public EvaluationException() { }
        public EvaluationException(string message) : base(message) { }
        public EvaluationException(string message, Exception inner) : base(message, inner) { }
    }

}