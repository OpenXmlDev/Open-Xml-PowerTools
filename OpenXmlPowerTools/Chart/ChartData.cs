// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace OpenXmlPowerTools
{
    public class ChartData
    {
        public string[] SeriesNames { get; set; }

        public ChartDataType CategoryDataType { get; set; }
        public int CategoryFormatCode { get; set; }
        public string[] CategoryNames { get; set; }

        public double[][] Values { get; set; }
    }
}