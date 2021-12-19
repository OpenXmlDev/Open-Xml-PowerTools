using JetBrains.Annotations;

namespace OpenXmlPowerTools
{
    [PublicAPI]
    public sealed class ChartData
    {
        public string[] SeriesNames;

        public ChartDataType CategoryDataType;

        public int CategoryFormatCode;

        public string[] CategoryNames;

        public double[][] Values;
    }
}
