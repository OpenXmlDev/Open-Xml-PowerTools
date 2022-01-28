namespace Codeuctivity.Chart
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