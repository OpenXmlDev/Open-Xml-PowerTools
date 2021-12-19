using JetBrains.Annotations;

namespace OpenXmlPowerTools
{
    [PublicAPI]
    public sealed class MetricsGetterSettings
    {
        public bool IncludeTextInContentControls { get; set; }

        public bool IncludeXlsxTableCellData { get; set; }

        public bool RetrieveNamespaceList { get; set; }

        public bool RetrieveContentTypeList { get; set; }
    }
}
