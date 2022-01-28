using System.Collections.Generic;

namespace Codeuctivity
{
    public partial class WmlDocument : OpenXmlPowerToolsDocument
    {
        public IEnumerable<WmlDocument> SplitOnSections() => Codeuctivity.DocumentBuilder.DocumentBuilder.SplitOnSections(this);
    }
}