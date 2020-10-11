

#define TestForUnsupportedDocuments
#define MergeStylesWithSameNames

using System.Collections.Generic;

namespace OpenXmlPowerTools
{
    public partial class WmlDocument : OpenXmlPowerToolsDocument
    {
        public IEnumerable<WmlDocument> SplitOnSections()
        {
            return DocumentBuilder.SplitOnSections(this);
        }
    }
}