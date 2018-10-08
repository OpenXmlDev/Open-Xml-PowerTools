using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    internal class ContentPartRelTypeIdTuple
    {
        public OpenXmlPart ContentPart { get; set; }
        public string RelationshipType { get; set; }
        public string RelationshipId { get; set; }
    }
}
