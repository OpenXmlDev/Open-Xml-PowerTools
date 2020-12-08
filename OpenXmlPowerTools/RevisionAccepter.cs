using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    public class RevisionAccepter
    {
        public static WmlDocument AcceptRevisions(WmlDocument document)
        {
            using var streamDoc = new OpenXmlMemoryStreamDocument(document);
            using (var doc = streamDoc.GetWordprocessingDocument())
            {
                AcceptRevisions(doc);
            }
            return streamDoc.GetModifiedWmlDocument();
        }

        public static void AcceptRevisions(WordprocessingDocument doc)
        {
            RevisionProcessor.AcceptRevisions(doc);
        }

        public static bool PartHasTrackedRevisions(OpenXmlPart part)
        {
            return RevisionProcessor.PartHasTrackedRevisions(part);
        }

        public static bool HasTrackedRevisions(WmlDocument document)
        {
            return RevisionProcessor.HasTrackedRevisions(document);
        }

        public static bool HasTrackedRevisions(WordprocessingDocument doc)
        {
            return RevisionProcessor.HasTrackedRevisions(doc);
        }
    }
}