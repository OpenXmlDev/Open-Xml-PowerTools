using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    /// <summary>
    /// This class is used to prevent duplication of media.
    /// </summary>
    internal class MediaData
    {
        public List<ContentPartRelTypeIdTuple> ContentPartRelTypeIdList = new List<ContentPartRelTypeIdTuple>();

        public MediaData(DataPart part)
        {
            ContentType = part.ContentType;
            using (Stream s = part.GetStream(FileMode.Open, FileAccess.Read))
            {
                Media = new byte[s.Length];
                s.Read(Media, 0, (int) s.Length);
            }
        }

        public DataPart DataPart { get; set; }
        private string ContentType { get; }
        private byte[] Media { get; }

        public void AddContentPartRelTypeResourceIdTupple(OpenXmlPart contentPart, string relationshipType, string relationshipId)
        {
            ContentPartRelTypeIdList.Add(
                new ContentPartRelTypeIdTuple
                {
                    ContentPart = contentPart,
                    RelationshipType = relationshipType,
                    RelationshipId = relationshipId
                });
        }

        public void WriteMedia(DataPart part)
        {
            using (Stream s = part.GetStream(FileMode.Create, FileAccess.ReadWrite))
                s.Write(Media, 0, Media.GetUpperBound(0) + 1);
        }

        public bool Compare(MediaData arg)
        {
            if (ContentType != arg.ContentType)
                return false;
            if (Media.GetLongLength(0) != arg.Media.GetLongLength(0))
                return false;

            // Compare the arrays byte by byte
            long length = Media.GetLongLength(0);
            byte[] media1 = Media;
            byte[] media2 = arg.Media;
            for (long n = 0; n < length; n++)
                if (media1[n] != media2[n])
                    return false;

            return true;
        }
    }
}