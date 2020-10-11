using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;
using System.IO;

namespace OpenXmlPowerTools
{
    // This class is used to prevent duplication of media
    internal class MediaData
    {
        private string ContentType { get; set; }
        private byte[] Media { get; set; }
        public DataPart DataPart { get; set; }
        public List<ContentPartRelTypeIdTuple> ContentPartRelTypeIdList = new List<ContentPartRelTypeIdTuple>();

        public MediaData(DataPart part)
        {
            ContentType = part.ContentType;
            using var s = part.GetStream(FileMode.Open, FileAccess.Read);
            Media = new byte[s.Length];
            s.Read(Media, 0, (int)s.Length);
        }

        public void AddContentPartRelTypeResourceIdTupple(OpenXmlPart contentPart, string relationshipType, string relationshipId)
        {
            ContentPartRelTypeIdList.Add(
                new ContentPartRelTypeIdTuple()
                {
                    ContentPart = contentPart,
                    RelationshipType = relationshipType,
                    RelationshipId = relationshipId,
                });
        }

        public void WriteMedia(DataPart part)
        {
            using var s = part.GetStream(FileMode.Create, FileAccess.ReadWrite);
            s.Write(Media, 0, Media.GetUpperBound(0) + 1);
        }

        public bool Compare(MediaData arg)
        {
            if (ContentType != arg.ContentType)
            {
                return false;
            }

            if (Media.GetLongLength(0) != arg.Media.GetLongLength(0))
            {
                return false;
            }
            // Compare the arrays byte by byte
            var length = Media.GetLongLength(0);
            var media1 = Media;
            var media2 = arg.Media;
            for (long n = 0; n < length; n++)
            {
                if (media1[n] != media2[n])
                {
                    return false;
                }
            }

            return true;
        }
    }
}