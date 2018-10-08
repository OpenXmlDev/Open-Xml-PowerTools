using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    /// <summary>
    /// This class is used to prevent duplication of images
    /// </summary>
    internal class ImageData
    {
        public List<ContentPartRelTypeIdTuple> ContentPartRelTypeIdList = new List<ContentPartRelTypeIdTuple>();

        public ImageData(ImagePart part)
        {
            ContentType = part.ContentType;
            using (Stream s = part.GetStream(FileMode.Open, FileAccess.Read))
            {
                Image = new byte[s.Length];
                s.Read(Image, 0, (int) s.Length);
            }
        }

        public OpenXmlPart ImagePart { get; set; }
        private string ContentType { get; }
        private byte[] Image { get; }

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

        public void WriteImage(ImagePart part)
        {
            using (Stream s = part.GetStream(FileMode.Create, FileAccess.ReadWrite))
                s.Write(Image, 0, Image.GetUpperBound(0) + 1);
        }

        public bool Compare(ImageData arg)
        {
            if (ContentType != arg.ContentType)
                return false;
            if (Image.GetLongLength(0) != arg.Image.GetLongLength(0))
                return false;

            // Compare the arrays byte by byte
            long length = Image.GetLongLength(0);
            byte[] image1 = Image;
            byte[] image2 = arg.Image;
            for (long n = 0; n < length; n++)
                if (image1[n] != image2[n])
                    return false;

            return true;
        }
    }
}