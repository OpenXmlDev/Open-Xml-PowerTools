#define TestForUnsupportedDocuments
#define MergeStylesWithSameNames

namespace Codeuctivity.DocumentBuilder
{
    public class Source
    {
        public WmlDocument WmlDocument { get; set; }
        public int Start { get; set; }
        public int Count { get; set; }
        public bool KeepSections { get; set; }
        public bool DiscardHeadersAndFootersInKeptSections { get; set; }
        public string InsertId { get; set; }

        public Source(string fileName)
        {
            WmlDocument = new WmlDocument(fileName);
            Start = 0;
            Count = int.MaxValue;
            KeepSections = false;
            InsertId = null;
        }

        public Source(WmlDocument source)
        {
            WmlDocument = source;
            Start = 0;
            Count = int.MaxValue;
            KeepSections = false;
            InsertId = null;
        }

        public Source(string fileName, bool keepSections)
        {
            WmlDocument = new WmlDocument(fileName);
            Start = 0;
            Count = int.MaxValue;
            KeepSections = keepSections;
            InsertId = null;
        }

        public Source(WmlDocument source, bool keepSections)
        {
            WmlDocument = source;
            Start = 0;
            Count = int.MaxValue;
            KeepSections = keepSections;
            InsertId = null;
        }

        public Source(string fileName, string insertId)
        {
            WmlDocument = new WmlDocument(fileName);
            Start = 0;
            Count = int.MaxValue;
            KeepSections = false;
            InsertId = insertId;
        }

        public Source(WmlDocument source, string insertId)
        {
            WmlDocument = source;
            Start = 0;
            Count = int.MaxValue;
            KeepSections = false;
            InsertId = insertId;
        }

        public Source(string fileName, int start, bool keepSections)
        {
            WmlDocument = new WmlDocument(fileName);
            Start = start;
            Count = int.MaxValue;
            KeepSections = keepSections;
            InsertId = null;
        }

        public Source(WmlDocument source, int start, bool keepSections)
        {
            WmlDocument = source;
            Start = start;
            Count = int.MaxValue;
            KeepSections = keepSections;
            InsertId = null;
        }

        public Source(string fileName, int start, string insertId)
        {
            WmlDocument = new WmlDocument(fileName);
            Start = start;
            Count = int.MaxValue;
            KeepSections = false;
            InsertId = insertId;
        }

        public Source(WmlDocument source, int start, string insertId)
        {
            WmlDocument = source;
            Start = start;
            Count = int.MaxValue;
            KeepSections = false;
            InsertId = insertId;
        }

        public Source(string fileName, int start, int count, bool keepSections)
        {
            WmlDocument = new WmlDocument(fileName);
            Start = start;
            Count = count;
            KeepSections = keepSections;
            InsertId = null;
        }

        public Source(WmlDocument source, int start, int count, bool keepSections)
        {
            WmlDocument = source;
            Start = start;
            Count = count;
            KeepSections = keepSections;
            InsertId = null;
        }

        public Source(string fileName, int start, int count, string insertId)
        {
            WmlDocument = new WmlDocument(fileName);
            Start = start;
            Count = count;
            KeepSections = false;
            InsertId = insertId;
        }

        public Source(WmlDocument source, int start, int count, string insertId)
        {
            WmlDocument = source;
            Start = start;
            Count = count;
            KeepSections = false;
            InsertId = insertId;
        }
    }
}