

// It is possible to optimize DescendantContentAtoms

/// Currently, the unid is set at the beginning of the algorithm.  It is used by the code that establishes correlation based on first rejecting
/// tracked revisions, then correlating paragraphs/tables.  It is requred for this algorithm - after finding a correlated sequence in the document with rejected
/// revisions, it uses the unid to find the same paragraph in the document without rejected revisions, then sets the correlated sha1 hash in that document.
///
/// But then when accepting tracked revisions, for certain paragraphs (where there are deleted paragraph marks) it is going to lose the unids.  But this isn't a
/// problem because when paragraph marks are deleted, the correlation is definitely no longer possible.  Any paragraphs that are in a range of paragraphs that
/// are coalesced can't be correlated to paragraphs in the other document via their hash.  At that point we no longer care what their unids are.
///
/// But after that it is only used to reconstruct the tree.  It is also used in the debugging code that
/// prints the various correlated sequences and comparison units - this is display for debugging purposes only.

/// The key idea here is that a given paragraph will always have the same ancestors, and it doesn't matter whether the content was deleted from the old document,
/// inserted into the new document, or set as equal.  At this point, we identify a paragraph as a sequential list of content atoms, terminated by a paragraph mark.
/// This entire list will for a single paragraph, regardless of whether the paragraph is a child of the body, or if the paragraph is in a cell in a table, or if
/// the paragraph is in a text box.  The list of ancestors, from the paragraph to the root of the XML tree will be the same for all content atoms in the paragraph.
///
/// Therefore:
///
/// Iterate through the list of content atoms backwards.  When the loop sees a paragraph mark, it gets the ancestor unids from the paragraph mark to the top of the
/// tree, and sets this as the same for all content atoms in the paragraph.  For descendants of the paragraph mark, it doesn't really matter if content is put into
/// separate runs or what not.  We don't need to be concerned about what the unids are for descendants of the paragraph.

namespace OpenXmlPowerTools
{
    public enum CorrelationStatus
    {
        Nil,
        Normal,
        Unknown,
        Inserted,
        Deleted,
        Equal,
        Group,
    }
}