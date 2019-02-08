// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

/***************************************************************************

Copyright (c) Microsoft Corporation 2016.

This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license can be found here:

http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx

Developer: Thomas Barnekow
Email: thomas@barnekow.info

***************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public class UnicodeMapper
    {
        // Unicode character values.
        public static readonly char StartOfHeading = '\u0001';
        public static readonly char HorizontalTabulation = '\u0009';
        public static readonly char LineFeed = '\u000A';
        public static readonly char FormFeed = '\u000C';
        public static readonly char CarriageReturn = '\u000D';
        public static readonly char SoftHyphen = '\u00AD';
        public static readonly char NonBreakingHyphen = '\u2011';

        // Unicode area boundaries.
        public static readonly char StartOfPrivateUseArea = '\uE000';
        public static readonly char StartOfSymbolArea = '\uF000';
        public static readonly char EndOfPrivateUseArea = '\uF8FF';

        // Dictionaries for w:sym stringification.
        private static readonly Dictionary<string, char> SymStringToUnicodeCharDictionary =
            new Dictionary<string, char>();

        private static readonly Dictionary<char, XElement> UnicodeCharToSymDictionary =
            new Dictionary<char, XElement>();

        // Represents the Unicode value that was last used to map an actual character
        // onto a special value in the private use area, which starts at U+E000.
        // In Open XML, U+F000 is added to the actual Unicode values, so we should be
        // well outside that range and would have to map 4096 different characters
        // to get into the area starting at U+F000.
        private static char _lastUnicodeChar = StartOfPrivateUseArea;

        /// <summary>
        /// Stringify an Open XML run, turning (a) w:t, w:br, w:cr, w:noBreakHyphen,
        /// w:softHyphen,  w:sym, and w:tab into their corresponding Unicode strings
        /// and (b) everything else into U+0001.
        /// </summary>
        /// <param name="element">An Open XML run or run child element.</param>
        /// <returns>The corresponding Unicode value or U+0001.</returns>
        public static string RunToString(XElement element)
        {
            if (element.Name == W.r && (element.Parent == null || element.Parent.Name != W.del))
                return element.Elements().Select(RunToString).StringConcatenate();

            // We need to ignore run properties.
            if (element.Name == W.rPr)
                return string.Empty;

            // For w:t elements, we obviously want the element's value.
            if (element.Name == W.t)
                return (string) element;

            // Turn elements representing special characters into their corresponding
            // unicode characters.
            if (element.Name == W.br)
            {
                XAttribute typeAttribute = element.Attribute(W.type);
                string type = typeAttribute != null ? typeAttribute.Value : null;
                if (type == null || type == "textWrapping")
                    return CarriageReturn.ToString();
                if (type == "page")
                    return FormFeed.ToString();
            }

            if (element.Name == W.cr)
                return CarriageReturn.ToString();
            if (element.Name == W.noBreakHyphen)
                return NonBreakingHyphen.ToString();
            if (element.Name == W.softHyphen)
                return SoftHyphen.ToString();
            if (element.Name == W.tab)
                return HorizontalTabulation.ToString();

            if (element.Name == W.fldChar)
            {
                var fldCharType = element.Attributes(W.fldCharType).Select(a => a.Value).FirstOrDefault();
                switch (fldCharType)
                {
                    case "begin":
                        return "{";
                    case "end":
                        return "}";
                    default:
                        return "_";
                }
            }

            if (element.Name == W.instrText)
                return "_";

            // Turn w:sym elements into Unicode character values. A w:char attribute
            // value can be stored (a) directly in its Unicode character value from
            // the font glyph or (b) in a Unicode character value created by adding
            // U+F000 to the character value, thereby shifting the value into the
            // Unicode private use area.
            if (element.Name == W.sym)
                return SymToChar(element).ToString();

            // Elements we don't recognize will be turned into a character that
            // doesn't typically appear in documents.
            return StartOfHeading.ToString();
        }

        /// <summary>
        /// Translate a symbol into a Unicode character, using the specified w:font attribute
        /// value and unicode value (represented by the w:sym element's w:char attribute),
        /// using a substitute value for the actual Unicode value if the same Unicode value
        /// is already used in conjunction with a different w:font attribute value.
        ///
        /// Add U+F000 to the Unicode value if the specified value is less than U+1000, which
        /// shifts the value into the Unicode private use area (which is also done by MS Word).
        /// </summary>
        /// <remarks>
        /// For w:sym elements, the w:char attribute value is typically greater than "F000",
        /// because U+F000 is added to the actual Unicode value to shift the value into
        /// the Unicode private use area.
        /// </remarks>
        /// <param name="fontAttributeValue">The w:font attribute value, e.g., "Wingdings".</param>
        /// <param name="unicodeValue">The unicode value.</param>
        /// <returns>The Unicode character used to represent the symbol.</returns>
        public static char SymToChar(string fontAttributeValue, char unicodeValue)
        {
            return SymToChar(fontAttributeValue, (int) unicodeValue);
        }

        /// <summary>
        /// Translate a symbol into a Unicode character, using the specified w:font attribute
        /// value and unicode value (represented by the w:sym element's w:char attribute),
        /// using a substitute value for the actual Unicode value if the same Unicode value
        /// is already used in conjunction with a different w:font attribute value.
        ///
        /// Add U+F000 to the Unicode value if the specified value is less than U+1000, which
        /// shifts the value into the Unicode private use area (which is also done by MS Word).
        /// </summary>
        /// <remarks>
        /// For w:sym elements, the w:char attribute value is typically greater than "F000",
        /// because U+F000 is added to the actual Unicode value to shift the value into
        /// the Unicode private use area.
        /// </remarks>
        /// <param name="fontAttributeValue">The w:font attribute value, e.g., "Wingdings".</param>
        /// <param name="unicodeValue">The unicode value.</param>
        /// <returns>The Unicode character used to represent the symbol.</returns>
        public static char SymToChar(string fontAttributeValue, int unicodeValue)
        {
            int effectiveUnicodeValue = unicodeValue < 0x1000 ? 0xF000 + unicodeValue : unicodeValue;
            return SymToChar(fontAttributeValue, effectiveUnicodeValue.ToString("X4"));
        }

        /// <summary>
        /// Translate a symbol into a Unicode character, using the specified w:font and
        /// w:char attribute values, using a substitute value for the actual Unicode
        /// value if the same Unicode value is already used in conjunction with a different
        /// w:font attribute value.
        ///
        /// Do not alter the w:char attribute value.
        /// </summary>
        /// <remarks>
        /// For w:sym elements, the w:char attribute value is typically greater than "F000",
        /// because U+F000 is added to the actual Unicode value to shift the value into
        /// the Unicode private use area.
        /// </remarks>
        /// <param name="fontAttributeValue">The w:font attribute value, e.g., "Wingdings".</param>
        /// <param name="charAttributeValue">The w:char attribute value, e.g., "F028".</param>
        /// <returns>The Unicode character used to represent the symbol.</returns>
        public static char SymToChar(string fontAttributeValue, string charAttributeValue)
        {
            if (string.IsNullOrEmpty(fontAttributeValue))
                throw new ArgumentException("Argument is null or empty.", "fontAttributeValue");
            if (string.IsNullOrEmpty(charAttributeValue))
                throw new ArgumentException("Argument is null or empty.", "charAttributeValue");

            return SymToChar(new XElement(W.sym,
                new XAttribute(W.font, fontAttributeValue),
                new XAttribute(W._char, charAttributeValue),
                new XAttribute(XNamespace.Xmlns + "w", W.w)));
        }

        /// <summary>
        /// Represent a w:sym element as a Unicode value, mapping the Unicode value
        /// specified in the w:char attribute to a substitute value to be able to
        /// use a Unicode value in conjunction with different fonts.
        /// </summary>
        /// <param name="sym">The w:sym element to be stringified.</param>
        /// <returns>A single-character Unicode string representing the w:sym element.</returns>
        public static char SymToChar(XElement sym)
        {
            if (sym == null)
                throw new ArgumentNullException("sym");
            if (sym.Name != W.sym)
                throw new ArgumentException(string.Format("Not a w:sym: {0}", sym.Name), "sym");

            XAttribute fontAttribute = sym.Attribute(W.font);
            string fontAttributeValue = fontAttribute != null ? fontAttribute.Value : null;
            if (fontAttributeValue == null)
                throw new ArgumentException("w:sym element has no w:font attribute.", "sym");

            XAttribute charAttribute = sym.Attribute(W._char);
            string charAttributeValue = charAttribute != null ? charAttribute.Value : null;
            if (charAttributeValue == null)
                throw new ArgumentException("w:sym element has no w:char attribute.", "sym");

            // Return Unicode value if it is in the dictionary.
            var standardizedSym = new XElement(W.sym,
                new XAttribute(W.font, fontAttributeValue),
                new XAttribute(W._char, charAttributeValue),
                new XAttribute(XNamespace.Xmlns + "w", W.w));
            string standardizedSymString = standardizedSym.ToString(SaveOptions.None);
            if (SymStringToUnicodeCharDictionary.ContainsKey(standardizedSymString))
                return SymStringToUnicodeCharDictionary[standardizedSymString];

            // Determine Unicode value to be used to represent the current w:sym element.
            // Use the actual Unicode value if it has not yet been used with another font.
            // Otherwise, create a special Unicode value in the private use area to represent
            // the current w:sym element.
            var unicodeChar = (char) Convert.ToInt32(charAttributeValue, 16);
            if (UnicodeCharToSymDictionary.ContainsKey(unicodeChar))
                unicodeChar = ++_lastUnicodeChar;

            SymStringToUnicodeCharDictionary.Add(standardizedSymString, unicodeChar);
            UnicodeCharToSymDictionary.Add(unicodeChar, standardizedSym);
            return unicodeChar;
        }

        /// <summary>
        /// Turn the specified text value into a list of runs with coalesced text elements.
        /// Each run will have the specified run properties.
        /// </summary>
        /// <param name="textValue">The text value to transform.</param>
        /// <param name="runProperties">The run properties to apply.</param>
        /// <returns>A list of runs representing the text value.</returns>
        public static List<XElement> StringToCoalescedRunList(string textValue, XElement runProperties)
        {
            return textValue
                .Select(CharToRunChild)
                .GroupAdjacent(e => e.Name == W.t)
                .SelectMany(grouping => grouping.Key
                    ? StringToSingleRunList(grouping.Select(t => (string) t).StringConcatenate(), runProperties)
                    : grouping.Select(e => new XElement(W.r, runProperties, e)))
                .ToList();
        }

        /// <summary>
        /// Turn the specified text value into a list consisting of a single run having one
        /// text element with that text value. The run will have the specified run properties.
        /// </summary>
        /// <param name="textValue">The text value to transform.</param>
        /// <param name="runProperties">The run properties to apply.</param>
        /// <returns>A list with a single run.</returns>
        public static IEnumerable<XElement> StringToSingleRunList(string textValue, XElement runProperties)
        {
            var run = new XElement(W.r,
                runProperties,
                new XElement(W.t, XmlUtil.GetXmlSpaceAttribute(textValue), textValue));
            return new List<XElement> { run };
        }

        /// <summary>
        /// Turn the specified text value into a list of runs, each having the specified
        /// run properties.
        /// </summary>
        /// <param name="textValue">The text value to transform.</param>
        /// <param name="runProperties">The run properties to apply.</param>
        /// <returns>A list of runs representing the text value.</returns>
        public static List<XElement> StringToRunList(string textValue, XElement runProperties)
        {
            return textValue.Select(character => CharToRun(character, runProperties)).ToList();
        }

        /// <summary>
        /// Create a w:r element from the specified character, which will be turned
        /// into a corresponding Open XML element (e.g., w:t, w:br, w:tab).
        /// </summary>
        /// <param name="character">The character.</param>
        /// <param name="runProperties">The w:rPr element to be added to the w:r element.</param>
        /// <returns>The w:r element.</returns>
        public static XElement CharToRun(char character, XElement runProperties)
        {
            return new XElement(W.r, runProperties, CharToRunChild(character));
        }

        /// <summary>
        /// Create an Open XML element (e.g., w:t, w:br, w:tab) from the specified
        /// character.
        /// </summary>
        /// <param name="character">The character.</param>
        /// <returns>The Open XML element or null, if the character equals <see cref="StartOfHeading" /> (U+0001).</returns>
        public static XElement CharToRunChild(char character)
        {
            // Ignore the special character that represents the Open XML elements we
            // wanted to ignore.
            if (character == StartOfHeading)
                return null;

            // Translate special characters into their corresponding Open XML elements.
            // Turn a Carriage Return into an empty w:br element, regardless of whether
            // the former was created from an equivalent w:cr element.
            if (character == CarriageReturn)
                return new XElement(W.br);
            if (character == FormFeed)
                return new XElement(W.br, new XAttribute(W.type, "page"));
            if (character == HorizontalTabulation)
                return new XElement(W.tab);
            if (character == NonBreakingHyphen)
                return new XElement(W.noBreakHyphen);
            if (character == SoftHyphen)
                return new XElement(W.softHyphen);

            // Translate symbol characters into their corresponding w:sym elements.
            if (UnicodeCharToSymDictionary.ContainsKey(character))
                return UnicodeCharToSymDictionary[character];

            // Turn "normal" characters into text elements.
            return new XElement(W.t, XmlUtil.GetXmlSpaceAttribute(character), character);
        }
    }
}
