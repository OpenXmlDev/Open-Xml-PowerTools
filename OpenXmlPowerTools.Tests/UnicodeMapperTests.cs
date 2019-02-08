// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

/***************************************************************************

Copyright (c) Microsoft Corporation 2016.

This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license can be found here:

http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx

Developer: Thomas Barnekow
Email: thomas@barnekow.info

***************************************************************************/

using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Xunit;

#if !ELIDE_XUNIT_TESTS

namespace OpenXmlPowerTools.Tests
{
    public class UnicodeMapperTests
    {
        [Fact]
        public void CanStringifyRunAndTextElements()
        {
            const string textValue = "Hello World!";
            var textElement = new XElement(W.t, textValue);
            var runElement = new XElement(W.r, textElement);
            var formattedRunElement = new XElement(W.r, new XElement(W.rPr, new XElement(W.b)), textElement);

            Assert.Equal(textValue, UnicodeMapper.RunToString(textElement));
            Assert.Equal(textValue, UnicodeMapper.RunToString(runElement));
            Assert.Equal(textValue, UnicodeMapper.RunToString(formattedRunElement));
        }

        [Fact]
        public void CanStringifySpecialElements()
        {
            Assert.Equal(UnicodeMapper.CarriageReturn,
                UnicodeMapper.RunToString(new XElement(W.cr)).First());
            Assert.Equal(UnicodeMapper.CarriageReturn,
                UnicodeMapper.RunToString(new XElement(W.br)).First());
            Assert.Equal(UnicodeMapper.FormFeed,
                UnicodeMapper.RunToString(new XElement(W.br, new XAttribute(W.type, "page"))).First());
            Assert.Equal(UnicodeMapper.NonBreakingHyphen,
                UnicodeMapper.RunToString(new XElement(W.noBreakHyphen)).First());
            Assert.Equal(UnicodeMapper.SoftHyphen,
                UnicodeMapper.RunToString(new XElement(W.softHyphen)).First());
            Assert.Equal(UnicodeMapper.HorizontalTabulation,
                UnicodeMapper.RunToString(new XElement(W.tab)).First());
        }

        [Fact]
        public void CanCreateRunChildElementsFromSpecialCharacters()
        {
            Assert.Equal(W.br, UnicodeMapper.CharToRunChild(UnicodeMapper.CarriageReturn).Name);
            Assert.Equal(W.noBreakHyphen, UnicodeMapper.CharToRunChild(UnicodeMapper.NonBreakingHyphen).Name);
            Assert.Equal(W.softHyphen, UnicodeMapper.CharToRunChild(UnicodeMapper.SoftHyphen).Name);
            Assert.Equal(W.tab, UnicodeMapper.CharToRunChild(UnicodeMapper.HorizontalTabulation).Name);

            XElement element = UnicodeMapper.CharToRunChild(UnicodeMapper.FormFeed);
            Assert.Equal(W.br, element.Name);
            Assert.Equal("page", element.Attribute(W.type).Value);

            Assert.Equal(W.br, UnicodeMapper.CharToRunChild('\r').Name);
        }

        [Fact]
        public void CanCreateCoalescedRuns()
        {
            const string textString = "This is only text.";
            const string mixedString = "First\tSecond\tThird";

            List<XElement> textRuns = UnicodeMapper.StringToCoalescedRunList(textString, null);
            List<XElement> mixedRuns = UnicodeMapper.StringToCoalescedRunList(mixedString, null);

            Assert.Single(textRuns);
            Assert.Equal(5, mixedRuns.Count);

            Assert.Equal("First", mixedRuns.Elements(W.t).Skip(0).First().Value);
            Assert.Equal("Second", mixedRuns.Elements(W.t).Skip(1).First().Value);
            Assert.Equal("Third", mixedRuns.Elements(W.t).Skip(2).First().Value);
        }

        [Fact]
        public void CanMapSymbols()
        {
            var sym1 = new XElement(W.sym,
                new XAttribute(W.font, "Wingdings"),
                new XAttribute(W._char, "F028"));
            char charFromSym1 = UnicodeMapper.SymToChar(sym1);
            XElement symFromChar1 = UnicodeMapper.CharToRunChild(charFromSym1);

            var sym2 = new XElement(W.sym,
                new XAttribute(W._char, "F028"),
                new XAttribute(W.font, "Wingdings"));
            char charFromSym2 = UnicodeMapper.SymToChar(sym2);

            var sym3 = new XElement(W.sym,
                new XAttribute(XNamespace.Xmlns + "w", W.w),
                new XAttribute(W.font, "Wingdings"),
                new XAttribute(W._char, "F028"));
            char charFromSym3 = UnicodeMapper.SymToChar(sym3);

            var sym4 = new XElement(W.sym,
                new XAttribute(XNamespace.Xmlns + "w", W.w),
                new XAttribute(W.font, "Webdings"),
                new XAttribute(W._char, "F028"));
            char charFromSym4 = UnicodeMapper.SymToChar(sym4);
            XElement symFromChar4 = UnicodeMapper.CharToRunChild(charFromSym4);

            Assert.Equal(charFromSym1, charFromSym2);
            Assert.Equal(charFromSym1, charFromSym3);
            Assert.NotEqual(charFromSym1, charFromSym4);

            Assert.Equal("F028", symFromChar1.Attribute(W._char).Value);
            Assert.Equal("Wingdings", symFromChar1.Attribute(W.font).Value);

            Assert.Equal("F028", symFromChar4.Attribute(W._char).Value);
            Assert.Equal("Webdings", symFromChar4.Attribute(W.font).Value);
        }

        [Fact]
        public void CanStringifySymbols()
        {
            char charFromSym1 = UnicodeMapper.SymToChar("Wingdings", '\uF028');
            char charFromSym2 = UnicodeMapper.SymToChar("Wingdings", 0xF028);
            char charFromSym3 = UnicodeMapper.SymToChar("Wingdings", "F028");

            XElement symFromChar1 = UnicodeMapper.CharToRunChild(charFromSym1);
            XElement symFromChar2 = UnicodeMapper.CharToRunChild(charFromSym2);
            XElement symFromChar3 = UnicodeMapper.CharToRunChild(charFromSym3);

            Assert.Equal(charFromSym1, charFromSym2);
            Assert.Equal(charFromSym1, charFromSym3);

            Assert.Equal(symFromChar1.ToString(SaveOptions.None), symFromChar2.ToString(SaveOptions.None));
            Assert.Equal(symFromChar1.ToString(SaveOptions.None), symFromChar3.ToString(SaveOptions.None));
        }
    }
}

#endif
