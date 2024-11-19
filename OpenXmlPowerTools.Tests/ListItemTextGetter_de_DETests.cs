// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using Xunit;
using OpenXmlPowerTools; 

#if !ELIDE_XUNIT_TESTS

namespace OpenXmlPowerTools.Tests
{
    public class ListItemTextGetter_de_DETests
    {
        [Theory]
        [InlineData(0, "Zero")]
        [InlineData(1, "1-ste")]
        [InlineData(2, "2-te")]
        [InlineData(3, "3-te")]
        [InlineData(4, "4-te")]
        [InlineData(5, "5-te")]
        [InlineData(6, "6-te")]
        [InlineData(7, "7-te")]
        [InlineData(8, "8-te")]
        [InlineData(9, "9-te")]
        [InlineData(10, "10-te")]
        [InlineData(11, "11-te")]
        [InlineData(12, "12-te")]
        [InlineData(13, "13-te")]
        [InlineData(14, "14-te")]
        [InlineData(16, "16-te")]
        [InlineData(17, "17-te")]
        [InlineData(18, "18-te")]
        [InlineData(19, "19-te")]
        [InlineData(20, "20-te")]
        [InlineData(23, "23-te")]
        [InlineData(25, "25-te")]
        [InlineData(50, "50-te")]
        [InlineData(56, "56-te")]
        [InlineData(67, "67-te")]
        [InlineData(78, "78-te")]
        [InlineData(100, "100-te")]
        [InlineData(123, "123-te")]
        [InlineData(125, "125-te")]
        [InlineData(1050, "1050-te")]
        public void GetListItemText_Ordinal(int integer, string expectedText)
        {
            string actualText = ListItemTextGetter_de_DE.GetListItemText("", integer, "ordinal"); 

            Assert.Equal(expectedText, actualText);
        }

        [Theory]
        [InlineData(0, "Zero")]
        [InlineData(1, "Erste")]
        [InlineData(2, "Zweite")]
        [InlineData(3, "Dritte")]
        [InlineData(4, "Vierte")]
        [InlineData(5, "Fünfte")]
        [InlineData(6, "Sechste")]
        [InlineData(7, "Siebte")]
        [InlineData(8, "Achte")]
        [InlineData(9, "Neunte")]
        [InlineData(10, "Zehnte")]
        [InlineData(11, "Elfte")]
        [InlineData(12, "Zwölfte")]
        [InlineData(13, "Dreizehnte")]
        [InlineData(14, "Vierzehnte")]
        [InlineData(16, "Sechzehnte")]
        [InlineData(17, "Siebzehnte")]
        [InlineData(18, "Achtzehnte")]
        [InlineData(19, "Neunzehnte")]
        [InlineData(20, "Zwanzigste")]
        [InlineData(23, "Dreiundzwanzigste")]
        [InlineData(25, "Fünfundzwanzigste")]
        [InlineData(50, "Fünfzigste")]
        [InlineData(56, "Sechsundfünfzigste")]
        [InlineData(67, "Siebenundsechzigste")]
        [InlineData(78, "Achtundsiebzigste")]
        [InlineData(100, "Einhundertste")]
        [InlineData(123, "Einhundertdreiundzwanzigste")]
        [InlineData(125, "Einhundertfünfundzwanzigste")]
        [InlineData(1050, "Eintausendfünfzigste")]
        public void GetListItemText_OrdinalText(int integer, string expectedText)
        {
            string actualText = ListItemTextGetter_de_DE.GetListItemText("", integer, "ordinalText"); 

            Assert.Equal(expectedText, actualText);
        }

        [Theory]
        [InlineData(0, "Zero")]
        [InlineData(1, "Eins")]
        [InlineData(2, "Zwei")]
        [InlineData(3, "Drei")]
        [InlineData(4, "Vier")]
        [InlineData(5, "Fünf")]
        [InlineData(6, "Sechs")]
        [InlineData(7, "Sieben")]
        [InlineData(8, "Acht")]
        [InlineData(9, "Neun")]
        [InlineData(10, "Zehn")]
        [InlineData(11, "Elf")]
        [InlineData(12, "Zwölf")]
        [InlineData(13, "Dreizehn")]
        [InlineData(14, "Vierzehn")]
        [InlineData(16, "Sechzehn")]
        [InlineData(17, "Siebzehn")]
        [InlineData(18, "Achtzehn")]
        [InlineData(19, "Neunzehn")]
        [InlineData(20, "Zwanzig")]
        [InlineData(23, "Dreiundzwanzig")]
        [InlineData(25, "Fünfundzwanzig")]
        [InlineData(50, "Fünfzig")]
        [InlineData(56, "Sechsundfünfzig")]
        [InlineData(67, "Siebenundsechzig")]
        [InlineData(78, "Achtundsiebzig")]
        [InlineData(100, "Einhundert")]
        [InlineData(123, "Einhundertdreiundzwanzig")]
        [InlineData(125, "Einhundertfünfundzwanzig")]
        [InlineData(1050, "Eintausendfünfzig")]
        public void GetListItemText_CardinalText(int integer, string expectedText)
        {
            string actualText = ListItemTextGetter_de_DE.GetListItemText("", integer, "cardinalText"); 

            Assert.Equal(expectedText, actualText);
        }
    }
}

#endif
