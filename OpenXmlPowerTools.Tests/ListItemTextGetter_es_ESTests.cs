// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using Xunit;
using OpenXmlPowerTools; 

#if !ELIDE_XUNIT_TESTS

namespace OpenXmlPowerTools.Tests
{
    public class ListItemTextGetter_es_ESTests
    {
        [Theory]
        [InlineData(0, "Cero")]
        [InlineData(1, "1-o")]
        [InlineData(2, "2-o")]
        [InlineData(3, "3-o")]
        [InlineData(4, "4-o")]
        [InlineData(5, "5-o")]
        [InlineData(6, "6-o")]
        [InlineData(7, "7-o")]
        [InlineData(8, "8-o")]
        [InlineData(9, "9-o")]
        [InlineData(10, "10-o")]
        [InlineData(11, "11-o")]
        [InlineData(12, "12-o")]
        [InlineData(13, "13-o")]
        [InlineData(14, "14-o")]
        [InlineData(16, "16-o")]
        [InlineData(17, "17-o")]
        [InlineData(18, "18-o")]
        [InlineData(19, "19-o")]
        [InlineData(20, "20-o")]
        [InlineData(23, "23-o")]
        [InlineData(25, "25-o")]
        [InlineData(50, "50-o")]
        [InlineData(56, "56-o")]
        [InlineData(67, "67-o")]
        [InlineData(78, "78-o")]
        [InlineData(100, "100-o")]
        [InlineData(123, "123-o")]
        [InlineData(125, "125-o")]
        [InlineData(1050, "1050-o")]
        public void GetListItemText_Ordinal(int integer, string expectedText)
        {
            string actualText = ListItemTextGetter_es_ES.GetListItemText("", integer, "ordinal"); 

            Assert.Equal(expectedText, actualText);
        }

        [Theory]
        [InlineData(0, "Cero")]
        [InlineData(1, "Primero")]
        [InlineData(2, "Segundo")]
        [InlineData(3, "Tercero")]
        [InlineData(4, "Cuarto")]
        [InlineData(5, "Quinto")]
        [InlineData(6, "Sexto")]
        [InlineData(7, "Séptimo")]
        [InlineData(8, "Octavo")]
        [InlineData(9, "Noveno")]
        [InlineData(10, "Décimo")]
        [InlineData(11, "Undécimo")]
        [InlineData(12, "Duodécimo")]
        [InlineData(13, "Decimotercio")]
        [InlineData(14, "Decimocuarto")]
        [InlineData(16, "Decimosexto")]
        [InlineData(17, "Decimoséptimo")]
        [InlineData(18, "Decimoctavo")]
        [InlineData(19, "Decimonoveno")]
        [InlineData(20, "Vigésimo")]
        [InlineData(23, "Vigésimo tercero")]
        [InlineData(25, "Vigésimo quinto")]
        [InlineData(50, "Quincuagésimo")]
        [InlineData(56, "Quincuagésimo sexto")]
        [InlineData(67, "Sexagésimo séptimo")]
        [InlineData(78, "Septuagésimo octavo")]
        [InlineData(100, "Centésimo")]
        [InlineData(123, "Centésimo vigésimo tercero")]
        [InlineData(125, "Centésimo vigésimo quinto")]
        [InlineData(1050, "Milesimo quincuagésimo")]
        public void GetListItemText_OrdinalText(int integer, string expectedText)
        {
            string actualText = ListItemTextGetter_es_ES.GetListItemText("", integer, "ordinalText"); 

            Assert.Equal(expectedText, actualText);
        }

        [Theory]
        [InlineData(0, "Cero")]
        [InlineData(1, "Uno")]
        [InlineData(2, "Dos")]
        [InlineData(3, "Tres")]
        [InlineData(4, "Cuatro")]
        [InlineData(5, "Cinco")]
        [InlineData(6, "Seis")]
        [InlineData(7, "Siete")]
        [InlineData(8, "Ocho")]
        [InlineData(9, "Nueve")]
        [InlineData(10, "Diez")]
        [InlineData(11, "Once")]
        [InlineData(12, "Doce")]
        [InlineData(13, "Trece")]
        [InlineData(14, "Catorce")]
        [InlineData(16, "Dieciséis")]
        [InlineData(17, "Diecisiete")]
        [InlineData(18, "Dieciocho")]
        [InlineData(19, "Diecinueve")]
        [InlineData(20, "Veinte")]
        [InlineData(23, "Veinte y tres")]
        [InlineData(25, "Veinte y cinco")]
        [InlineData(50, "Cincuenta")]
        [InlineData(56, "Cincuenta y seis")]
        [InlineData(67, "Sesenta y siete")]
        [InlineData(78, "Setenta y ocho")]
        [InlineData(100, "Cien")]
        [InlineData(123, "Ciento veinte y tres")]
        [InlineData(125, "Ciento veinte y cinco")]
        [InlineData(1050, "Mil cincuenta")]
        public void GetListItemText_CardinalText(int integer, string expectedText)
        {
            string actualText = ListItemTextGetter_es_ES.GetListItemText("", integer, "cardinalText"); 

            Assert.Equal(expectedText, actualText);
        }
    }
}

#endif
