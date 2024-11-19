// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using Xunit;
using OpenXmlPowerTools; 

#if !ELIDE_XUNIT_TESTS

namespace OpenXmlPowerTools.Tests
{
    public class ListItemTextGetter_ru_RUTests
    {
        [Theory]
        [InlineData(0, "Ноль")]
        [InlineData(1, "1-ый")]
        [InlineData(2, "2-ой")]
        [InlineData(3, "3-ий")]
        [InlineData(4, "4-ый")]
        [InlineData(5, "5-ый")]
        [InlineData(6, "6-ой")]
        [InlineData(7, "7-ой")]
        [InlineData(8, "8-ой")]
        [InlineData(9, "9-ый")]
        [InlineData(10, "10-ый")]
        [InlineData(11, "11-ый")]
        [InlineData(12, "12-ый")]
        [InlineData(13, "13-ый")]
        [InlineData(14, "14-ый")]
        [InlineData(16, "16-ый")]
        [InlineData(17, "17-ый")]
        [InlineData(18, "18-ый")]
        [InlineData(19, "19-ый")]
        [InlineData(20, "20-ый")]
        [InlineData(23, "23-ий")]
        [InlineData(25, "25-ый")]
        [InlineData(50, "50-ый")]
        [InlineData(56, "56-ой")]
        [InlineData(67, "67-ой")]
        [InlineData(78, "78-ой")]
        [InlineData(100, "100-ый")]
        [InlineData(123, "123-ий")]
        [InlineData(125, "125-ый")]
        [InlineData(1050, "1050-ый")]
        public void GetListItemText_Ordinal(int integer, string expectedText)
        {
            string actualText = ListItemTextGetter_ru_RU.GetListItemText("", integer, "ordinal"); 

            Assert.Equal(expectedText, actualText);
        }

        [Theory]
        [InlineData(0, "Ноль")]
        [InlineData(1, "Первый")]
        [InlineData(2, "Второй")]
        [InlineData(3, "Третий")]
        [InlineData(4, "Четвертый")]
        [InlineData(5, "Пятый")]
        [InlineData(6, "Шестой")]
        [InlineData(7, "Седьмой")]
        [InlineData(8, "Восьмой")]
        [InlineData(9, "Девятый")]
        [InlineData(10, "Десятый")]
        [InlineData(11, "Одиннадцатый")]
        [InlineData(12, "Двенадцатый")]
        [InlineData(13, "Тринадцатый")]
        [InlineData(14, "Четырнадцатый")]
        [InlineData(16, "Шестнадцатый")]
        [InlineData(17, "Семнадцатый")]
        [InlineData(18, "Восемнадцатый")]
        [InlineData(19, "Девятнадцатый")]
        [InlineData(20, "Двадцатый")]
        [InlineData(23, "Двадцать третий")]
        [InlineData(25, "Двадцать пятый")]
        [InlineData(50, "Пятидесятый")]
        [InlineData(56, "Пятьдесят шестой")]
        [InlineData(67, "Шестьдесят седьмой")]
        [InlineData(78, "Семьдесят восьмой")]
        [InlineData(100, "Сотый")]
        [InlineData(123, "Сто двадцать третий")]
        [InlineData(125, "Сто двадцать пятый")]
        [InlineData(1050, "Одна тысяча пятидесятый")]
        public void GetListItemText_OrdinalText(int integer, string expectedText)
        {
            string actualText = ListItemTextGetter_ru_RU.GetListItemText("", integer, "ordinalText"); 

            Assert.Equal(expectedText, actualText);
        }

        [Theory]
        [InlineData(0, "Ноль")]
        [InlineData(1, "Один")]
        [InlineData(2, "Два")]
        [InlineData(3, "Три")]
        [InlineData(4, "Четыре")]
        [InlineData(5, "Пять")]
        [InlineData(6, "Шесть")]
        [InlineData(7, "Семь")]
        [InlineData(8, "Восемь")]
        [InlineData(9, "Девять")]
        [InlineData(10, "Десять")]
        [InlineData(11, "Одиннадцать")]
        [InlineData(12, "Двенадцать")]
        [InlineData(13, "Тринадцать")]
        [InlineData(14, "Четырнадцать")]
        [InlineData(16, "Шестнадцать")]
        [InlineData(17, "Семнадцать")]
        [InlineData(18, "Восемнадцать")]
        [InlineData(19, "Девятнадцать")]
        [InlineData(20, "Двадцать")]
        [InlineData(23, "Двадцать три")]
        [InlineData(25, "Двадцать пять")]
        [InlineData(50, "Пятьдесят")]
        [InlineData(56, "Пятьдесят шесть")]
        [InlineData(67, "Шестьдесят семь")]
        [InlineData(78, "Семьдесят восемь")]
        [InlineData(100, "Сто")]
        [InlineData(123, "Сто двадцать три")]
        [InlineData(125, "Сто двадцать пять")]
        [InlineData(1050, "Одна тысяча пятьдесят")]
        public void GetListItemText_CardinalText(int integer, string expectedText)
        {
            string actualText = ListItemTextGetter_ru_RU.GetListItemText("", integer, "cardinalText"); 

            Assert.Equal(expectedText, actualText);
        }
    }
}

#endif
