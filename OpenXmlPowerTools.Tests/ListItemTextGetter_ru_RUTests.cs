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
    }
}

#endif
