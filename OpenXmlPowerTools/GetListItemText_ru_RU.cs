// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OpenXmlPowerTools
{
    public class ListItemTextGetter_ru_RU
    {
        private static string[] OneThroughNineteen = {
            "один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь",
            "девять", "десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать",
            "пятнадцать", "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать"
        };

        private static string[] Tens = {
            "десять", "двадцать", "тридцать", "сорок", "пятьдесят", "шестьдесят", "семьдесят",
            "восемьдесят", "девяносто"
        };

        private static string[] OrdinalOneThroughNineteen = {
            "первый", "второй", "третий", "четвертый", "пятый", "шестой",
            "седьмой", "восьмой", "девятый", "десятый", "одиннадцатый", "двенадцатый",
            "тринадцатый", "четырнадцатый", "пятнадцатый", "шестнадцатый",
            "семнадцатый", "восемнадцатый", "девятнадцатый"
        };

        private static string[] OrdinalTenths = {
            "десятый", "двадцатый", "тридцатый", "сороковой", "пятидесятый",
            "шестидесятый", "семидесятый", "восьмидесятый", "девяностый"
        };

        // TODO this is not correct for values above 99

        public static string GetListItemText(string languageCultureName, int levelNumber, string numFmt)
        {
            if (numFmt == "cardinalText")
            {
                string result = "";
                int t1 = levelNumber / 1000;
                int t2 = levelNumber % 1000;
                if (t1 >= 1)
                    result += OneThroughNineteen[t1 - 1] + " thousand";
                if (t1 >= 1 && t2 == 0)
                    return result.Substring(0, 1).ToUpper() +
                        result.Substring(1);
                if (t1 >= 1)
                    result += " ";
                int h1 = (levelNumber % 1000) / 100;
                int h2 = levelNumber % 100;
                if (h1 >= 1)
                    result += OneThroughNineteen[h1 - 1] + " hundred";
                if (h1 >= 1 && h2 == 0)
                    return result.Substring(0, 1).ToUpper() +
                        result.Substring(1);
                if (h1 >= 1)
                    result += " ";
                int z = levelNumber % 100;
                if (z <= 19)
                    result += OneThroughNineteen[z - 1];
                else
                {
                    int x = z / 10;
                    int r = z % 10;
                    result += Tens[x - 1];
                    if (r >= 1)
                        result += "-" + OneThroughNineteen[r - 1];
                }
                return result.Substring(0, 1).ToUpper() +
                    result.Substring(1);
            }
            if (numFmt == "ordinalText")
            {
                string result = "";
                int t1 = levelNumber / 1000;
                int t2 = levelNumber % 1000;
                if (t1 >= 1 && t2 != 0)
                    result += OneThroughNineteen[t1 - 1] + " thousand";
                if (t1 >= 1 && t2 == 0)
                {
                    result += OneThroughNineteen[t1 - 1] + " thousandth";
                    return result.Substring(0, 1).ToUpper() +
                        result.Substring(1);
                }
                if (t1 >= 1)
                    result += " ";
                int h1 = (levelNumber % 1000) / 100;
                int h2 = levelNumber % 100;
                if (h1 >= 1 && h2 != 0)
                    result += OneThroughNineteen[h1 - 1] + " hundred";
                if (h1 >= 1 && h2 == 0)
                {
                    result += OneThroughNineteen[h1 - 1] + " hundredth";
                    return result.Substring(0, 1).ToUpper() +
                        result.Substring(1);
                }
                if (h1 >= 1)
                    result += " ";
                int z = levelNumber % 100;
                if (z <= 19)
                    result += OrdinalOneThroughNineteen[z - 1];
                else
                {
                    int x = z / 10;
                    int r = z % 10;
                    if (r == 0)
                        result += OrdinalTenths[x - 1];
                    else
                        result += Tens[x - 1];
                    if (r >= 1)
                        result += " " + OrdinalOneThroughNineteen[r - 1];
                }
                return result.Substring(0, 1).ToUpper() +
                    result.Substring(1);
            }
            return null;
        }
    }
}
