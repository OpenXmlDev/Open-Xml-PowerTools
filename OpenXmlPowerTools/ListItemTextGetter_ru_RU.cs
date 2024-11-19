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

        private static string[] Hundreds = {
            "сто", "двести", "триста", "четыреста", "пятьсот", "шестьсот", "семьсот",
            "восемьсот", "девятьсот"
        };

        private static string[] OrdinalOneThroughNineteen = {
            "первый", "второй", "третий", "четвертый", "пятый", "шестой",
            "седьмой", "восьмой", "девятый", "десятый", "одиннадцатый", "двенадцатый",
            "тринадцатый", "четырнадцатый", "пятнадцатый", "шестнадцатый",
            "семнадцатый", "восемнадцатый", "девятнадцатый"
        };

        private static string[] OrdinalTens = {
            "десятый", "двадцатый", "тридцатый", "сороковой", "пятидесятый",
            "шестидесятый", "семидесятый", "восьмидесятый", "девяностый"
        };
        
        private static string[] OrdinalOneThroughNineteenHT = {
            "одно", "двух", "трёх", "четырёх", "пяти", "шести", "семи", "восьми", "девяти", 
            "десяти", "одиннадцати", "двеннадцати", "четырнадцати", "пятнадцати", "шестнадцати", 
            "семнадцати", "восемьнадцати", "девятнадцати"
        };

        public static string GetListItemText(string languageCultureName, int levelNumber, string numFmt)
        {
			if (levelNumber > 19999)
				throw new ArgumentOutOfRangeException("levelNumber", "Converting a number greater than 19999 to text is not supported");
			if (levelNumber == 0)
				return "Ноль";
			if (levelNumber < 0)
				throw new ArgumentOutOfRangeException("levelNumber", "Converting a negative number to text is not supported");

            if (numFmt == "ordinal")
                return GetOrdinal(levelNumber); 
            if (numFmt == "cardinalText")
                return GetCardinalText(levelNumber); 
            if (numFmt == "ordinalText")
                return GetOrdinalText(levelNumber); 
            return null;
        }

        private static string GetOrdinal(int levelNumber)
        {
            string suffix;
            if (levelNumber % 100 == 12 || levelNumber % 100 == 13 || levelNumber % 100 == 16 || 
                levelNumber % 100 == 17 || levelNumber % 100 == 18)
                suffix = "-ый";
            else if (levelNumber % 10 == 2 || levelNumber % 10 == 6 || levelNumber % 10 == 7 || levelNumber % 10 == 8)
                suffix = "-ой";
            else if (levelNumber % 10 == 3)
                suffix = "-ий";
            else
                suffix = "-ый";
            return levelNumber.ToString() + suffix;
        }

        private static string GetCardinalText(int levelNumber)
        {
            string result = "";

            // Get thousands.
            int t1 = levelNumber / 1000;
            int t2 = levelNumber % 1000;
            if (t1 >= 1)
            {
                if (t1 == 1) 
                    result += "одна тысяча";
                else if (t1 == 2)
                    result += "две тысячи";
                else if (t1 == 3 || t1 == 4)
                    result += OneThroughNineteen[t1 - 1] + " тысячи";
                else
                    result += OneThroughNineteen[t1 - 1] + " тысяч";
            }
            if (t1 >= 1 && t2 == 0)
                return result.Substring(0, 1).ToUpper() + result.Substring(1);
            if (t1 >= 1)
                result += " ";
            
            // Get hundreds.
            int h1 = (levelNumber % 1000) / 100;
            int h2 = levelNumber % 100;
            if (h1 >= 1)
                result += Hundreds[h1 - 1];
            if (h1 >= 1 && h2 == 0)
                return result.Substring(0, 1).ToUpper() + result.Substring(1);
            if (h1 >= 1)
                result += " ";
            
            // Tens and ones.
            int z = levelNumber % 100;
            if (z <= 19)
                result += OneThroughNineteen[z - 1];
            else
            {
                int x = z / 10;
                int r = z % 10;
                result += Tens[x - 1];
                if (r >= 1)
                    result += " " + OneThroughNineteen[r - 1];
            }
            return result.Substring(0, 1).ToUpper() + result.Substring(1);
        }

        private static string GetOrdinalText(int levelNumber)
        {
            string result = "";

            // Get thousands.
            int t1 = levelNumber / 1000;
            int t2 = levelNumber % 1000;
            if (t1 >= 1 && t2 != 0)
            {
                if (t1 == 1) 
                    result += "одна тысяча";
                else if (t1 == 2)
                    result += "две тысячи";
                else if (t1 == 3 || t1 == 4)
                    result += OneThroughNineteen[t1 - 1] + " тысячи";
                else
                    result += OneThroughNineteen[t1 - 1] + " тысяч";
            }
            if (t1 >= 1 && t2 == 0)
            {
                result += OrdinalOneThroughNineteenHT[t1 - 1] + "тысячный";
                return result.Substring(0, 1).ToUpper() + result.Substring(1);
            }
            if (t1 >= 1)
                result += " ";

            // Get hundreds.
            int h1 = (levelNumber % 1000) / 100;
            int h2 = levelNumber % 100;
            if (h1 >= 1 && h2 != 0)
                result += Hundreds[h1 - 1];
            if (h1 >= 1 && h2 == 0)
            {
                result += (h1 == 1 ? "" : OrdinalOneThroughNineteenHT[h1 - 1]) + "сотый";
                return result.Substring(0, 1).ToUpper() + result.Substring(1);
            }
            if (h1 >= 1)
                result += " ";

            // Get tens and ones.
            int z = levelNumber % 100;
            if (z <= 19)
                result += OrdinalOneThroughNineteen[z - 1];
            else
            {
                int x = z / 10;
                int r = z % 10;
                if (r == 0)
                    result += OrdinalTens[x - 1];
                else
                    result += Tens[x - 1];
                if (r >= 1)
                    result += " " + OrdinalOneThroughNineteen[r - 1];
            }
            return result.Substring(0, 1).ToUpper() + result.Substring(1);
        }
    }
}
