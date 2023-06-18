﻿// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OpenXmlPowerTools
{
    public class ListItemTextGetter_tr_TR
    {
        private static string[] RomanOnes =
        {
            "", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX"
        };

        private static string[] RomanTens =
        {
            "", "X", "XX", "XXX", "XL", "L", "LX", "LXX", "LXXX", "XC"
        };

        private static string[] RomanHundreds =
        {
            "", "C", "CC", "CCC", "CD", "D", "DC", "DCC", "DCCC", "CM", "M"
        };

        private static string[] RomanThousands =
        {
            "", "M", "MM", "MMM", "MMMM", "MMMMM", "MMMMMM", "MMMMMMM", "MMMMMMMM",
            "MMMMMMMMM", "MMMMMMMMMM"
        };

        private static string[] OneThroughNineteen = {
            "bir", "iki", "üç", "dört", "beş", "altı", "yedi", "sekiz",
            "dokuz", "on", "onbir", "oniki", "onüç", "ondört",
            "onbeş", "onaltı", "onyedi", "onsekiz", "ondokuz"
        };

        private static string[] Tens = {
            "on", "yirmi", "otuz", "kırk", "elli", "altmış", "yetmiş",
            "seksen", "doksan"
        };

        private static string[] OrdinalOneThroughNineteen = {
            "birinci", "ikinci", "üçüncü", "dördüncü", "beşinci", "altıncı",
            "yedinci", "sekizinci", "dokuzuncu", "onuncu", "onbirinci", "onikinci",
            "onüçüncü", "ondördüncü", "onbeşinci", "onaltıncı",
            "onyedinci", "onsekizinci", "ondokuzuncu"
        };

        private static string[] TwoThroughNineteen = {
            "", "iki", "üç", "dört", "beş", "altı", "yedi", "sekiz",
            "dokuz", "on", "onbir", "oniki", "onüç", "ondört",
            "onbeş", "onaltı", "onyedi", "onsekiz", "ondokuz"
        };

        private static string[] OrdinalTenths = {
            "onuncu", "yirminci", "otuzuncu", "kırkıncı", "ellinci",
            "altmışıncı", "yetmişinci", "sekseninci", "doksanıncı"
        };

        public static string GetListItemText(string languageCultureName, int levelNumber, string numFmt)
        {
            #region
            if (numFmt == "decimal")
            {
                return levelNumber.ToString();
            }
            if (numFmt == "decimalZero")
            {
                if (levelNumber <= 9)
                    return "0" + levelNumber.ToString();
                else
                    return levelNumber.ToString();
            }
            if (numFmt == "upperRoman")
            {
                int ones = levelNumber % 10;
                int tens = (levelNumber % 100) / 10;
                int hundreds = (levelNumber % 1000) / 100;
                int thousands = levelNumber / 1000;
                return RomanThousands[thousands] + RomanHundreds[hundreds] +
                    RomanTens[tens] + RomanOnes[ones];
            }
            if (numFmt == "lowerRoman")
            {
                int ones = levelNumber % 10;
                int tens = (levelNumber % 100) / 10;
                int hundreds = (levelNumber % 1000) / 100;
                int thousands = levelNumber / 1000;
                return (RomanThousands[thousands] + RomanHundreds[hundreds] +
                    RomanTens[tens] + RomanOnes[ones]).ToLower();
            }
            if (numFmt == "upperLetter")
            {
                string a = "ABCÇDEFGĞHIİJKLMNOÖPRSŞTUÜVYZ";
                //string a = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
                int c = (levelNumber - 1) / 29;
                int n = (levelNumber - 1) % 29;
                char x = a[n];
                return "".PadRight(c + 1, x);
            }
            if (numFmt == "lowerLetter")
            {
                string a = "abcçdefgğhıijklmnoöprsştuüvyz";
                int c = (levelNumber - 1) / 29;
                int n = (levelNumber - 1) % 29;
                char x = a[n];
                return "".PadRight(c + 1, x);
            }
            if (numFmt == "ordinal")
            {
                string suffix;
                /*if (levelNumber % 100 == 11 || levelNumber % 100 == 12 ||
                    levelNumber % 100 == 13)
                    suffix = "th";
                else if (levelNumber % 10 == 1)
                    suffix = "st";
                else if (levelNumber % 10 == 2)
                    suffix = "nd";
                else if (levelNumber % 10 == 3)
                    suffix = "rd";
                else
                    suffix = "th";*/
                suffix = ".";
                return levelNumber.ToString() + suffix;
            }
            if (numFmt == "cardinalText")
            {
                string result = "";
                int t1 = levelNumber / 1000;
                int t2 = levelNumber % 1000;
                if (t1 >= 1)
                    result += OneThroughNineteen[t1 - 1] + " yüz";
                if (t1 >= 1 && t2 == 0)
                    return result.Substring(0, 1).ToUpper() +
                        result.Substring(1);
                if (t1 >= 1)
                    result += " ";
                int h1 = (levelNumber % 1000) / 100;
                int h2 = levelNumber % 100;
                if (h1 >= 1)
                    result += OneThroughNineteen[h1 - 1] + " bin";
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
                        result += /*"-" + */OneThroughNineteen[r - 1];
                }
                return result.Substring(0, 1).ToUpper() +
                    result.Substring(1);
            }
            #endregion
            if (numFmt == "ordinalText")
            {
                string result = "";
                int t1 = levelNumber / 1000;
                int t2 = levelNumber % 1000;
                if (t1 >= 1 && t2 != 0)
                    result += TwoThroughNineteen[t1 - 1] + "bin";
                if (t1 >= 1 && t2 == 0)
                {
                    result += TwoThroughNineteen[t1 - 1] + "bininci";
                    return result.Substring(0, 1).ToUpper() +
                        result.Substring(1);
                }
                //if (t1 >= 1)
                //    result += " ";
                int h1 = (levelNumber % 1000) / 100;
                int h2 = levelNumber % 100;
                if (h1 >= 1 && h2 != 0)
                    result += TwoThroughNineteen[h1 - 1] + "yüz";
                if (h1 >= 1 && h2 == 0)
                {
                    result += TwoThroughNineteen[h1 - 1] + "yüzüncü";
                    return result.Substring(0, 1).ToUpper() +
                        result.Substring(1);
                }
                //if (h1 >= 1)
                //    result += " ";
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
                        result += OrdinalOneThroughNineteen[r - 1]; //result += "-" + OrdinalOneThroughNineteen[r - 1];
                }
                return result.Substring(0, 1).ToUpper() +
                    result.Substring(1);
            }
            if (numFmt == "0001, 0002, 0003, ...")
            {
                return string.Format("{0:0000}", levelNumber);
            }
            if (numFmt == "bullet")
                return "";
            return levelNumber.ToString();
        }
    }
}
