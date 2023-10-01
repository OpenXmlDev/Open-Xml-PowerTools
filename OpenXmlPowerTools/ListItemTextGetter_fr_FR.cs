﻿// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OpenXmlPowerTools
{
    public class ListItemTextGetter_fr_FR
    {
        private static string[] OneThroughNineteen = {
            "", "un", "deux", "trois", "quatre", "cinq", "six", "sept", "huit",
            "neuf", "dix", "onze", "douze", "treize", "quatorze",
            "quinze", "seize", "dix-sept", "dix-huit", "dix-neuf"
        };

        private static string[] Tens = {
            "", "dix", "vingt", "trente", "quarante", "cinquante", "soixante", "soixante-dix",
            "quatre-vingt", "quatre-vingt-dix"
        };

        private static string[] OrdinalOneThroughNineteen = {
            "", "unième", "deuxième", "troisième", "quatrième", "cinquième", "sixième",
            "septième", "huitième", "neuvième", "dixième", "onzième", "douzième",
            "treizième", "quatorzième", "quinzième", "seizième",
            "dix-septième", "dix-huitième", "dix-neuvième"
        };

        private static string[] OrdinalTenths = {
            "", "dixième", "vingtième", "trentième", "quarantième", "cinquantième",
            "soixantième", "soixante-dixième", "quatre-vingtième", "quatre-vingt-dixième"
        };

        private static string[] OrdinalTenthsPlus = {
            "", "", "vingt", "trente", "quarante", "cinquante",
            "soixante", "", "quatre-vingt", ""
        };

        public static string GetListItemText(string languageCultureName, int levelNumber, string numFmt)
        {
            if (numFmt == "cardinalText")
            {
                string result = "";

                var sLevel = (levelNumber + 10000).ToString();
                int thousands = int.Parse(sLevel.Substring(1, 1));
                int hundreds = int.Parse(sLevel.Substring(2, 1));
                int tens = int.Parse(sLevel.Substring(3, 1));
                int ones = int.Parse(sLevel.Substring(4, 1));

                /* exact thousands */
                if (levelNumber == 1000)
                    return "Mille";
                if (levelNumber > 1000 && hundreds == 0 && tens == 0 && ones == 0)
                {
                    result = OneThroughNineteen[thousands] + " mille";
                    return result.Substring(0, 1).ToUpper() + result.Substring(1);
                }

                /* > 1000 */
                if (levelNumber > 1000 && levelNumber < 2000)
                    result = "mille ";
                else if (levelNumber > 2000 && levelNumber < 10000)
                    result = OneThroughNineteen[thousands] + " mille ";

                /* exact hundreds */
                if (hundreds > 0 && tens == 0 && ones == 0)
                {
                    if (hundreds == 1)
                        result += "cent";
                    else
                        result += OneThroughNineteen[hundreds] + " cents";
                    return result.Substring(0, 1).ToUpper() + result.Substring(1);
                }

                /* > 100 */
                if (hundreds > 0)
                {
                    if (hundreds == 1)
                        result += "cent ";
                    else
                        result += OneThroughNineteen[hundreds] + " cent ";
                }

                /* exact tens */
                if (tens > 0 && ones == 0)
                {
                    if (tens == 8)
                        result += "quatre-vingts";
                    else
                        result += Tens[tens];
                    return result.Substring(0, 1).ToUpper() + result.Substring(1);
                }

                /* 71-79 */
                if (tens == 7)
                {
                    if (ones == 1)
                        result += Tens[6] + " et " + OneThroughNineteen[ones + 10];
                    else
                        result += Tens[6] + "-" + OneThroughNineteen[ones + 10];
                    return result.Substring(0, 1).ToUpper() + result.Substring(1);
                }

                /* 91-99 */
                if (tens == 9)
                {
                    result += Tens[8] + "-" + OneThroughNineteen[ones + 10];
                    return result.Substring(0, 1).ToUpper() + result.Substring(1);
                }

                if (tens >= 2)
                {
                    if (ones == 1 && tens != 8 && tens != 9)
                        result += Tens[tens] + " et un";
                    else
                        result += Tens[tens] + "-" + OneThroughNineteen[ones];
                    return result.Substring(0, 1).ToUpper() + result.Substring(1);
                }

                if (levelNumber < 20)
                {
                    result += OneThroughNineteen[tens * 10 + ones];
                    return result.Substring(0, 1).ToUpper() + result.Substring(1);
                }

                result += OneThroughNineteen[tens * 10 + ones];
                return result.Substring(0, 1).ToUpper() + result.Substring(1);
            }
            if (numFmt == "ordinal")
            {
                string suffix;
                if (levelNumber == 1)
                    suffix = "er";
                else
                    suffix = "e";
                return levelNumber.ToString() + suffix;
            }
            if (numFmt == "ordinalText")
            {
                string result = "";

                if (levelNumber == 1)
                    return "Premier";

                var sLevel = (levelNumber + 10000).ToString();
                int thousands = int.Parse(sLevel.Substring(1, 1));
                int hundreds = int.Parse(sLevel.Substring(2, 1));
                int tens = int.Parse(sLevel.Substring(3, 1));
                int ones = int.Parse(sLevel.Substring(4, 1));

                /* exact thousands */
                if (levelNumber == 1000)
                    return "Millième";
                if (levelNumber > 1000 && hundreds == 0 && tens == 0 && ones == 0)
                {
                    result = OneThroughNineteen[thousands] + " millième";
                    return result.Substring(0, 1).ToUpper() + result.Substring(1);
                }

                /* > 1000 */
                if (levelNumber > 1000 && levelNumber < 2000)
                    result = "mille ";
                else if (levelNumber > 2000 && levelNumber < 10000)
                    result = OneThroughNineteen[thousands] + " mille ";

                /* exact hundreds */
                if (hundreds > 0 && tens == 0 && ones == 0)
                {
                    if (hundreds == 1)
                        result += "centième";
                    else
                        result += OneThroughNineteen[hundreds] + " centième";
                    return result.Substring(0, 1).ToUpper() + result.Substring(1);
                }

                /* > 100 */
                if (hundreds > 0)
                {
                    if (hundreds == 1)
                        result += "cent ";
                    else
                        result += OneThroughNineteen[hundreds] + " cent ";
                }

                /* exact tens */
                if (tens > 0 && ones == 0)
                {
                    result += OrdinalTenths[tens];
                    return result.Substring(0, 1).ToUpper() + result.Substring(1);
                }

                /* 71-79 */
                if (tens == 7)
                {
                    if (ones == 1)
                        result += OrdinalTenthsPlus[6] + " et " + OrdinalOneThroughNineteen[ones + 10];
                    else
                        result += OrdinalTenthsPlus[6] + "-" + OrdinalOneThroughNineteen[ones + 10];
                    return result.Substring(0, 1).ToUpper() + result.Substring(1);
                }

                /* 91-99 */
                if (tens == 9)
                {
                    result += OrdinalTenthsPlus[8] + "-" + OrdinalOneThroughNineteen[ones + 10];
                    return result.Substring(0, 1).ToUpper() + result.Substring(1);
                }

                if (tens >= 2)
                {
                    if (ones == 1 && tens != 8 && tens != 9)
                        result += OrdinalTenthsPlus[tens] + " et unième";
                    else
                        result += OrdinalTenthsPlus[tens] + "-" + OrdinalOneThroughNineteen[ones];
                    return result.Substring(0, 1).ToUpper() + result.Substring(1);
                }

                if (levelNumber < 20)
                {
                    result += OrdinalOneThroughNineteen[tens * 10 + ones];
                    return result.Substring(0, 1).ToUpper() + result.Substring(1);
                }

                result += OrdinalOneThroughNineteen[tens * 10 + ones];
                return result.Substring(0, 1).ToUpper() + result.Substring(1);
            }
            return null;
        }
    }
}
