// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OpenXmlPowerTools
{
    public class ListItemTextGetter_de_DE
    {
        private static string[] OneThroughNineteen = {
            "eins", "zwei", "drei", "vier", "fünf", "sechs", "sieben", "acht",
            "neun", "zehn", "elf", "zwölf", "dreizehn", "vierzehn",
            "fünfzehn", "sechzehn", "siebzehn", "achtzehn", "neunzehn"
        };

        private static string[] Tens = {
            "zehn", "zwanzig", "dreißig", "vierzig", "fünfzig", "sechzig", "siebzig",
            "achtzig", "neunzig"
        };
        
        private static string[] OrdinalOneThroughNineteen = {
            "erste", "zweite", "dritte", "vierte", "fünfte", "sechste",
            "siebte", "achte", "neunte", "zehnte", "elfte", "zwölfte",
            "dreizehnte", "vierzehnte", "fünfzehnte", "sechzehnte",
            "siebzehnte", "achtzehnte", "neunzehnte"
        };

        private static string[] OrdinalTens = {
            "zehnte", "zwanzigste", "dreißigste", "vierzigste", "fünfzigste",
            "sechzigste", "siebzigste", "achtzigste", "neunzigste"
        };

        public static string GetListItemText(string languageCultureName, int levelNumber, string numFmt)
        {
			if (levelNumber > 19999)
				throw new ArgumentOutOfRangeException("levelNumber", "Converting a number greater than 19999 to text is not supported");
			if (levelNumber == 0)
				return "Zero";
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
            if (levelNumber % 100 == 11)
                suffix = "-te";
            else if (levelNumber % 10 == 1)
                suffix = "-ste";
            else
                suffix = "-te";
            return levelNumber.ToString() + suffix;
        }

        private static string GetCardinalText(int levelNumber)
        {
            string result = "";

            // Get thousands.
            int t1 = levelNumber / 1000;
            int t2 = levelNumber % 1000;
            if (t1 >= 1)
                result += (t1 == 1 ? "ein" : OneThroughNineteen[t1 - 1]) + "tausend";
            if (t1 >= 1 && t2 == 0)
                return result.Substring(0, 1).ToUpper() + result.Substring(1);
            
            // Get hundreds.
            int h1 = (levelNumber % 1000) / 100;
            int h2 = levelNumber % 100;
            if (h1 >= 1)
                result += (h1 == 1 ? "ein" : OneThroughNineteen[h1 - 1]) + "hundert";
            if (h1 >= 1 && h2 == 0)
                return result.Substring(0, 1).ToUpper() + result.Substring(1);
            
            // Tens and ones.
            int z = levelNumber % 100;
            if (z <= 19)
                result += OneThroughNineteen[z - 1];
            else
            {
                int x = z / 10;
                int r = z % 10;
                if (r >= 1)
                    result += (r == 1 ? "ein" : OneThroughNineteen[r - 1]) + "und";
                result += Tens[x - 1];
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
                result += (t1 == 1 ? "ein" : OneThroughNineteen[t1 - 1]) + "tausend";
            if (t1 >= 1 && t2 == 0)
            {
                result += (t1 == 1 ? "ein" : OneThroughNineteen[t1 - 1]) + "tausendste";
                return result.Substring(0, 1).ToUpper() + result.Substring(1);
            }

            // Get hundreds.
            int h1 = (levelNumber % 1000) / 100;
            int h2 = levelNumber % 100;
            if (h1 >= 1 && h2 != 0)
                result += (h1 == 1 ? "ein" : OneThroughNineteen[h1 - 1]) + "hundert";
            if (h1 >= 1 && h2 == 0)
            {
                result += (h1 == 1 ? "ein" : OneThroughNineteen[h1 - 1])  + "hundertste";
                return result.Substring(0, 1).ToUpper() + result.Substring(1);
            }

            // Get tens and ones.
            int z = levelNumber % 100;
            if (z <= 19)
                result += OrdinalOneThroughNineteen[z - 1];
            else
            {
                int x = z / 10;
                int r = z % 10;
                if (r >= 1)
                    result += (r == 1 ? "ein" : OneThroughNineteen[r - 1]) + "und";
                result += OrdinalTens[x - 1];
            }
            return result.Substring(0, 1).ToUpper() + result.Substring(1);
        }
    }
}
