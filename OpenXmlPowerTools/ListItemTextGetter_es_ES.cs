// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OpenXmlPowerTools
{
    public class ListItemTextGetter_es_ES
    {
        private static string[] OneThroughNineteen = {
            "uno", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho",
            "nueve", "diez", "once", "doce", "trece", "catorce",
            "quince", "dieciséis", "diecisiete", "dieciocho", "diecinueve"
        };

        private static string[] Tens = {
            "diez", "veinte", "treinta", "cuarenta", "cincuenta", "sesenta", "setenta",
            "ochenta", "noventa"
        };

        private static string[] Hundreds = {
            "cien", "doscientos", "trescientos", "cuatrocientos", "quinientos", "seiscientos", "setecientos",
            "ochocientos", "novecientos"
        };

        private static string[] OrdinalOneThroughNineteen = {
            "primero", "segundo", "tercero", "cuarto", "quinto", "sexto",
            "séptimo", "octavo", "noveno", "décimo", "undécimo", "duodécimo",
            "decimotercio", "decimocuarto", "decimoquinto", "decimosexto",
            "decimoséptimo", "decimoctavo", "decimonoveno"
        };

        private static string[] OrdinalTens = {
            "décimo", "vigésimo", "trigésimo", "cuadragésimo", "quincuagésimo",
            "sexagésimo", "septuagésimo", "octogésimo", "nonagésimo"
        };
        
        private static string[] OrdinalHundreds = {
            "centésimo", "ducentesimo", "tricentesimo", "cuadringentesimo", "quingentesimo", "sexcentesimo", "septingentesimo",
            "octingentesimo", "noningentesimo"
        };

        public static string GetListItemText(string languageCultureName, int levelNumber, string numFmt)
        {
			if (levelNumber > 19999)
				throw new ArgumentOutOfRangeException("levelNumber", "Converting a number greater than 19999 to text is not supported");
			if (levelNumber == 0)
				return "Cero";
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
            return levelNumber.ToString() + "-o";
        }

        private static string GetCardinalText(int levelNumber)
        {
            string result = "";

            // Get thousands.
            int t1 = levelNumber / 1000;
            int t2 = levelNumber % 1000;
            if (t1 >= 1)
                result += (t1 == 1 ? "" : OneThroughNineteen[t1 - 1] + " ") + "mil";
            if (t1 >= 1 && t2 == 0)
                return result.Substring(0, 1).ToUpper() + result.Substring(1);
            if (t1 >= 1)
                result += " ";
            
            // Get hundreds.
            int h1 = (levelNumber % 1000) / 100;
            int h2 = levelNumber % 100;
            if (h1 >= 1)
                result += h1 == 1 && h2 != 0 ? "ciento" : Hundreds[h1 - 1];
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
                    result += " y " + OneThroughNineteen[r - 1];
            }
            return result.Substring(0, 1).ToUpper() + result.Substring(1);
        }

        private static string GetOrdinalText(int levelNumber)
        {
            string result = "";

            // Get thousands.
            int t1 = levelNumber / 1000;
            int t2 = levelNumber % 1000;
            if (t1 >= 1)
                result += (t1 == 1 ? "" : OneThroughNineteen[t1 - 1] + " ") + "milesimo";
            if (t1 >= 1 && t2 == 0)
                return result.Substring(0, 1).ToUpper() + result.Substring(1);
            if (t1 >= 1)
                result += " ";

            // Get hundreds.
            int h1 = (levelNumber % 1000) / 100;
            int h2 = levelNumber % 100;
            if (h1 >= 1)
                result += OrdinalHundreds[h1 - 1];
            if (h1 >= 1 && h2 == 0)
                return result.Substring(0, 1).ToUpper() + result.Substring(1);
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
                result += OrdinalTens[x - 1];
                if (r >= 1)
                    result += " " + OrdinalOneThroughNineteen[r - 1];
            }
            return result.Substring(0, 1).ToUpper() + result.Substring(1);
        }
    }
}
