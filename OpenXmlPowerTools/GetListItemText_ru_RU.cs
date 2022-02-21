namespace Codeuctivity.OpenXmlPowerTools
{
    public class ListItemTextGetterRuRU
    {
        private static readonly string[] OneThroughNineteen = {
            "один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь",
            "девять", "десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать",
            "пятнадцать", "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать"
        };

        private static readonly string[] Tens = {
            "десять", "двадцать", "тридцать", "сорок", "пятьдесят", "шестьдесят", "семьдесят",
            "восемьдесят", "девяносто"
        };

        private static readonly string[] OrdinalOneThroughNineteen = {
            "первый", "второй", "третий", "четвертый", "пятый", "шестой",
            "седьмой", "восьмой", "девятый", "десятый", "одиннадцатый", "двенадцатый",
            "тринадцатый", "четырнадцатый", "пятнадцатый", "шестнадцатый",
            "семнадцатый", "восемнадцатый", "девятнадцатый"
        };

        private static readonly string[] OrdinalTenths = {
            "десятый", "двадцатый", "тридцатый", "сороковой", "пятидесятый",
            "шестидесятый", "семидесятый", "восьмидесятый", "девяностый"
        };

        // TODO this is not correct for values above 99

        public static string? GetListItemText(int levelNumber, string numFmt)
        {
            if (numFmt == "cardinalText")
            {
                var result = "";
                var t1 = levelNumber / 1000;
                var t2 = levelNumber % 1000;
                if (t1 >= 1)
                {
                    result += OneThroughNineteen[t1 - 1] + " thousand";
                }

                if (t1 >= 1 && t2 == 0)
                {
                    return result.Substring(0, 1).ToUpper() +
                        result.Substring(1);
                }

                if (t1 >= 1)
                {
                    result += " ";
                }

                var h1 = levelNumber % 1000 / 100;
                var h2 = levelNumber % 100;
                if (h1 >= 1)
                {
                    result += OneThroughNineteen[h1 - 1] + " hundred";
                }

                if (h1 >= 1 && h2 == 0)
                {
                    return result.Substring(0, 1).ToUpper() +
                        result.Substring(1);
                }

                if (h1 >= 1)
                {
                    result += " ";
                }

                var z = levelNumber % 100;
                if (z <= 19)
                {
                    result += OneThroughNineteen[z - 1];
                }
                else
                {
                    var x = z / 10;
                    var r = z % 10;
                    result += Tens[x - 1];
                    if (r >= 1)
                    {
                        result += "-" + OneThroughNineteen[r - 1];
                    }
                }
                return result.Substring(0, 1).ToUpper() +
                    result.Substring(1);
            }
            if (numFmt == "ordinalText")
            {
                var result = "";
                var t1 = levelNumber / 1000;
                var t2 = levelNumber % 1000;
                if (t1 >= 1 && t2 != 0)
                {
                    result += OneThroughNineteen[t1 - 1] + " thousand";
                }

                if (t1 >= 1 && t2 == 0)
                {
                    result += OneThroughNineteen[t1 - 1] + " thousandth";
                    return result.Substring(0, 1).ToUpper() +
                        result.Substring(1);
                }
                if (t1 >= 1)
                {
                    result += " ";
                }

                var h1 = levelNumber % 1000 / 100;
                var h2 = levelNumber % 100;
                if (h1 >= 1 && h2 != 0)
                {
                    result += OneThroughNineteen[h1 - 1] + " hundred";
                }

                if (h1 >= 1 && h2 == 0)
                {
                    result += OneThroughNineteen[h1 - 1] + " hundredth";
                    return result.Substring(0, 1).ToUpper() +
                        result.Substring(1);
                }
                if (h1 >= 1)
                {
                    result += " ";
                }

                var z = levelNumber % 100;
                if (z <= 19)
                {
                    result += OrdinalOneThroughNineteen[z - 1];
                }
                else
                {
                    var x = z / 10;
                    var r = z % 10;
                    if (r == 0)
                    {
                        result += OrdinalTenths[x - 1];
                    }
                    else
                    {
                        result += Tens[x - 1];
                    }

                    if (r >= 1)
                    {
                        result += " " + OrdinalOneThroughNineteen[r - 1];
                    }
                }
                return result.Substring(0, 1).ToUpper() +
                    result.Substring(1);
            }
            return null;
        }
    }
}