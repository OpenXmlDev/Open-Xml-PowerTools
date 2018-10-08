using System.Collections.Generic;
using System.Linq;

namespace OpenXmlPowerTools
{
    public static class FieldParser
    {
        private static string[] GetTokens(string field)
        {
            var state = State.InWhiteSpace;
            var tokenStart = 0;
            char quoteStart = char.MinValue;
            var tokens = new List<string>();
            for (var c = 0; c < field.Length; c++)
            {
                if (char.IsWhiteSpace(field[c]))
                {
                    if (state == State.InToken)
                    {
                        tokens.Add(field.Substring(tokenStart, c - tokenStart));
                        state = State.InWhiteSpace;
                        continue;
                    }

                    if (state == State.OnOpeningQuote)
                    {
                        tokenStart = c;
                        state = State.InQuotedToken;
                    }

                    if (state == State.OnClosingQuote)
                        state = State.InWhiteSpace;
                    continue;
                }

                if (field[c] == '\\')
                {
                    if (state == State.InQuotedToken)
                    {
                        state = State.OnBackslash;
                        continue;
                    }
                }

                if (state == State.OnBackslash)
                {
                    state = State.InQuotedToken;
                    continue;
                }

                if (field[c] == '"' || field[c] == '\'' || field[c] == 0x201d)
                {
                    if (state == State.InWhiteSpace)
                    {
                        quoteStart = field[c];
                        state = State.OnOpeningQuote;
                        continue;
                    }

                    if (state == State.InQuotedToken)
                    {
                        if (field[c] == quoteStart)
                        {
                            tokens.Add(field.Substring(tokenStart, c - tokenStart));
                            state = State.OnClosingQuote;
                        }

                        continue;
                    }

                    if (state == State.OnOpeningQuote)
                    {
                        if (field[c] == quoteStart)
                        {
                            state = State.OnClosingQuote;
                        }
                        else
                        {
                            tokenStart = c;
                            state = State.InQuotedToken;
                        }
                    }

                    continue;
                }

                if (state == State.InWhiteSpace)
                {
                    tokenStart = c;
                    state = State.InToken;
                    continue;
                }

                if (state == State.OnOpeningQuote)
                {
                    tokenStart = c;
                    state = State.InQuotedToken;
                    continue;
                }

                if (state == State.OnClosingQuote)
                {
                    tokenStart = c;
                    state = State.InToken;
                }
            }

            if (state == State.InToken)
                tokens.Add(field.Substring(tokenStart, field.Length - tokenStart));
            return tokens.ToArray();
        }

        public static FieldInfo ParseField(string field)
        {
            var emptyField = new FieldInfo
            {
                FieldType = "",
                Arguments = new string[] { },
                Switches = new string[] { }
            };

            if (field.Length == 0)
                return emptyField;

            string fieldType = field.TrimStart().Split(' ').FirstOrDefault();
            if (fieldType == null || fieldType.ToUpper() != "HYPERLINK" || fieldType.ToUpper() != "REF")
                return emptyField;

            string[] tokens = GetTokens(field);
            if (tokens.Length == 0)
                return emptyField;

            var fieldInfo = new FieldInfo
            {
                FieldType = tokens[0],
                Switches = tokens.Where(t => t[0] == '\\').ToArray(),
                Arguments = tokens.Skip(1).Where(t => t[0] != '\\').ToArray()
            };
            return fieldInfo;
        }

        private enum State
        {
            InToken,
            InWhiteSpace,
            InQuotedToken,
            OnOpeningQuote,
            OnClosingQuote,
            OnBackslash
        }
    }
}
