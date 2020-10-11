using System;
using System.IO;
using System.Linq;
using System.Text;

namespace OpenXmlPowerTools
{
    public class Base64
    {
        public static string ConvertToBase64(string fileName)
        {
            var ba = File.ReadAllBytes(fileName);
            var base64String = (Convert.ToBase64String(ba))
                .Select
                (
                    (c, i) => new
                    {
                        Chunk = i / 76,
                        Character = c
                    }
                )
                .GroupBy(c => c.Chunk)
                .Aggregate(
                    new StringBuilder(),
                    (s, i) =>
                        s.Append(
                            i.Aggregate(
                                new StringBuilder(),
                                (seed, it) => seed.Append(it.Character),
                                sb => sb.ToString()
                            )
                        )
                        .Append(Environment.NewLine),
                    s =>
                    {
                        s.Length -= Environment.NewLine.Length;
                        return s.ToString();
                    }
                );

            return base64String;
        }

        public static byte[] ConvertFromBase64(string b64)
        {
            var b64b = b64.Replace("\r\n", "");
            var ba = Convert.FromBase64String(b64b);
            return ba;
        }
    }
}