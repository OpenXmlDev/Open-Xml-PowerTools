/***************************************************************************

Copyright (c) EricWhite.com 2018.

Developer: Eric White
Blog: http://www.ericwhite.com
Twitter: @EricWhiteDev
Email: eric@ericwhite.com

***************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public class PtpSHA1Util
    {
        public static string SHA1HashStringForUTF8String(string s)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(s);
            var sha1 = SHA1.Create();
            byte[] hashBytes = sha1.ComputeHash(bytes);
            return HexStringFromBytes(hashBytes);
        }

        public static string SHA1HashStringForByteArray(byte[] bytes)
        {
            var sha1 = SHA1.Create();
            byte[] hashBytes = sha1.ComputeHash(bytes);
            return HexStringFromBytes(hashBytes);
        }

        public static string HexStringFromBytes(byte[] bytes)
        {
            var sb = new StringBuilder();
            foreach (byte b in bytes)
            {
                var hex = b.ToString("x2");
                sb.Append(hex);
            }
            return sb.ToString();
        }
    }

    public class Base64Util
    {
        private class Bs64Tupple
        {
            public char Bs64Character;
            public int Bs64Chunk;
        }

        public static string Convert76CharLineLength(byte[] byteArray)
        {
            string base64String = (System.Convert.ToBase64String(byteArray))
                .Select
                (
                    (c, i) => new Bs64Tupple()
                    {
                        Bs64Character = c,
                        Bs64Chunk = i / 76
                    }
                )
                .GroupBy(c => c.Bs64Chunk)
                .Aggregate(
                    new StringBuilder(),
                    (s, i) =>
                        s.Append(
                            i.Aggregate(
                                new StringBuilder(),
                                (seed, it) => seed.Append(it.Bs64Character),
                                sb => sb.ToString()
                            )
                        )
                        .Append(Environment.NewLine),
                    s => s.ToString()
                );
            return base64String;
        }
    }
}
