/***************************************************************************

Copyright (c) Microsoft Corporation 2012-2015.

This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license can be found here:

http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx

Published at http://OpenXmlDeveloper.org
Resource Center and Documentation: http://openxmldeveloper.org/wiki/w/wiki/powertools-for-open-xml.aspx

Developer: Eric White
Blog: http://www.ericwhite.com
Twitter: @EricWhiteDev
Email: eric@ericwhite.com

***************************************************************************/

using System.Drawing;

namespace OpenXmlPowerTools
{
    public static class ColorParser
    {
        public static Color FromName(string name)
        {
            return Color.FromName(name);
        }

        public static bool TryFromName(string name, out Color color)
        {
            try
            {
                color = Color.FromName(name);

                return color.IsNamedColor;
            }
            catch
            {
                color = default(Color);

                return false;
            }
        }

        public static bool IsValidName(string name)
        {
            return TryFromName(name, out _);
        }
    }
}
