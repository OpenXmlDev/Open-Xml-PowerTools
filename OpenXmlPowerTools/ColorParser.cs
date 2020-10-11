

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