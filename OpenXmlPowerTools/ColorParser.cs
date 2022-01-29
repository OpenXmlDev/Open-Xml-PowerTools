using SixLabors.ImageSharp;

namespace Codeuctivity.OpenXmlPowerTools
{
    public static class ColorParser
    {
        public static Color FromName(string name)
        {
            return Color.Parse(name);
        }

        public static bool TryFromName(string name, out Color color)
        {
            try
            {
                color = Color.Parse(name);

                return true;
            }
            catch
            {
                color = default;

                return false;
            }
        }

        public static bool IsValidName(string name)
        {
            return TryFromName(name, out _);
        }
    }
}