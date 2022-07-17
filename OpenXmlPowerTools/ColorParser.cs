using SixLabors.ImageSharp;

namespace Codeuctivity.OpenXmlPowerTools
{
    public static class ColorParser
    {
        public static Color FromName(string name)
        {
            return Color.Parse(name);
        }

        public static bool TryFromName(string? name, out Color color)
        {
            return Color.TryParse(name, out color);
        }

        public static bool IsValidName(string name)
        {
            return TryFromName(name, out _);
        }
    }
}