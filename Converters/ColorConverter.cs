using System.Collections.Generic;
using Szakdolgozat.Properties;

namespace Szakdolgozat.Converters
{
    /// <summary>
    /// A class that converts color names to their corresponding hexadecimal color codes.
    /// </summary>
    internal class ColorConverter
    {
        /// <summary>
        /// Color list for configuration page.
        /// </summary>
        public List<string> colorList = new List<string>{
            Resources.Black, Resources.White, Resources.Gray, Resources.Gold, Resources.Brown,
            Resources.Blue, Resources.Cyan, Resources.AliceBlue, Resources.Red, Resources.Green,
            Resources.LimeGreen, Resources.Purple, Resources.Pink, Resources.Yellow, Resources.Orange };


        /// <summary>
        /// A dictionary that maps color names to their corresponding hexadecimal color codes.
        /// </summary>
        private static readonly Dictionary<string, string> colorMap = new Dictionary<string, string>
        {
            {Resources.Black,       Resources.BlackHexValue},
            {Resources.White,       Resources.WhiteHexValue},
            {Resources.Gray,        Resources.GrayHexValue},
            {Resources.Gold,        Resources.GoldHexValue},
            {Resources.Brown,       Resources.BrownHexValue},
            {Resources.Blue,        Resources.BlueHexValue},
            {Resources.Cyan,        Resources.CyanHexValue},
            {Resources.AliceBlue,   Resources.AliceBlueHexValue},
            {Resources.Red,         Resources.RedHexValue},
            {Resources.Green,       Resources.GreenHexValue},
            {Resources.LimeGreen,   Resources.LimeGreenHexValue},
            {Resources.Purple,      Resources.PurpleHexValue},
            {Resources.Pink,        Resources.PinkHexValue},
            {Resources.Yellow,      Resources.YellowHexValue},
            {Resources.Orange,      Resources.OrangeHexValue}
        };

        /// <summary>
        /// Converts a color name to its corresponding hexadecimal color code.
        /// If the color code name is not found in the colorMap dictionary, it returns black ("#000000").
        /// </summary>
        /// <param name="colorObject">The color name as an object.</param>
        /// <returns>The hexadecimal color code as a string.</returns>
        public string ColorNameToHex(object colorObject)
        {
            if (colorObject is string colorString &&
                colorMap.TryGetValue(colorString, out string hexValue))
            {
                return hexValue;
            }

            return Resources.BlackHexValue;
        }
    }
}
