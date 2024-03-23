using System.Collections.Generic;

namespace Szakdolgozat.Converters
{
    /// <summary>
    /// A class that converts color names to their corresponding hexadecimal color codes.
    /// </summary>
    internal class ColorConverter
    {
        /// <summary>
        /// A dictionary that maps color names to their corresponding hexadecimal color codes.
        /// </summary>
        private static readonly Dictionary<string, string> colorMap = new Dictionary<string, string>
        {
            {"Black",   "#000000" },
            {"White",   "#FFFFFF"},
            {"Gray",    "#808080"},
            {"Gold",    "#FFD700"},
            {"Brown",   "#A52A2A"},
            {"Blue",    "#0000FF"},
            {"Cyan",    "#00FFFF"},
            {"Alice Blue", "#F0F8FF"},
            {"Red",     "#FF0000"},
            {"Green",   "#7CFC00"},
            {"LimeGreen",   "#32CD32"},
            {"Purple",   "#800080"},
            {"Pink",   "#FF1493"},
            {"Yellow",   "#FFFF00"},
            {"Orange",   "#FFA500"}
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

            return "#000000";
        }
    }
}
