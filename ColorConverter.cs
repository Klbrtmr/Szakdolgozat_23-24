using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Szakdolgozat
{
    internal class ColorConverter
    {
        public string ConvertFromString(object colorObject)
        {

            string colorString = colorObject.ToString();

            // private List<string> colorList = new List<string>{ "Black", "White", "Grey", "Blond", "Brown", "Dark Blue", "Cyan", "Ice, "Red", "Green", "LimeGreen", "Purple", "Pink", "Yellow" };

            if (colorString == "Black")
            {
                return "#000000";
            }
            else if (colorString == "White")
            {
                return "#ffffff";
            }
            else if (colorString == "Grey")
            {
                return "#333333";
            }
            else if (colorString == "Blond")
            {
                return "#faf0be";
            }
            else if (colorString == "Brown")
            {
                return "#964b00";
            }
            else if (colorString == "Dark Blue")
            {
                return "#0000ff";
            }
            else if (colorString == "Cyan")
            {
                return "#00ffff";
            }
            else if (colorString == "Ice")
            {
                return "#c6e2ff";
            }
            else if (colorString == "Red")
            {
                return "#ff0000";
            }
            else if (colorString == "Green")
            {
                return "#00ff00";
            }
            else if (colorString == "LimeGreen")
            {
                return "#32cd32";
            }
            else if (colorString == "Purple")
            {
                return "#a020f0";
            }
            else if (colorString == "Pink")
            {
                return "#ff6ec7";
            }
            else if (colorString == "Yellow")
            {
                return "#ffff00";
            }
            else if (colorString == "Orange")
            {
                return "#ff9500";
            }

            return "#000000";
        }
    }
}
