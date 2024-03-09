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

            // private List<string> colorList = new List<string>{ "Black", "White", "Grey", "Gold", "Brown", "Dark Blue", "Cyan", "Ice", "Red", "Green", "LimeGreen", "Purple", "Pink", "Yellow", "Orange" };

            if (colorString == "Black")
            {
                return "#000000";
            }
            else if (colorString == "White")
            {
                return "#FFFFFF";
            }
            else if (colorString == "Gray")
            {
                return "#808080";
            }
            else if (colorString == "Gold")
            {
                return "#FFD700";
            }
            else if (colorString == "Brown")
            {
                return "#A52A2A";
            }
            else if (colorString == "Blue")
            {
                return "#0000FF";
            }
            else if (colorString == "Cyan")
            {
                return "#00FFFF";
            }
            else if (colorString == "Alice Blue")
            {
                return "#F0F8FF";
            }
            else if (colorString == "Red")
            {
                return "#FF0000";
            }
            else if (colorString == "Green")
            {
                return "#7CFC00";
            }
            else if (colorString == "LimeGreen")
            {
                return "#32CD32";
            }
            else if (colorString == "Purple")
            {
                return "#800080";
            }
            else if (colorString == "Pink")
            {
                return "#FF1493";
            }
            else if (colorString == "Yellow")
            {
                return "#FFFF00";
            }
            else if (colorString == "Orange")
            {
                return "#FFA500";
            }

            return "#000000";
        }
    }
}
