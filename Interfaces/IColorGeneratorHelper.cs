using System.Windows.Media;

namespace Szakdolgozat.Interfaces
{
    internal interface IColorGeneratorHelper
    {
        /// <summary>
        /// Generates a random color by creating a Color object with random RGB values.
        /// Each RGB value is a random byte, which means it can range from 0 to 255.
        /// </summary>
        /// <returns>A Color object with random RGB values.</returns>
        Color GenerateRandomColorForFiles();

        Color GetDisplayColorForFile(string newFileName);
    }
}
