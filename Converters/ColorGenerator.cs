using System;
using System.Linq;
using System.Windows.Media;

namespace Szakdolgozat.Converters
{
    /// <summary>
    /// A class that generates random colors.
    /// It uses a Random object to generate random RGB values for the Color object.
    /// </summary>
    internal class ColorGenerator
    {
        /// <summary>
        /// A Random object used to generate random numbers.
        /// </summary>
        private static readonly Random m_Random = new Random();

        private MainWindow m_MainWindow;

        public ColorGenerator(MainWindow mainWindow)
        {
            m_MainWindow = mainWindow;
        }

        /// <summary>
        /// Generates a random color by creating a Color object with random RGB values.
        /// Each RGB value is a random byte, which means it can range from 0 to 255.
        /// </summary>
        /// <returns>A Color object with random RGB values.</returns>
        public Color GenerateRandomColorForFiles()
        {
            return Color.FromRgb(
                (byte)m_Random.Next(256),
                (byte)m_Random.Next(256),
                (byte)m_Random.Next(256));
        }

        public Color GetDisplayColorForFile(string newFileName)
        {
            return m_MainWindow.SelectedFiles.FirstOrDefault(file => file.FileName == newFileName)?.DisplayColor
                   ?? GenerateRandomColorForFiles();
        }
    }
}
