using System;
using System.Linq;
using System.Windows.Media;
using Szakdolgozat.Interfaces;

namespace Szakdolgozat.Helper
{
    /// <summary>
    /// A class that generates random colors.
    /// It uses a Random object to generate random RGB values for the Color object.
    /// </summary>
    internal class ColorGeneratorHelper : IColorGeneratorHelper
    {
        /// <summary>
        /// A Random object used to generate random numbers.
        /// </summary>
        private static readonly Random m_Random = new Random();

        private MainWindow m_MainWindow;

        public ColorGeneratorHelper(MainWindow mainWindow)
        {
            m_MainWindow = mainWindow;
        }

        /// <inheritdoc cref="IColorGeneratorHelper.GenerateRandomColorForFiles"/>
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
