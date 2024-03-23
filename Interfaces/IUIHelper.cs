using System.Windows.Controls;
using Szakdolgozat.Model;

namespace Szakdolgozat.Interfaces
{
    internal interface IUIHelper
    {
        /// <summary>
        /// Creates a StackPanel for a file.
        /// The StackPanel includes an Ellipse filled with the file's display color,
        /// and a TextBlock with the file's name.
        /// </summary>
        /// <param name="importedFile">The file for which to create the StackPanel.</param>
        /// <returns>The created StackPanel.</returns>
        StackPanel CreateFilePanel(ImportedFile importedFile);

        /// <summary>
        /// Creates an Image from an image path.
        /// </summary>
        /// <param name="imagePath">The path of the image</param>
        /// <param name="width">The width of the image</param>
        /// <param name="height">The height of the image</param>
        /// <returns></returns>
        Image CreateImage(string imagePath, int width = 16, int height = 16);

        /// <summary>
        /// Sets the icon images for the MainWindow's buttons for dark mode.
        /// </summary>
        void SetIconImagesForDarkMode();

        /// <summary>
        /// Sets the default icon images for the MainWindow's buttons.
        /// </summary>
        void SetIconImagesForLightMode();
    }
}
