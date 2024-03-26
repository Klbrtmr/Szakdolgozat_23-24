using System;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Szakdolgozat.Interfaces;
using Szakdolgozat.Model;

namespace Szakdolgozat.View
{
    /// <summary>
    /// A helper class that provides methods to create and update UI elements in the MainWindow.
    /// It includes methods to create a StackPanel for a file, create an Image from an image path,
    /// and set the icon images for the MainWindow's buttons.
    /// </summary>
    internal class UIHelper : IUIHelper
    {
        /// <summary>
        /// The MainWindow instance whose UI elements are to be created or updated.
        /// This instance is used to access the UI elements that need to be created or updated.
        /// </summary>
        private MainWindow m_MainWindow;

        /// <summary>
        /// Initializes a new instance of the UIHelper class.
        /// </summary>
        /// <param name="mainWindow">The MainWindow whose UI elements are to be created or updated.</param>
        public UIHelper(MainWindow mainWindow)
        {
            m_MainWindow = mainWindow;
        }

        /// <inheritdoc cref="IUIHelper.CreateFilePanel"/>
        public StackPanel CreateFilePanel(ImportedFile importedFile)
        {
            Ellipse ellipse = new Ellipse();
            ellipse.Width = 15;
            ellipse.Height = 15;

            SolidColorBrush brush = new SolidColorBrush(importedFile.DisplayColor);

            ellipse.Fill = brush;

            StackPanel panel = new StackPanel();
            panel.Orientation = Orientation.Horizontal;
            panel.Children.Add(ellipse);
            panel.Children.Add(new TextBlock() { Text = importedFile.FileName });

            return panel;
        }

        /// <inheritdoc cref="IUIHelper.CreateImage"/>
        public Image CreateImage(string imagePath, int width = 16, int height = 16)
        {
            return new Image
            {
                Source = new BitmapImage(new Uri(imagePath, UriKind.RelativeOrAbsolute)),
                Width = width,
                Height = height
            };
        }

        /// <inheritdoc cref="IUIHelper.SetIconImagesForDarkMode"/>
        public void SetIconImagesForDarkMode()
        {
            SetIconImages("light");
        }

        /// <inheritdoc cref="IUIHelper.SetIconImages"/>
        public void SetIconImagesForLightMode()
        {
            SetIconImages("");
        }

        /// <summary>
        /// Sets the icon images for the MainWindow's buttons.
        /// The mode parameter determines the suffix of the image file names.
        /// </summary>
        /// <param name="mode">The mode that determines the suffix of the image file names.</param>
        private void SetIconImages(string mode)
        {
            string suffix = string.IsNullOrEmpty(mode) ? "" : $"_{mode}";

            m_MainWindow.cleanButton.Icon = CreateImage($"Assets/broom{suffix}.png");
            m_MainWindow.importExcel.Icon = CreateImage($"Assets/document{suffix}.png");
            m_MainWindow.importProject.Icon = CreateImage($"Assets/file-import{suffix}.png");
            m_MainWindow.exporttoexcel.Icon = CreateImage($"Assets/export{suffix}.png");
            m_MainWindow.exportproject.Icon = CreateImage($"Assets/exportfile{suffix}.png");
            m_MainWindow.closeProject.Icon = CreateImage($"Assets/exit{suffix}.png");
        }
    }
}
