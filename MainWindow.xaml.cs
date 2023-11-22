using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Szakdolgozat
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// The list in which the imported files are stored.
        /// </summary>
        private List<ImportedFile> selectedFiles = new List<ImportedFile>();
        
        /// <summary>
        /// Listed all file what we imported. This method created an ellipse to every file.
        /// </summary>
        private void ListFiles()
        {
            filesListing.ItemsSource = null;
            filesListing.Items.Clear();

            foreach (var importedFile in selectedFiles)
            {
                Ellipse ellipse = new Ellipse();
                ellipse.Width = 15;
                ellipse.Height = 15;

                SolidColorBrush brush = new SolidColorBrush(importedFile.DisplayColor);

                ellipse.Fill = brush;

                StackPanel panel = new StackPanel();
                panel.Orientation = Orientation.Horizontal;
                panel.Children.Add(ellipse);
                panel.Children.Add(new TextBlock() { Text = importedFile.FileName});

                filesListing.Items.Add(panel);
            }
        }

        private void filesListing_Loaded(object sender, RoutedEventArgs e)
        {
            ListFiles();
        }

        /// <summary>
        /// Import excel button click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void importExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";

            if (openFileDialog.ShowDialog() == true)
            {
                string excelFilePath = openFileDialog.FileName;
                string fileName = System.IO.Path.GetFileNameWithoutExtension(excelFilePath);
                string fileExtension = System.IO.Path.GetExtension(excelFilePath);

                int counter = 1;
                string newFileName = fileName;

                while (selectedFiles.Any(file => file.FileName == newFileName))
                {
                    newFileName = $"{fileName}_{counter}";
                    counter++;
                }

                Color displayColor = selectedFiles.FirstOrDefault(file => file.FileName == newFileName)?.DisplayColor ?? GenerateRandomColor();

                ImportedFile importedFile = new ImportedFile
                {
                    FileName = newFileName,
                    FilePath = excelFilePath,
                    DisplayColor = displayColor
                };

                selectedFiles.Add(importedFile);

                ListFiles();
            }
        }

        /// <summary>
        /// Generate color for files.
        /// </summary>
        /// <returns>A color.</returns>
        private Color GenerateRandomColor()
        {
            Random random = new Random();
            return Color.FromRgb((byte)random.Next(256), (byte)random.Next(256), (byte)random.Next(256));
        }

        /// <summary>
        /// Delete file with context menu from list.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DeleteMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (filesListing.SelectedItem != null && filesListing.SelectedItem is StackPanel selectedStackPanel) 
            {
                if (selectedStackPanel.Children[1] is TextBlock textBlock)
                {
                    string selectedFileName = textBlock.Text;

                    ImportedFile m_selectedImportedFile = selectedFiles.FirstOrDefault(file => file.FileName == selectedFileName);

                    if (m_selectedImportedFile != null)
                    {
                        selectedFiles.Remove(m_selectedImportedFile);
                        ListFiles();
                    }
                }
            }
        }

        /// <summary>
        /// Context menu.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void filesListing_ContextMenuOpening(object sender, ContextMenuEventArgs e)
        {
            if (filesListing.SelectedItem == null)
            {
                e.Handled = true;
            }
        }

        /// <summary>
        /// Double click to file.
        /// TODO: More functions will be implement later.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void filesListing_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left && filesListing.SelectedItem != null && filesListing.SelectedItem is StackPanel selectedStackPanel)
            {
                if (selectedStackPanel.Children[1] is TextBlock textBlock)
                {
                    string selectedFileName = textBlock.Text;
                
                    ImportedFile m_selectedImportedFile = selectedFiles.FirstOrDefault(file => file.FileName == selectedFileName);

                    if (m_selectedImportedFile != null)
                    {
                        tabControlBorder.BorderBrush = new SolidColorBrush(m_selectedImportedFile.DisplayColor);
                    }
                }
            }
        }
    }
}
