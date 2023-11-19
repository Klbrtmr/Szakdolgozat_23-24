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
            //filesListing.ContextMenu = CreateContextMenu();
        }

        private List<ImportedFile> selectedFiles = new List<ImportedFile>();
        /*
        private ContextMenu CreateContextMenu()
        {
            ContextMenu contextMenu = new ContextMenu();
            MenuItem deleteMenuItem = new MenuItem();
            deleteMenuItem.Header = "Delete";
            deleteMenuItem.Click += DeleteMenuItem_Click;
            contextMenu.Items.Add(deleteMenuItem);

            return contextMenu;
        }*/

        private void ListFiles()
        {
            /*
             * 0. próbálkozás
             * Képlistázós dolog
            string path = "../../Assets/";
            string[] files = Directory.GetFiles(path, "*.jpg");
            
            foreach (string file in files) 
            {
                filesListing.Items.Add(System.IO.Path.GetFileName(file));
            }

            if (filesListing.Items.Count == 0)
            {
                tabControlBorder.BorderBrush = Brushes.Red;
            }
            else
            {
                tabControlBorder.BorderBrush = Brushes.Green;
            }*/

            /*
             * 1. próbálkozás

            filesListing.Items.Clear();
            
            foreach (var file in files)
            {
                filesListing.Items.Add(System.IO.Path.GetFileName(file));
            }

            if (filesListing.Items.Count == 0)
            {
                tabControlBorder.BorderBrush = Brushes.Red;
            }
            else
            {
                tabControlBorder.BorderBrush = Brushes.Green;
            }*/

            filesListing.ItemsSource = null;

            filesListing.Items.Clear();

            /*
            foreach (string file in selectedFiles)
            {
                filesListing.Items.Add(System.IO.Path.GetFileName(file));
            }*/

            foreach (ImportedFile importedFile in selectedFiles)
            {
                filesListing.Items.Add(importedFile.FileName);
            }

            if (filesListing.Items.Count == 0)
            {
                tabControlBorder.BorderBrush = Brushes.Red;
            }
            else
            {
                tabControlBorder.BorderBrush = Brushes.Green;
            }

        }

        private void filesListing_Loaded(object sender, RoutedEventArgs e)
        {/*
          * 0. próbálkozás
          * Képlistázós dolog, plusz a kör hozzá 
            ListFiles();

            List<UIElement> modifiedItems = new List<UIElement>();

            Random random = new Random();

            foreach (var item in filesListing.Items)
            {
                Ellipse ellipse = new Ellipse();
                ellipse.Width = 15;
                ellipse.Height = 15;

                Color randomColor = Color.FromRgb((byte)random.Next(256), (byte)random.Next(256), (byte)random.Next(256));
                SolidColorBrush brush = new SolidColorBrush(randomColor);

                ellipse.Fill = brush;

                StackPanel stackPanel = new StackPanel();
                stackPanel.Orientation = Orientation.Horizontal;
                stackPanel.Children.Add(ellipse);
                stackPanel.Children.Add(new TextBlock() { Text = item.ToString() });
                
                modifiedItems.Add(stackPanel);
            }

            filesListing.Items.Clear();

            foreach (var item in modifiedItems)
            {
                filesListing.Items.Add(item);
            }*/

            /*
             * 1. próbálkozás
            List<UIElement> modifiedItems = new List<UIElement>();

            Random random = new Random();
            foreach (var item in filesListing.Items)
            {
                TextBlock textBlock = new TextBlock() { Text = item.ToString()};
                modifiedItems.Add(textBlock);
            }

            filesListing.Items.Clear();

            foreach (var item in modifiedItems)
            {
                filesListing.Items.Add(item);
            }*/

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

                ImportedFile importedFile = new ImportedFile
                {
                    FileName = newFileName,
                    FilePath = excelFilePath
                };

                selectedFiles.Add(importedFile);

                ListFiles();
            }
        }

        private void DeleteMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (filesListing.SelectedItem != null) 
            {
                string selectedFileName = filesListing.SelectedItem.ToString();

                ImportedFile selectedImportedFile = selectedFiles.FirstOrDefault(file => file.FileName == selectedFileName);

                if (selectedImportedFile != null)
                {
                    selectedFiles.Remove(selectedImportedFile);
                    ListFiles();
                }
            }
        }

        private void filesListing_ContextMenuOpening(object sender, ContextMenuEventArgs e)
        {
            if (filesListing.SelectedItem == null)
            {
                e.Handled = true;
            }
        }
    }
}
