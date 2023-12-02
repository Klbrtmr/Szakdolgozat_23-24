using ExcelDataReader;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
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
using LiveCharts;
using LiveCharts.Wpf;

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

        //indoklások az elrendezés miatt a szakdogaban
        //struktura szarmazasok

        /// <summary>
        /// The list in which the imported files are stored.
        /// </summary>
        private List<ImportedFile> selectedFiles = new List<ImportedFile>();

        public object[,] cellValues { get; private set; }

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

                

                using (var streamval = File.Open(excelFilePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(streamval))
                    {
                        var configuration = new ExcelDataSetConfiguration
                        {
                            ConfigureDataTable = _ => new ExcelDataTableConfiguration
                            {
                                UseHeaderRow = false
                            }
                        };
                        var dataSet = reader.AsDataSet(configuration);

                        if (dataSet.Tables.Count>0)
                        {
                            var dataTable = dataSet.Tables[0];
                        }

                        cellValues = new object[dataSet.Tables[0].Rows.Count, dataSet.Tables[0].Columns.Count];

                        for (int i = 0; i < dataSet.Tables[0].Rows.Count; i++)
                        {
                            for (int j = 0; j < dataSet.Tables[0].Columns.Count; j++)
                            {
                                cellValues[i, j] = dataSet.Tables[0].Rows[i].ItemArray[j]; 
                            }
                        }
                    }
                }

                ImportedFile importedFile = new ImportedFile
                {
                    FileName = newFileName,
                    FilePath = excelFilePath,
                    DisplayColor = displayColor,
                    ExcelData = cellValues,
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
                        DisplayExcelData(m_selectedImportedFile);
                    }
                }
            }
        }
        
        private void mainTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (mainTabControl.SelectedItem is TabItem selectedTabItem)
            {
                if (selectedTabItem.DataContext is ImportedFile selectedImportedFile)
                {
                    //DisplayExcelData(selectedImportedFile);
                }
            }
        }

        private void DisplayExcelData(ImportedFile importedFile)
        {
            DataTable dataTable = ConvertArrayToDataTable(importedFile.ExcelData);
            excelDataGrid.ItemsSource = dataTable.DefaultView;

            DataTable customDataTable = ConvertArrayToDataTable(importedFile.ExcelData);
            excelCustomDataGrid.ItemsSource = customDataTable.DefaultView;
        }

        private DataTable ConvertArrayToDataTable(object[,] array)
        {
            DataTable dataTable = new DataTable();
            
            for (int i = 0; i < array.GetLength(1); i++)
            {
                dataTable.Columns.Add($"{array[0,i]}");
            }

            for (int i = 1; i < array.GetLength(0); i++)
            {
                DataRow dataRow = dataTable.NewRow();
                for (int j = 0; j < array.GetLength(1); j++)
                {
                    dataRow[j] = array[i, j];
                }
                dataTable.Rows.Add(dataRow);
            }

            return dataTable;
        }
    }
}
