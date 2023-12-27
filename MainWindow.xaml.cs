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
using LiveCharts.Defaults;
using Steema.TeeChart;
using Steema.TeeChart.WPF;
using Steema.TeeChart.Themes;
using InteractiveDataDisplay.WPF;
using InteractiveDataDisplay;
using System.Reactive.Linq;
using ClosedXML.Excel;

namespace Szakdolgozat
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public SeriesCollection ChartValues { get; set; } = new SeriesCollection();
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

        private ObservablePoint observablePoint;

        //Actual importedFile helper
        private ImportedFile m_ImportedFile;
        private DataTable m_CustomDataTable;

        // For scaling, zoom-in, zoom-out
        private TransformGroup transformGroup = new TransformGroup();

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
                            var firstCellValue = dataSet.Tables[0].Rows[0].ItemArray[0];
                            if (firstCellValue == null || !firstCellValue.ToString().Equals("Sample", StringComparison.OrdinalIgnoreCase))
                            {
                                MessageBox.Show("Error: Wrong file structure or file is corrupt.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                return;
                            }

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
                    CustomExcelData = cellValues,
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="importedFile"></param>
        private void DisplayExcelData(ImportedFile importedFile)
        {
            DataTable dataTable = ConvertArrayToDataTable(importedFile.ExcelData);
            excelDataGrid.ItemsSource = dataTable.DefaultView;

            //DataTable customDataTable = ConvertArrayToDataTable(importedFile.ExcelData);
            m_CustomDataTable = ConvertArrayToDataTable(importedFile.CustomExcelData);
            excelCustomDataGrid.ItemsSource = m_CustomDataTable.DefaultView;

            this.m_ImportedFile = importedFile;
            UpdateChart(importedFile, dataTable);
            UpdateCustomChart(importedFile, m_CustomDataTable);
        }

        /// <summary>
        /// Convert from the imported file datas to datatable for a easier handle.
        /// </summary>
        /// <param name="array">Two-dimension array from imported file.</param>
        /// <returns>Returns the datatable from excel data.</returns>
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
                    if (j == 0)
                    {
                        dataTable.Columns[j].ReadOnly = true;
                    }
                }
                dataTable.Rows.Add(dataRow);
            }


            // innen folytatni...
            /*
            if (dataTable.Rows.Count > 0)
            {
                DataRow lastRow = dataTable.Rows[dataTable.Rows.Count - 1];
                foreach (DataColumn column in dataTable.Columns)
                {
                    lastRow[column] = lastRow[column];
                }
            }*/

            return dataTable;
        }



        /// <summary>
        /// Created chart view from actual datas when the file is open.
        /// </summary>
        /// <param name="importedFile">Imported excel file.</param>
        /// <param name="dataTable">Datatable from imported excel file.</param>
        private void UpdateChart(ImportedFile importedFile, DataTable dataTable)
        {
            //Clear for other call
            lines1.Children.Clear();

            //Chart title
            plotter.Title = importedFile.FileName;

            if (dataTable.Rows.Count>0)
            {
                var xColumn = dataTable.Columns[0];
                var yColumns = dataTable.Columns.Cast<DataColumn>().Skip(1).ToArray();

                var x = dataTable.AsEnumerable().Select(row =>
                {
                    var xValue = row[xColumn];
                    return xValue != DBNull.Value ? Convert.ToDouble(xValue) : 0.0;
                }).ToArray();

                foreach ( var yColumn in yColumns)
                {
                    var lg = new LineGraph();
                    lines1.Children.Add(lg);
                    lg.Stroke = new SolidColorBrush(Color.FromArgb(255, 0, 0, 255));
                    lg.Description = String.Format($"{yColumn.ColumnName}");
                    lg.StrokeThickness = 2;

                    var y = dataTable.AsEnumerable().Select(row =>
                    {
                        var yValue = row[yColumn];
                        if (yValue != DBNull.Value)
                        {
                            double parsedValue;
                            if (double.TryParse(yValue.ToString(), out parsedValue))
                            {
                                return parsedValue;
                            }
                            else
                            {
                                // Logic for invalid values
                                row[yColumn] = 0.0;
                                return 0.0;
                            }
                        }
                        else
                        {
                            // Logic for empty values
                            row[yColumn] = 0.0;
                            return 0.0;
                        }
                    }).ToArray();

                    lg.Plot(x, y);
                }
            }
        }

        private void UpdateCustomChart(ImportedFile importedFile, DataTable dataTable)
        {
            //Clear for other call
            customLines1.Children.Clear();

            //Chart title
            customplotter.Title = importedFile.FileName;

            bool invalidvalues = false;

            if (dataTable.Rows.Count > 0)
            {
                var xColumn = dataTable.Columns[0];
                var yColumns = dataTable.Columns.Cast<DataColumn>().Skip(1).ToArray();

                var x = dataTable.AsEnumerable().Select(row =>
                {
                    var xValue = row[xColumn];
                    return xValue != DBNull.Value ? Convert.ToDouble(xValue) : 0.0;
                }).ToArray();

                foreach (var yColumn in yColumns)
                {
                    var lg = new LineGraph();
                    customLines1.Children.Add(lg);
                    lg.Stroke = new SolidColorBrush(Color.FromArgb(255, 0, 0, 255));
                    lg.Description = String.Format($"{yColumn.ColumnName}");
                    lg.StrokeThickness = 2;

                    var y = dataTable.AsEnumerable().Select(row =>
                    {
                        var yValue = row[yColumn];
                        if (yValue != DBNull.Value)
                        {
                            double parsedValue;
                            if (double.TryParse(yValue.ToString(), out parsedValue))
                            {
                                return parsedValue;
                            }
                            else
                            {
                                // Logic for invalid values
                                invalidvalues = true;
                                row[yColumn] = 0.0;
                                return 0.0;
                            }
                        }
                        else
                        {
                            // Logic for empty values
                            invalidvalues = true;
                            row[yColumn] = 0.0;
                            return 0.0;
                        }
                    }).ToArray();

                    lg.Plot(x, y);
                }
            }
            if (invalidvalues == true)
            {
                MessageBox.Show("Excel contains invalid values! Invalid values are set to 0.", "Import warning!" , MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        
        private void excelCustomDataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {

            if (e.Column.DisplayIndex != 0 && e.Row.Item is DataRowView rowView)
            {
                object modifiedValue = ((TextBox)e.EditingElement).Text;

                if (modifiedValue != DBNull.Value)
                {
                    double parsedValue;
                    if (double.TryParse(modifiedValue.ToString(), out parsedValue))
                    {
                        rowView.Row[e.Column.DisplayIndex] = parsedValue;
                        UpdateCustomChart(m_ImportedFile, m_CustomDataTable);
                    }
                    else
                    {
                        MessageBox.Show("Invalid value!", "Invalid value error!", MessageBoxButton.OK, MessageBoxImage.Error);
                        ((TextBox)e.EditingElement).Text = rowView.Row[e.Column.DisplayIndex].ToString();
                    }
                }

            }
        }

        private void ZoomIn_Click(object sender, RoutedEventArgs e)
        {
            ApplyZoom(1.1);
        }

        private void ZoomOut_Click(object sender, RoutedEventArgs e)
        {
            ApplyZoom(0.9);
        }

        private void ZoomOutCompletly_Click(object sender, RoutedEventArgs e)
        {
            ResetZoom();
        }

        private void ApplyZoom(double factor)
        {
            ScaleTransform scaleTransform = new ScaleTransform(factor, factor);
            transformGroup.Children.Add(scaleTransform);
            lines1.LayoutTransform = transformGroup;
        }

        private void ResetZoom()
        {
            transformGroup.Children.Clear();
            transformGroup.Children.Add(new ScaleTransform(1.0, 1.0));
        }
    }
}