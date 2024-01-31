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
using OfficeOpenXml;
using ICSharpCode.SharpZipLib.Zip;
using System.IO.Compression;

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
            importProgressBar = FindName("importProgressBar") as ProgressBar;
        }

        //indoklások az elrendezés miatt a szakdogaban
        //struktura szarmazasok

        /// <summary>
        /// The list in which the imported files are stored.
        /// </summary>
        private List<ImportedFile> selectedFiles = new List<ImportedFile>();

        /// <summary>
        /// Two-dimension array for original excel datas.
        /// </summary>
        public object[,] cellValues { get; private set; }

        /// <summary>
        /// Two-dimension array for custom changed excel datas.
        /// </summary>
        public object[,] customCellValues { get; private set; }

        /// <summary>
        /// Actual importedFile.
        /// </summary>
        private ImportedFile m_ImportedFile;

        /// <summary>
        /// Actual opened dataTable.
        /// </summary>
        private DataTable m_CustomDataTable;

        /// <summary>
        /// TransformGroup for scaling, zoom-in, zoom-out
        /// </summary>
        private TransformGroup transformGroup = new TransformGroup();

        /// <summary>
        /// Number of imported file from one package file.
        /// </summary>
        private int m_importedFileNumber;

        private int currentID = 0;

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

        /// <summary>
        /// Event handler to ListFiles() method.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void filesListing_Loaded(object sender, RoutedEventArgs e)
        {
            ListFiles();
        }

        /// <summary>
        /// Import excel button click.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void importExcel_Click(object sender, RoutedEventArgs e)
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

                        importProgressBar.Visibility = Visibility.Visible;

                        try
                        {
                            await Task.Run(() =>
                            {
                                for (int i = 0; i < dataSet.Tables[0].Rows.Count; i++)
                                {
                                    for (int j = 0; j < dataSet.Tables[0].Columns.Count; j++)
                                    {
                                        cellValues[i, j] = dataSet.Tables[0].Rows[i].ItemArray[j];
                                    }
                                }
                                System.Threading.Thread.Sleep(2000);
                            });
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                        finally
                        {
                            importProgressBar.Visibility = Visibility.Collapsed;
                        }

                        
                    }
                }

                customCellValues = cellValues;

                int newID = GenerateNewID();
                ImportedFile importedFile = new ImportedFile
                {
                    ID = newID,
                    FileName = newFileName,
                    FilePath = excelFilePath,
                    DisplayColor = displayColor,
                    ExcelData = cellValues,
                    CustomExcelData = customCellValues,
                };

                Dispatcher.Invoke(() =>
                {
                    selectedFiles.Add(importedFile);
                    ListFiles();
                });
            }
        }

        private int GenerateNewID()
        {
            return currentID++;
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

                    ImportedFile m_selectedImportedFile = selectedFiles.FirstOrDefault(file => file.FileName.Equals(selectedFileName));

                    if (m_selectedImportedFile != null)
                    {
                        m_ImportedFile = m_selectedImportedFile;
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
        /// Display Excel Data in a dataTable.
        /// </summary>
        /// <param name="importedFile">Selected imported file.</param>
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
            UpdateAllCharts(importedFile, dataTable, m_CustomDataTable);
        }

        /// <summary>
        /// Convert from the imported file datas to datatable for an easier handle.
        /// </summary>
        /// <param name="array">Two-dimension array from imported file.</param>
        /// <returns>Returns the datatable from excel data.</returns>
        /*private DataTable ConvertArrayToDataTable(object[,] array)
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
            }
            return dataTable;
        }*/

        
        private DataTable ConvertArrayToDataTable(object[,] array)
        {
            DataTable dataTable = new DataTable();

            for (int i = 0; i < array.GetLength(1); i++)
            {
                dataTable.Columns.Add($"{array[0, i]}");
            }

            for (int i = 1; i < array.GetLength(0); i++)
            {
                DataRow dataRow = dataTable.NewRow();

                for (int j = 0; j < array.GetLength(1); j++)
                {
                    dataRow[j] = (double)array[i, j];
                }

                dataTable.Rows.Add(dataRow);
            }
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

        /// <inheritdoc cref="UpdateChart(ImportedFile, DataTable)"/>
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

        private void UpdateAllCharts(ImportedFile importedFile, DataTable dataTable, DataTable customDataTable)
        {

        }
        
        /// <summary>
        /// Handle when in the custom data grid edited cell(s). Save the new value to the customCellValues array.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

                        int rowIndex = e.Row.GetIndex() + 1;
                        int columnIndex = e.Column.DisplayIndex;

                        ImportedFile selectedFile = m_ImportedFile;

                        if (selectedFile != null)
                        {
                            selectedFile.CustomExcelData[rowIndex, columnIndex] = parsedValue;
                            UpdateCustomChart(selectedFile, m_CustomDataTable);
                        }
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

        /// <summary>
        /// Exported custom data table to the new excel file.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void exporttoexcel_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel file (*.xlsx)|*.xlsx",
                Title = "Save",
                FileName = System.IO.Path.GetFileNameWithoutExtension(m_ImportedFile.FileName) + "_customtable.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                string outputPath = saveFileDialog.FileName;

                using (var stream = File.Create(outputPath))
                {
                    using (var package = new ExcelPackage(stream))
                    {
                        var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                        for (int i = 0; i < m_ImportedFile.CustomExcelData.GetLength(0); i++)
                        {
                            for (int j = 0; j < m_ImportedFile.CustomExcelData.GetLength(1); j++)
                            {
                                worksheet.Cells[i + 1, j + 1].Value = m_ImportedFile.CustomExcelData[i,j];
                            }
                        }

                        package.Save();
                    }

                    MessageBox.Show("File saved successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
        }

        /// <summary>
        /// Close project.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void closeProject_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show("Are you sure?", "Close", MessageBoxButton.OKCancel, MessageBoxImage.Hand);
            if (result == MessageBoxResult.OK)
            {
                Application.Current.Shutdown();
            }
        }

        /// <summary>
        /// "Save as PNG" toolbar button eventhandler. 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void saveaspng_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "PNG Files (*.png)|*.png",
                Title = "Save as PNG"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                SaveChartAsPng(saveFileDialog.FileName, plotter);
            }
        }

        /// <inheritdoc cref="saveaspng_Click(object, RoutedEventArgs)"/>
        private void customsaveaspng_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "PNG Files (*.png)|*.png",
                Title = "Save as PNG"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                SaveChartAsPng(saveFileDialog.FileName, customplotter);
            }
        }

        /// <summary>
        /// Save chart to PNG.
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="chartName"></param>
        private void SaveChartAsPng(string filePath, InteractiveDataDisplay.WPF.Chart chartName)
        {
            // Create a RenderTargetBitmap for the chart
            var renderTargetBitmap = new RenderTargetBitmap((int)chartName.ActualWidth, (int)chartName.ActualHeight, 96, 96, PixelFormats.Pbgra32);
            renderTargetBitmap.Render(chartName);

            // Create a PNG encoder
            var pngEncoder = new PngBitmapEncoder();
            pngEncoder.Frames.Add(BitmapFrame.Create(renderTargetBitmap));

            // Save to file
            using (var stream = File.Create(filePath))
            {
                pngEncoder.Save(stream);
            }
        }
        /*
        private void exportproject_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Package file (*.zip)|*.zip";

            if (saveFileDialog.ShowDialog() == true)
            {
                string zipFileName = saveFileDialog.FileName;

                try
                {
                    // Create temporary directory
                    string tempDirectory = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid().ToString());
                    Directory.CreateDirectory(tempDirectory);

                    // Export to temporary directory
                    foreach (var selectedFile in selectedFiles)
                    {
                        // Export to temporary excel file
                        ExportToExcel(selectedFile, tempDirectory);
                    }

                    // Create package file
                    FastZip fastZip = new FastZip();
                    fastZip.CreateZip(zipFileName, tempDirectory, true, "");

                    // Delete temporary directory
                    Directory.Delete(tempDirectory, true);

                    MessageBox.Show("Files saved and compressed successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"An error occurred while exporting and compressing files: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }*/

        private void exportproject_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "EDF file (*.edf)|*.edf";

            if (saveFileDialog.ShowDialog() == true)
            {
                string edfFileName = saveFileDialog.FileName;

                try
                {
                    // Create temporary directory
                    string tempDirectory = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid().ToString());
                    Directory.CreateDirectory(tempDirectory);

                    // Export to temporary directory
                    foreach (var selectedFile in selectedFiles)
                    {
                        // Export to temporary excel file
                        ExportToExcel(selectedFile, tempDirectory);
                    }

                    // Create EDF file (not an actual ZIP file, just a custom extension)
                    FastZip fastZip = new FastZip();
                    fastZip.CreateZip(edfFileName, tempDirectory, true, "");

                    // Delete temporary directory
                    Directory.Delete(tempDirectory, true);

                    MessageBox.Show("Files saved successfully in EDF format.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"An error occurred while exporting files: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }


            private void ExportToExcel(ImportedFile importedFile, string outputDirectory)
        {
            string outputPath = System.IO.Path.Combine(outputDirectory, importedFile.FileName + "_customtable.xlsx");

            using (var stream = File.Create(outputPath))
            {
                using (var package = new ExcelPackage(stream))
                {
                    var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                    for (int i = 0; i < importedFile.CustomExcelData.GetLength(0); i++)
                    {
                        for (int j = 0; j < importedFile.CustomExcelData.GetLength(1); j++)
                        {
                            worksheet.Cells[i + 1, j + 1].Value = importedFile.CustomExcelData[i, j];
                        }
                    }

                    package.Save();
                }
            }
        }
        /*
        private void importProject_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Package file (*.zip)|*.zip";

            if (openFileDialog.ShowDialog() == true)
            {
                string zipFilePath = openFileDialog.FileName;

                try
                {
                    // Create temporary directory
                    string tempDirectory = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid().ToString());
                    m_importedFileNumber = 0;
                    Directory.CreateDirectory(tempDirectory);

                    // Open Package file
                    using (ZipInputStream zipInputStream = new ZipInputStream(File.OpenRead(zipFilePath)))
                    {
                        ZipEntry entry;
                        while ((entry = zipInputStream.GetNextEntry()) != null)
                        {
                            string entryPath = System.IO.Path.Combine(tempDirectory, entry.Name);

                            if (!entry.IsDirectory)
                            {
                                using (FileStream entryStream = File.Create(entryPath))
                                {
                                    zipInputStream.CopyTo(entryStream);
                                }
                            }
                            else
                            {
                                Directory.CreateDirectory(entryPath);
                            }
                        }
                    }

                    foreach (var excelFile in Directory.GetFiles(tempDirectory, "*.xlsx"))
                    {
                        // Check for file is excel file
                        if (IsValidExcelFile(excelFile))
                        {
                            // Import
                            ImportExcelFile(excelFile);
                        }
                        else
                        {
                            MessageBox.Show($"Error: {excelFile} is not a valid Excel file.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }

                    // Delete temporary directory
                    Directory.Delete(tempDirectory, true);

                    if (m_importedFileNumber == 0)
                    {
                        MessageBox.Show("Package file is not valid. 0 file imported.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    else if (m_importedFileNumber == 1)
                    {
                        MessageBox.Show("File imported successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else if (m_importedFileNumber > 1)
                    {
                        MessageBox.Show("Files imported successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"An error occurred while importing files: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
        */
        private void importProject_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "EDF file (*.edf)|*.edf";

            if (openFileDialog.ShowDialog() == true)
            {
                string edfFilePath = openFileDialog.FileName;

                try
                {
                    // Create temporary directory
                    string tempDirectory = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid().ToString());
                    m_importedFileNumber = 0;
                    Directory.CreateDirectory(tempDirectory);

                    // Copy files from EDF file to temporary directory
                    using (ZipInputStream zipInputStream = new ZipInputStream(File.OpenRead(edfFilePath)))
                    {
                        ZipEntry entry;
                        while ((entry = zipInputStream.GetNextEntry()) != null)
                        {
                            string entryPath = System.IO.Path.Combine(tempDirectory, entry.Name);

                            if (!entry.IsDirectory)
                            {
                                using (FileStream entryStream = File.Create(entryPath))
                                {
                                    zipInputStream.CopyTo(entryStream);
                                }
                            }
                            else
                            {
                                Directory.CreateDirectory(entryPath);
                            }
                        }
                    }

                    foreach (var excelFile in Directory.GetFiles(tempDirectory, "*.xlsx"))
                    {
                        // Check for file is excel file
                        if (IsValidExcelFile(excelFile))
                        {
                            // Import
                            ImportExcelFile(excelFile);
                        }
                        else
                        {
                            MessageBox.Show($"Error: {excelFile} is not a valid Excel file.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }

                    // Delete temporary directory
                    Directory.Delete(tempDirectory, true);

                    if (m_importedFileNumber == 0)
                    {
                        MessageBox.Show("EDF file is not valid. 0 file imported.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    else if (m_importedFileNumber == 1)
                    {
                        MessageBox.Show("File imported successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else if (m_importedFileNumber > 1)
                    {
                        MessageBox.Show("Files imported successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"An error occurred while importing files: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private bool IsValidExcelFile(string filePath)
        {
            try
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        return true;
                    }
                }
            }
            catch
            {
                return false;
            }
        }

        private void ImportExcelFile(string excelFilePath)
        {
            try
            {
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

                        if (dataSet.Tables.Count > 0)
                        {
                            var firstCellValue = dataSet.Tables[0].Rows[0].ItemArray[0];
                            if (firstCellValue == null || !firstCellValue.ToString().Equals("Sample", StringComparison.OrdinalIgnoreCase))
                            {
                                MessageBox.Show("Error: Wrong file structure or file is corrupt.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                return;
                            }

                            var dataTable = dataSet.Tables[0];
                        }

                        object[,] cellValues = new object[dataSet.Tables[0].Rows.Count, dataSet.Tables[0].Columns.Count];

                        for (int i = 0; i < dataSet.Tables[0].Rows.Count; i++)
                        {
                            for (int j = 0; j < dataSet.Tables[0].Columns.Count; j++)
                            {
                                cellValues[i, j] = dataSet.Tables[0].Rows[i].ItemArray[j];
                            }
                        }

                                int newID = GenerateNewID();
                                Color displayColor = selectedFiles.FirstOrDefault(file => file.FileName == System.IO.Path.GetFileNameWithoutExtension(excelFilePath))?.DisplayColor ?? GenerateRandomColor();
                                ImportedFile importedFile = new ImportedFile
                                {
                                    ID = newID,
                                    FileName = System.IO.Path.GetFileNameWithoutExtension(excelFilePath),
                                    FilePath = excelFilePath,
                                    DisplayColor = displayColor,
                                    ExcelData = cellValues,
                                    CustomExcelData = cellValues,
                                };

                                    selectedFiles.Add(importedFile);
                                    m_importedFileNumber++;
                                    ListFiles();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while importing the file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void allChartsSaveAsPng_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "PNG Files (*.png)|*.png",
                Title = "Save as PNG"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                SaveChartAsPng(saveFileDialog.FileName, allChartsPlotter);
            }
        }
    }
}