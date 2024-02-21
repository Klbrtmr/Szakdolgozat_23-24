using ExcelDataReader;
using ICSharpCode.SharpZipLib.Zip;
using LiveCharts;
using Microsoft.Win32;
using OfficeOpenXml;
using ScottPlot;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reactive.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

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
        /// Actual opened original dataTable.
        /// </summary>
        private DataTable m_OriginalDataTable;

        /// <summary>
        /// Actual opened custom dataTable.
        /// </summary>
        private DataTable m_CustomDataTable;

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
                panel.Orientation = System.Windows.Controls.Orientation.Horizontal;
                panel.Children.Add(ellipse);
                panel.Children.Add(new TextBlock() { Text = importedFile.FileName });

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

                System.Windows.Media.Color displayColor = selectedFiles.FirstOrDefault(file => file.FileName == newFileName)?.DisplayColor ?? GenerateRandomColorForFiles();



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
                            var secondCellValue = dataSet.Tables[0].Rows[0].ItemArray[1];
                            if (firstCellValue == null || !firstCellValue.ToString().Equals("Sample", StringComparison.OrdinalIgnoreCase) &&
                                secondCellValue == null || !secondCellValue.ToString().Equals("Events", StringComparison.OrdinalIgnoreCase))
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
                                for (int i = 0; i < dataSet.Tables[0].Columns.Count; i++)
                                {
                                    cellValues[0,i] = dataSet.Tables[0].Rows[0].ItemArray[i];
                                }

                                for (int i = 1; i < dataSet.Tables[0].Rows.Count; i++)
                                {
                                    for (int j = 0; j < dataSet.Tables[0].Columns.Count; j++)
                                    {
                                        var actualValue = dataSet.Tables[0].Rows[i].ItemArray[j];

                                        if (double.TryParse(actualValue.ToString(), out double parsedValue))
                                        {
                                            cellValues[i,j] = parsedValue;
                                        }
                                        else if (actualValue.ToString().EndsWith("Event"))
                                        {
                                            cellValues[i, j] = actualValue.ToString();
                                        }
                                        else
                                        {
                                            // Logic for invalid values
                                            cellValues[i,j] = 0.0;
                                        }
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

                // customCellValues = cellValues;

                customCellValues = new object[cellValues.GetLength(0), cellValues.GetLength(1)];
                Array.Copy(cellValues, customCellValues, cellValues.Length);

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
        private System.Windows.Media.Color GenerateRandomColorForFiles()
        {
            Random random = new Random();
            return System.Windows.Media.Color.FromRgb((byte)random.Next(256), (byte)random.Next(256), (byte)random.Next(256));
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
            m_OriginalDataTable = ConvertArrayToDataTable(importedFile.ExcelData);
            excelDataGrid.ItemsSource = m_OriginalDataTable.DefaultView;

            //DataTable customDataTable = ConvertArrayToDataTable(importedFile.ExcelData);
            m_CustomDataTable = ConvertArrayToDataTable(importedFile.CustomExcelData);
            excelCustomDataGrid.ItemsSource = m_CustomDataTable.DefaultView;

            DataTable configTable = ConvertArrayToDataTableConfigPage(importedFile.ExcelData);
            configGrid.ItemsSource = configTable.DefaultView;

            this.m_ImportedFile = importedFile;
            UpdateChart(importedFile, m_OriginalDataTable);
            UpdateCustomChart(importedFile, m_CustomDataTable);
            // UpdateSpecialChart(importedFile, m_OriginalDataTable);
        }

        /// <summary>
        /// Convert from the imported file datas to datatable for an easier handle.
        /// </summary>
        /// <param name="array">Two-dimension array from imported file.</param>
        /// <returns>Returns the datatable from excel data.</returns>
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
                    var actualValue = array[i, j];
                    if (double.TryParse(actualValue.ToString(), out double parsedValue))
                    {
                        dataRow[j] = parsedValue;
                    }
                    else if (actualValue.ToString().EndsWith("Event"))
                    {
                        dataRow[j] = actualValue.ToString();
                    }
                }

                dataTable.Rows.Add(dataRow);
            }
            return dataTable;
        }

        private DataTable ConvertArrayToDataTableConfigPage(object[,] array)
        {
            DataTable dataTable = new DataTable();

            dataTable.Columns.Add("Signal Name", typeof(string));

            DataGridTemplateColumn templateColumn = new DataGridTemplateColumn();
            templateColumn.Header = "Colors";

            // DataTemplate létrehozása, ami tartalmazza a ComboBox-ot
            FrameworkElementFactory comboBoxFactory = new FrameworkElementFactory(typeof(ComboBox));
            comboBoxFactory.SetValue(ComboBox.ItemsSourceProperty, Enumerable.Range(1, 5));
            comboBoxFactory.SetBinding(ComboBox.SelectedValueProperty, new Binding("Colors"));
            DataTemplate cellTemplate = new DataTemplate { VisualTree = comboBoxFactory };
            templateColumn.CellTemplate = cellTemplate;

            if (configGrid.Columns.Count < 2)
            {
                configGrid.Columns.Add(templateColumn);
            }
            
            for (int columnIndex = 1; columnIndex < array.GetLength(1); columnIndex++)
            {
                DataRow dataRow = dataTable.NewRow();

                dataRow["Signal Name"] = array[0, columnIndex];
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
            originalChart.Plot.Clear();

            //Chart Titles
            originalChart.Plot.Title($"{importedFile.FileName}");
            originalChart.Plot.XLabel("Samples");
            originalChart.Plot.YLabel("Values");


            if (dataTable.Rows.Count > 0)
            {
                var xColumn = dataTable.Columns[0];
                var eventColumn = dataTable.Columns[1];
                // var eventColumn = dataTable.Columns.Cast<DataColumn>().Skip(1).Take(1).ToArray();
                // List<int> eventSamples = new List<int>();
                var yColumns = dataTable.Columns.Cast<DataColumn>().Skip(2).ToArray();

                var x = dataTable.AsEnumerable().Select(row =>
                {
                    var xValue = row[xColumn];
                    return xValue != DBNull.Value ? Convert.ToDouble(xValue) : 0.0;
                }).ToArray();

                var events = dataTable.AsEnumerable().Select(row =>
                {
                    var eventValue = row[eventColumn];
                    return eventValue != DBNull.Value ? eventValue : 0.0;
                }).ToArray();

                foreach (var yColumn in yColumns)
                {
                    var y = dataTable.AsEnumerable().Select(row =>
                    {
                        var yValue = row[yColumn];
                        if (yValue != DBNull.Value)
                        {
                            if (double.TryParse(yValue.ToString(), out double parsedValue))
                            {
                                return parsedValue;
                            }
                            else if (yValue.ToString().EndsWith("Event"))
                            {
                                return yValue;
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


                    // var currentLine = originalChart.Plot.Add.Scatter(x, y, ScottPlot.Color.FromHex("#000000"));
                    var currentLine = originalChart.Plot.Add.Scatter(x, y);
                    currentLine.Label = yColumn.ColumnName;
                    // currentLine.Color = ScottPlot.Color.FromHex("#000000");
                    currentLine.MarkerStyle.IsVisible = false;
                    for (int i = 0; i < events.Length; i++)
                    {
                        var xIndex = x[i];
                        var yIndex = (double)y[i];
                        var marker = originalChart.Plot.Add.Marker(xIndex, yIndex);
                        if (events[i].ToString() == "Alarm_Event")
                        {
                            marker.MarkerStyle.Fill.Color = ScottPlot.Color.FromARGB(4278190219); // Dark Blue

                        }
                        else if (events[i].ToString() == "Error_Event")
                        {
                            marker.MarkerStyle.Shape = MarkerShape.FilledTriangleUp;
                            marker.MarkerStyle.Fill.Color = ScottPlot.Color.FromARGB(4294901760); // Red
                        }
                        else
                        {
                            marker.MarkerStyle.IsVisible = false;
                        }
                    }
                }
            }
            originalChart.Plot.Axes.AutoScaleX();
            originalChart.Plot.Axes.AutoScaleY();
            originalChart.Plot.ShowLegend();
            originalChart.Refresh();
        }

        /// <inheritdoc cref="UpdateChart(ImportedFile, DataTable)"/>
        private void UpdateCustomChart(ImportedFile importedFile, DataTable dataTable)
        {
            //Clear for other call
            CustomChart.Plot.Clear();

            //Chart Titles
            CustomChart.Plot.Title($"{importedFile.FileName}");
            CustomChart.Plot.XLabel("Samples");
            CustomChart.Plot.YLabel("Values");

            bool invalidvalues = false;

            if (dataTable.Rows.Count > 0)
            {
                var xColumn = dataTable.Columns[0];
                var eventColumn = dataTable.Columns[1];
                var yColumns = dataTable.Columns.Cast<DataColumn>().Skip(2).ToArray();

                var x = dataTable.AsEnumerable().Select(row =>
                {
                    var xValue = row[xColumn];
                    return xValue != DBNull.Value ? Convert.ToDouble(xValue) : 0.0;
                }).ToArray();

                var events = dataTable.AsEnumerable().Select(row =>
                {
                    var eventValue = row[eventColumn];
                    return eventValue != DBNull.Value ? eventValue : 0.0;
                }).ToArray();

                foreach (var yColumn in yColumns)
                {
                    var y = dataTable.AsEnumerable().Select(row =>
                    {
                        var yValue = row[yColumn];
                        if (yValue != DBNull.Value)
                        {
                            if (double.TryParse(yValue.ToString(), out double parsedValue))
                            {
                                return parsedValue;
                            }
                            else if (yValue.ToString().EndsWith("Event"))
                            {
                                return yValue;
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

                    var currentLine = CustomChart.Plot.Add.Scatter(x, y);
                    currentLine.Label = yColumn.ColumnName;
                    currentLine.MarkerStyle.IsVisible = false;
                    for (int i = 0; i < events.Length; i++)
                    {
                        var xIndex = x[i];
                        var yIndex = (double)y[i];
                        var marker = CustomChart.Plot.Add.Marker(xIndex, yIndex);
                        if (events[i].ToString() == "Alarm_Event")
                        {
                            marker.MarkerStyle.Fill.Color = ScottPlot.Color.FromARGB(4278190219); // Dark Blue
                            
                        }
                        else if (events[i].ToString() == "Error_Event")
                        {
                            marker.MarkerStyle.Shape = MarkerShape.FilledTriangleUp;
                            marker.MarkerStyle.Fill.Color = ScottPlot.Color.FromARGB(4294901760); // Red
                        }
                        else
                        {
                            marker.MarkerStyle.IsVisible = false;
                        }
                    }
                }
            }
            if (invalidvalues == true)
            {
                MessageBox.Show("Excel contains invalid values! Invalid values are set to 0.", "Import warning!", MessageBoxButton.OK, MessageBoxImage.Warning);
            }

            CustomChart.Plot.Axes.AutoScaleX();
            CustomChart.Plot.Axes.AutoScaleY();
            CustomChart.Plot.ShowLegend();
            CustomChart.Refresh();
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
                    if (double.TryParse(modifiedValue.ToString(), out double parsedValue) ||
                        (modifiedValue.ToString().EndsWith("Event") && e.Column.DisplayIndex == 1))
                    {
                        if (parsedValue != 0)
                        {
                            rowView.Row[e.Column.DisplayIndex] = parsedValue;
                        }
                        else
                        {
                            rowView.Row[e.Column.DisplayIndex] = modifiedValue;
                        }


                        int rowIndex = e.Row.GetIndex() + 1;
                        int columnIndex = e.Column.DisplayIndex;

                        ImportedFile selectedFile = m_ImportedFile;

                        if (selectedFile != null)
                        {
                            if (modifiedValue.ToString().EndsWith("Event"))
                            {
                                selectedFile.CustomExcelData[rowIndex, columnIndex] = modifiedValue;
                            }
                            else
                            {
                                selectedFile.CustomExcelData[rowIndex, columnIndex] = parsedValue;
                            }

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
            else if (e.Column.DisplayIndex == 0 && e.Row.Item is DataRowView rowView2)
            {
                ((TextBox)e.EditingElement).Text = rowView2.Row[e.Column.DisplayIndex].ToString();
                MessageBox.Show("Invalid value!", "Invalid value error!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
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
                                worksheet.Cells[i + 1, j + 1].Value = m_ImportedFile.CustomExcelData[i, j];
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
                // SaveChartAsPng(saveFileDialog.FileName, plotter);
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
                //SaveChartAsPng(saveFileDialog.FileName, customplotter);
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
                            var secondCellValue = dataSet.Tables[0].Rows[0].ItemArray[1];
                            if (firstCellValue == null || !firstCellValue.ToString().Equals("Sample", StringComparison.OrdinalIgnoreCase) &&
                                secondCellValue == null || !secondCellValue.ToString().Equals("Events", StringComparison.OrdinalIgnoreCase))
                            {
                                MessageBox.Show("Error: Wrong file structure or file is corrupt.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                return;
                            }

                            var dataTable = dataSet.Tables[0];
                        }

                        object[,] cellValues = new object[dataSet.Tables[0].Rows.Count, dataSet.Tables[0].Columns.Count];

                        for (int i = 0; i < dataSet.Tables[0].Columns.Count; i++)
                        {
                            cellValues[0, i] = dataSet.Tables[0].Rows[0].ItemArray[i];
                        }

                        for (int i = 1; i < dataSet.Tables[0].Rows.Count; i++)
                        {
                            for (int j = 0; j < dataSet.Tables[0].Columns.Count; j++)
                            {
                                var actualValue = dataSet.Tables[0].Rows[i].ItemArray[j];

                                if (double.TryParse(actualValue.ToString(), out double parsedValue))
                                {
                                    cellValues[i, j] = parsedValue;
                                }
                                else if (actualValue.ToString().EndsWith("Event"))
                                {
                                    cellValues[i, j] = actualValue;
                                }
                                else
                                {
                                    // Logic for invalid values
                                    cellValues[i, j] = 0.0;
                                }
                            }
                        }

                        int newID = GenerateNewID();
                        System.Windows.Media.Color displayColor = selectedFiles.FirstOrDefault(file => file.FileName == System.IO.Path.GetFileNameWithoutExtension(excelFilePath))?.DisplayColor ?? GenerateRandomColorForFiles();

                        var customCellValue = new object[cellValues.GetLength(0), cellValues.GetLength(1)];
                        Array.Copy(cellValues, customCellValue, cellValues.Length);

                        ImportedFile importedFile = new ImportedFile
                        {
                            ID = newID,
                            FileName = System.IO.Path.GetFileNameWithoutExtension(excelFilePath),
                            FilePath = excelFilePath,
                            DisplayColor = displayColor,
                            ExcelData = cellValues,
                            CustomExcelData = customCellValue,
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

        private void myPlot_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            originalChart.Plot.Axes.AutoScaleX();
            originalChart.Plot.Axes.AutoScaleY();
            originalChart.Refresh();
            CustomChart.Plot.Axes.AutoScaleX();
            CustomChart.Plot.Axes.AutoScaleY();
            CustomChart.Refresh();
        }
    }
}