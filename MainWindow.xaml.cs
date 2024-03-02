using ExcelDataReader;
using ICSharpCode.SharpZipLib.Zip;
using InteractiveDataDisplay.WPF;
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
using Hardcodet.Wpf.TaskbarNotification;

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
            notifyIcon = new TaskbarIcon();
            
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

        private TaskbarIcon notifyIcon;

        /// <summary>
        /// Number of imported file from one package file.
        /// </summary>
        private int m_importedFileNumber;

        private int currentID = 0;

        private List<string> colorList = new List<string>{ 
            "Black", "White", "Grey", "Blond", "Brown", 
            "Dark Blue", "Cyan", "Ice", "Red", "Green", 
            "LimeGreen", "Purple", "Pink", "Yellow", "Orange" };

        private double m_Stamp = 1.0;

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

                System.Windows.Media.Color displayColor = 
                    selectedFiles.FirstOrDefault(file => file.FileName == newFileName)?.DisplayColor 
                    ?? GenerateRandomColorForFiles();



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
                            if (firstCellValue == null || !firstCellValue.ToString()
                                .Equals("Sample", StringComparison.OrdinalIgnoreCase) &&
                                secondCellValue == null || !secondCellValue.ToString()
                                .Equals("Events", StringComparison.OrdinalIgnoreCase))
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
            comboBoxFactory.SetValue(ComboBox.ItemsSourceProperty, colorList);
            comboBoxFactory.SetBinding(ComboBox.SelectedValueProperty, new Binding("Colors"));

            comboBoxFactory.AddHandler(ComboBox.SelectionChangedEvent, new SelectionChangedEventHandler(ComboBox_SelectionChanged));

            DataTemplate cellTemplate = new DataTemplate { VisualTree = comboBoxFactory };
            templateColumn.CellTemplate = cellTemplate;

            if (configGrid.Columns.Count < 2)
            {
                configGrid.Columns.Add(templateColumn);
            }
            
            for (int columnIndex = 2; columnIndex < array.GetLength(1); columnIndex++)
            {
                DataRow dataRow = dataTable.NewRow();

                dataRow["Signal Name"] = array[0, columnIndex];
                dataTable.Rows.Add(dataRow);
            }

            configGrid.ItemsSource = dataTable.DefaultView;

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
            if (sample_RadioButton.IsChecked == true)
            {
                originalChart.Plot.XLabel("Samples");
            }
            else if (time_RadioButton.IsChecked == true)
            {
                originalChart.Plot.XLabel("Time (s)");
            }
            originalChart.Plot.YLabel("Values");


            if (dataTable.Rows.Count > 0)
            {
                var xColumn = dataTable.Columns[0];
                var eventColumn = dataTable.Columns[1];
                var yColumns = dataTable.Columns.Cast<DataColumn>().Skip(2).ToArray();

                var x = dataTable.AsEnumerable().Select(row =>
                {
                    object xValue = row[xColumn];
                    if (xValue != DBNull.Value && double.TryParse(xValue.ToString(), out double doubleValue))
                    {
                        xValue = doubleValue * m_Stamp;
                    }
                    else
                    {
                        xValue = 0.0;
                    }
                    return Convert.ToDouble(xValue);
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

                    if (DarkModeToggleButton.IsChecked == true)
                    {
                        currentLine.Color = ScottPlot.Color.FromHex("#ffffff");
                    }
                    else if (DarkModeToggleButton.IsChecked == false)
                    {
                        currentLine.Color = ScottPlot.Color.FromHex("#000000");
                    }

                    currentLine.MarkerStyle.IsVisible = false;
                    for (int i = 0; i < events.Length; i++)
                    {
                        var xIndex = x[i];
                        var yIndex = (double)y[i];
                        var marker = originalChart.Plot.Add.Marker(xIndex, yIndex);
                        if (events[i].ToString() == "Alarm_Event")
                        {
                            marker.MarkerStyle.Shape = MarkerShape.FilledCircle;
                            marker.MarkerStyle.Size = 10;
                            marker.MarkerStyle.Fill.Color = ScottPlot.Color.FromARGB(4278190219); // Dark Blue
                        }
                        else if (events[i].ToString() == "Error_Event")
                        {
                            marker.MarkerStyle.Shape = MarkerShape.FilledTriangleUp;
                            marker.MarkerStyle.Size = 10;
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
            var alarmEventNumber = 0;
            var errorEventNumber = 0;


            //Clear for other call
            CustomChart.Plot.Clear();

            //Chart Titles
            CustomChart.Plot.Title($"{importedFile.FileName}");
            if (sample_RadioButton.IsChecked == true)
            {
                CustomChart.Plot.XLabel("Samples");
            }
            else if(time_RadioButton.IsChecked == true)
            {
                CustomChart.Plot.XLabel("Time (s)");
            }
            
            CustomChart.Plot.YLabel("Values");

            bool invalidvalues = false;

            if (dataTable.Rows.Count > 0)
            {
                var xColumn = dataTable.Columns[0];
                var eventColumn = dataTable.Columns[1];
                var yColumns = dataTable.Columns.Cast<DataColumn>().Skip(2).ToArray();

                var x = dataTable.AsEnumerable().Select(row =>
                {
                    object xValue = row[xColumn];
                    if (xValue != DBNull.Value && double.TryParse(xValue.ToString(), out double doubleValue))
                    {
                        xValue = doubleValue * m_Stamp;
                    }
                    else
                    {
                        xValue = 0.0;
                    }
                    return Convert.ToDouble(xValue);
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

                    if (DarkModeToggleButton.IsChecked == true)
                    {
                        currentLine.Color = ScottPlot.Color.FromHex("#ffffff");
                    }
                    else if (DarkModeToggleButton.IsChecked == false)
                    {
                        currentLine.Color = ScottPlot.Color.FromHex("#000000");
                    }

                    currentLine.MarkerStyle.IsVisible = false;
                    for (int i = 0; i < events.Length; i++)
                    {
                        var xIndex = x[i];
                        var yIndex = (double)y[i];
                        var marker = CustomChart.Plot.Add.Marker(xIndex, yIndex);
                        if (events[i].ToString() == "Alarm_Event")
                        {
                            marker.MarkerStyle.Fill.Color = ScottPlot.Color.FromARGB(4278190219); // Dark Blue
                            if (alarmEventNumber == 0)
                            {
                                marker.MarkerStyle.IsVisible = true;
                                marker.Label = events[i].ToString();
                                alarmEventNumber++;
                            }
                            
                        }
                        else if (events[i].ToString() == "Error_Event")
                        {
                            marker.MarkerStyle.Shape = MarkerShape.FilledTriangleUp;
                            marker.MarkerStyle.Fill.Color = ScottPlot.Color.FromARGB(4294901760); // Red
                            if (errorEventNumber == 0)
                            {
                                marker.MarkerStyle.IsVisible = true;
                                marker.Label = events[i].ToString();
                                errorEventNumber++;
                            }
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

            CustomChart.Plot.Legend.ManualItems.Reverse();

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
                        System.Windows.Media.Color displayColor = selectedFiles
                            .FirstOrDefault(file => file.FileName == System.IO.Path.GetFileNameWithoutExtension(excelFilePath))?.DisplayColor 
                            ?? GenerateRandomColorForFiles();

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

        private void DarkModeToggleButton_Click(object sender, RoutedEventArgs e)
        {
            if (DarkModeToggleButton.IsChecked == true)
            {
                DarkModeToggleButton.Content = ImageForModeSwitch("Assets/brightness.png");

                originalChart.Plot.Style.DarkMode();
                CustomChart.Plot.Style.DarkMode();

                DarkModeForBackGround();
                originalChart.Refresh();
                CustomChart.Refresh();

            }
            else if (DarkModeToggleButton.IsChecked == false)
            {
                DarkModeToggleButton.Content = ImageForModeSwitch("Assets/moon.png");

                LightModeForCharts();
                LightModeForBackGround();
                originalChart.Refresh();
                CustomChart.Refresh();
            }
        }

        private void LightModeForCharts()
        {
            originalChart.Plot.Style.ColorAxes(ScottPlot.Color.FromHex("#000000"));

            originalChart.Plot.Style.ColorGrids(ScottPlot.Color.FromHex("#e0e0e0"));

            originalChart.Plot.Style.Background(
                figure: ScottPlot.Color.FromHex("#ffffff"),
                data: ScottPlot.Color.FromHex("#ffffff"));

            originalChart.Plot.Style.ColorLegend(
                background: ScottPlot.Color.FromHex("#ffffff"),
                foreground: ScottPlot.Color.FromHex("#000000"),
                border: ScottPlot.Color.FromHex("#000000"));

            CustomChart.Plot.Style.ColorAxes(ScottPlot.Color.FromHex("#000000"));

            CustomChart.Plot.Style.ColorGrids(ScottPlot.Color.FromHex("#e0e0e0"));

            CustomChart.Plot.Style.Background(
                figure: ScottPlot.Color.FromHex("#ffffff"),
                data: ScottPlot.Color.FromHex("#ffffff"));

            CustomChart.Plot.Style.ColorLegend(
                background: ScottPlot.Color.FromHex("#ffffff"),
                foreground: ScottPlot.Color.FromHex("#000000"),
                border: ScottPlot.Color.FromHex("#000000"));
        }

        private void DarkModeForBackGround()
        {
            SolidColorBrush darkBackground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(40, 40, 40)); // dark gray
            SolidColorBrush soliddarkBackground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(50, 50, 50)); // dark gray
            SolidColorBrush whiteColor = new SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 255, 255));
            BackroundBehindTabs.Background = darkBackground;
            mainTabControl.Background = soliddarkBackground;
            controllerStackPanel.Background = darkBackground;
            myNameText.Foreground = whiteColor;
            homeButton.Foreground = whiteColor;
            importExcel.Foreground = whiteColor;
            importProject.Foreground = whiteColor;
            exporttoexcel.Foreground = whiteColor;
            exportproject.Foreground = whiteColor;
            closeProject.Foreground = whiteColor;
            filesListing.Background = soliddarkBackground;
            filesListing.Foreground = whiteColor;
            sample_RadioButton.Foreground = whiteColor;
            time_RadioButton.Foreground = whiteColor;

            homeButton.Icon = ImageForModeSwitch("Assets/home_light.png");
            importExcel.Icon = ImageForModeSwitch("Assets/document_light.png");
            importProject.Icon = ImageForModeSwitch("Assets/file-import_light.png");
            exporttoexcel.Icon = ImageForModeSwitch("Assets/export_light.png");
            exportproject.Icon = ImageForModeSwitch("Assets/exportfile_light.png");
            closeProject.Icon = ImageForModeSwitch("Assets/exit_light.png");
        }

        private void LightModeForBackGround()
        {
            SolidColorBrush lightBackground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(128, 128, 128)); // gray
            SolidColorBrush lighterBackground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(169, 169 ,169));
            SolidColorBrush blackColor = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 0, 0));
            SolidColorBrush whiteColor = new SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 255, 255));
            BackroundBehindTabs.Background = lightBackground;
            mainTabControl.Background = lighterBackground;
            controllerStackPanel.Background= lighterBackground;
            myNameText.Foreground = blackColor;
            homeButton.Foreground = blackColor;
            importExcel.Foreground = blackColor;
            importProject.Foreground = blackColor;
            exporttoexcel.Foreground = blackColor;
            exportproject.Foreground = blackColor;
            closeProject.Foreground = blackColor;
            filesListing.Background = whiteColor;
            filesListing.Foreground = blackColor;
            sample_RadioButton.Foreground = blackColor;
            time_RadioButton.Foreground = blackColor;

            homeButton.Icon = ImageForModeSwitch("Assets/home.png");
            importExcel.Icon = ImageForModeSwitch("Assets/document.png");
            importProject.Icon = ImageForModeSwitch("Assets/file-import.png");
            exporttoexcel.Icon = ImageForModeSwitch("Assets/export.png");
            exportproject.Icon = ImageForModeSwitch("Assets/exportfile.png");
            closeProject.Icon = ImageForModeSwitch("Assets/exit.png");
        }

        private System.Windows.Controls.Image ImageForModeSwitch(string root)
        {
            System.Windows.Controls.Image icon = new System.Windows.Controls.Image();
            icon.Source = new BitmapImage(new Uri(root, UriKind.RelativeOrAbsolute));
            icon.Width = 16;
            icon.Height = 16;
            return icon;
        }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ColorConverter colorConverter = new ColorConverter();
            ComboBox comboBox = sender as ComboBox;

            if (comboBox != null)
            {
                // Get the DataGridCell for the given ComboBox
                DataGridCell cell = FindVisualParent<DataGridCell>(comboBox);

                // Get the DataGridRow for the given DataGridCell
                DataGridRow row = FindVisualParent<DataGridRow>(cell);

                if (row != null)
                {
                    var selectedValue = comboBox.SelectedValue;
                    var colorInHexa = colorConverter.ConvertFromString(selectedValue);

                    if (originalChart != null)
                    {
                        int rowIndex = row.GetIndex();
                        int seriesIndex = rowIndex;

                        var scatterSeriesList = originalChart.Plot.GetPlottables()
                            .Where(p => p is ScottPlot.Plottables.Scatter scatter).Cast<ScottPlot.Plottables.Scatter>().ToList();

                        var customScatterSeriesList = CustomChart.Plot.GetPlottables()
                            .Where(p => p is ScottPlot.Plottables.Scatter scatter).Cast<ScottPlot.Plottables.Scatter>().ToList();

                        if ((seriesIndex >= 0 && seriesIndex < scatterSeriesList.Count)
                            && (seriesIndex >= 0 && seriesIndex < customScatterSeriesList.Count))
                        {
                            // Modify the color of the specified series in the plot
                            scatterSeriesList[seriesIndex].LineStyle.Color = ScottPlot.Color.FromHex(colorInHexa);
                            customScatterSeriesList[seriesIndex].LineStyle.Color = ScottPlot.Color.FromHex(colorInHexa);
                            originalChart.Refresh();
                            CustomChart.Refresh();
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Helper method for find visual parent
        /// </summary>
        /// <typeparam name="T">Generic parameter</typeparam>
        /// <param name="child"></param>
        /// <returns></returns>
        private T FindVisualParent<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject parentObject = VisualTreeHelper.GetParent(child);

            if (parentObject == null)
                return null;

            T parent = parentObject as T;
            if (parent != null)
            {
                return parent;
            }
            else
            {
                return FindVisualParent<T>(parentObject);
            }
        }

        private void originalChart_MouseMove(object sender, MouseEventArgs e)
        {
            var mousePos = Mouse.GetPosition(originalChart);

            // var plotMousePos = originalChart.Plot.GetCoordinates((float)mousePos.X, (float)mousePos.Y);

            foreach (var marker in originalChart.Plot.GetPlottables<ScottPlot.Plottables.Marker>())
            {
                if (marker.MarkerStyle.Shape == MarkerShape.FilledTriangleUp)
                {
                    double markerX = marker.X;
                    double markerY = marker.Y;
                    
                    Coordinates coordinates = new Coordinates(markerX, markerY);

                    var markerPixelPos = originalChart.Plot.GetPixel(coordinates);

                    // Adott távolság (pl. 5 pixel) környékén van az egér a markertől
                    if (Math.Abs(markerPixelPos.X - mousePos.X) < 50 && Math.Abs(markerPixelPos.Y - mousePos.Y) < 50)
                    {
                        // Ha közel van, akkor megjelenítjük a tooltip-et
                        ShowTooltip("Error Event", mousePos.X, mousePos.Y);
                        return;
                        
                    }
                }
            }

            // Ha nem találunk közel lévő markert, elrejtjük a tooltip-et
            HideTooltip();
        }


        private void ShowTooltip(string message, double x, double y)
        {
            // Létrehoz egy új ToolTip-t
            var m_toolTip = new ToolTip();

            // Hozzáad egy TextBlock-ot a ToolTip-hez
            TextBlock textBlock = new TextBlock();
            textBlock.Text = message;
            m_toolTip.Content = textBlock;

            // Beállítja a ToolTip pozícióját a megadott koordinátáknak megfelelően
            m_toolTip.Placement = System.Windows.Controls.Primitives.PlacementMode.Right;
            m_toolTip.PlacementTarget = originalChart;  // Az eredeti diagram a célpont
            m_toolTip.HorizontalOffset = x;
            m_toolTip.VerticalOffset = y;
            m_toolTip.Height = 40;
            m_toolTip.Width = 80;

            notifyIcon.ToolTipText = textBlock.Text;
            notifyIcon.ToolTip = m_toolTip;
            // Beállítja a ToolTip-et az eredeti diagramhoz
            // ToolTipService.SetToolTip(originalChart, m_toolTip);
        }

        private void HideTooltip()
        {
            //m_toolTip = null;
            //ToolTipService.SetToolTip(originalChart, m_toolTip);
            notifyIcon.ToolTip = null;
        }

        private void RadioButtonMode_Checked(object sender, RoutedEventArgs e)
        {
            if (sample_RadioButton.IsChecked == true)
            {
                m_Stamp = 1;
                if (m_ImportedFile != null && m_OriginalDataTable != null && timeUnitTextBox != null)
                {
                    timeUnitTextBox.IsReadOnly = true;
                }
            }
            else if (time_RadioButton.IsChecked == true)
            {
                if (m_ImportedFile != null && m_OriginalDataTable != null && timeUnitTextBox != null)
                {
                    timeUnitTextBox.IsReadOnly = false;
                }
            }
        }

        private void saveTimeUnitButton_Click(object sender, RoutedEventArgs e)
        {
            if (m_ImportedFile != null && m_OriginalDataTable != null && m_CustomDataTable != null)
            {
                var getTimeStamp = timeUnitTextBox.Text;

                if (sample_RadioButton.IsChecked == true)
                {
                    getTimeStamp = "1";
                }

                if (double.TryParse(getTimeStamp, out double result))
                {
                    m_Stamp = result;
                    invalidvalueTextBlock.Visibility = Visibility.Hidden;
                }
                else
                {
                    invalidvalueTextBlock.Visibility = Visibility.Visible;
                    invalidvalueTextBlock.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 0, 0));
                }

                UpdateChart(m_ImportedFile, m_OriginalDataTable);
                UpdateCustomChart(m_ImportedFile, m_CustomDataTable);
            }
        }
    }
}