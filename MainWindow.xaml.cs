using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using OfficeOpenXml;
using Szakdolgozat.Converters;
using Szakdolgozat.Helper;
using Szakdolgozat.Model;
using Szakdolgozat.View;
using Szakdolgozat.ViewModel;
using ColorConverter = Szakdolgozat.Converters.ColorConverter;

namespace Szakdolgozat
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private ColorGeneratorHelper m_ColorGenerator;
        private ColorConverter m_ColorConverter;
        private FileHandler m_FileHandler;
        private ChildParentHelper m_ChildParentHelper;
        private UIHelper m_UIHelper;
        private UIColorUpdate m_UIColorUpdate;
        private ChartControl m_ChartControl;
        private ImportControl m_ImportControl;
        private ExportControl m_ExportControl;

        /// <summary>
        /// Initializes a new instance of the MainWindow class.
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
            importProgressBar = FindName("importProgressBar") as ProgressBar;

            m_ColorGenerator = new ColorGeneratorHelper(this);
            m_ColorConverter = new ColorConverter();
            m_FileHandler = new FileHandler(this);
            m_ChildParentHelper = new ChildParentHelper();
            m_UIHelper = new UIHelper(this);
            m_UIColorUpdate = new UIColorUpdate(this, m_ChildParentHelper);
            m_ChartControl = new ChartControl(this);
            m_ImportControl = new ImportControl(this, m_FileHandler, m_ColorGenerator);
            m_ExportControl = new ExportControl(this);
        }

        /// <summary>
        /// The list in which the imported files are stored.
        /// </summary>
        private List<ImportedFile> m_SelectedFiles = new List<ImportedFile>();

        public List<ImportedFile> SelectedFiles
        {
            get => m_SelectedFiles;
        }

        /// <summary>
        /// Two-dimension array for original excel datas.
        /// </summary>
        public object[,] m_CellValues { get; set; }

        /// <summary>
        /// Two-dimension array for custom changed excel datas.
        /// </summary>
        public object[,] m_CustomCellValues { get; set; }

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
        /// Selected values for color of combobox.
        /// </summary>
        private List<object> selectedValues = new List<object>();

        /// <summary>
        /// Listed all file what we imported. This method created an ellipse to every file.
        /// </summary>
        public void ListFiles()
        {
            filesListing.ItemsSource = null;
            filesListing.Items.Clear();

            foreach (ImportedFile importedFile in m_SelectedFiles)
            {
                StackPanel panel = m_UIHelper.CreateFilePanel(importedFile);
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
                string newFileName = m_FileHandler.GenerateUniqueFileName(excelFilePath, m_SelectedFiles);

                try
                {
                    importProgressBar.Visibility = Visibility.Visible;


                    m_CellValues = await m_FileHandler.ReadExcelFile(excelFilePath);
                    m_CustomCellValues = (object[,])m_CellValues.Clone();
                    IDictionary<double, string> localNamedValues = m_FileHandler.GetValuesForNamedValues(m_CellValues);

                    ImportedFile importedFile = m_FileHandler.CreateSingleImportedFile(
                        newFileName,
                        excelFilePath,
                        m_CellValues,
                        m_CustomCellValues,
                        m_FileHandler.GenerateNewID(),
                        m_ColorGenerator.GetDisplayColorForFile(newFileName),
                        localNamedValues);
                    Dispatcher.Invoke(() =>
                    {
                        m_SelectedFiles.Add(importedFile);
                        ListFiles();
                    });

                     importProgressBar.Visibility = Visibility.Collapsed;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    importProgressBar.Visibility = Visibility.Collapsed;
                }
            }
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

                    ImportedFile m_selectedImportedFile = m_SelectedFiles.FirstOrDefault(file => file.FileName == selectedFileName);

                    if (m_selectedImportedFile != null)
                    {
                        m_SelectedFiles.Remove(m_selectedImportedFile);
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
            e.Handled = filesListing?.SelectedItem == null;
        }

        /// <summary>
        /// Double click to file and opened it.
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

                    ImportedFile m_selectedImportedFile = m_SelectedFiles.FirstOrDefault(file => file.FileName.Equals(selectedFileName));

                    if (m_selectedImportedFile != null)
                    {
                        m_ImportedFile = m_selectedImportedFile;
                        tabControlBorder.BorderBrush = new SolidColorBrush(m_selectedImportedFile.DisplayColor);
                        DisplayExcelData(m_selectedImportedFile);
                    }
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

            m_CustomDataTable = ConvertArrayToDataTable(importedFile.CustomExcelData);
            excelCustomDataGrid.ItemsSource = m_CustomDataTable.DefaultView;

            DataTable configTable = ConvertArrayToDataTableConfigPage(importedFile.ExcelData);
            configGrid.ItemsSource = configTable.DefaultView;

            DataTable namedValuesDataTable = ConvertDictionaryToDataTable(importedFile.NamedValues);
            namedValuesDataGrid.ItemsSource = namedValuesDataTable.DefaultView;

            this.m_ImportedFile = importedFile;
            m_ChartControl.UpdateCharts(m_ImportedFile, m_OriginalDataTable, m_CustomDataTable);
        }

        /// <summary>
        /// Convert from the imported file datas to datatable for an easier handle.
        /// </summary>
        /// <param name="array">Two-dimension array from imported file.</param>
        /// <returns>Returns the datatable from excel data.</returns>
        private DataTable ConvertArrayToDataTable(object[,] array)
        {
            DataTable dataTable = new DataTable();

            selectedValues.Clear();

            for (int i = 0; i < array.GetLength(1); i++)
            {
                dataTable.Columns.Add($"{array[0, i]}");

                if (i>=2)
                {
                    selectedValues.Add((object)"Black");
                }
            }

            for (int i = 1; i < array.GetLength(0); i++)
            {
                DataRow dataRow = dataTable.NewRow();

                for (int j = 0; j < array.GetLength(1); j++)
                {
                    object actualValue = array[i, j];
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

        /// <summary>
        /// Convert from the imported file datas to datatable for config page.
        /// </summary>
        /// <param name="array"></param>
        /// <returns>Returns with config page grid for signal color choice</returns>
        private DataTable ConvertArrayToDataTableConfigPage(object[,] array)
        {
            DataTable dataTable = new DataTable();

            dataTable.Columns.Add("Signal Name", typeof(string));
            dataTable.Columns.Add("Color Preview", typeof(Brush));

            DataGridTemplateColumn templateColumn = new DataGridTemplateColumn();
            templateColumn.Header = "Colors";

            // Create DataTemplate, which contains the ComboBox
            FrameworkElementFactory comboBoxFactory = new FrameworkElementFactory(typeof(ComboBox));
            comboBoxFactory.SetValue(ComboBox.ItemsSourceProperty,m_ColorConverter.colorList);
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
        /// Convert from the dictionary to datatable for an easier handle.
        /// </summary>
        /// <param name="dictionary"></param>
        /// <returns></returns>
        private DataTable ConvertDictionaryToDataTable(IDictionary<double, string> dictionary)
        {
            DataTable dataTable = new DataTable();

            dataTable.Columns.Add("Values", typeof(double));
            dataTable.Columns.Add("Named Values", typeof(string));

            foreach (var kvp in dictionary)
            {
                dataTable.Rows.Add(kvp.Key, kvp.Value);
            }

            return dataTable;
        }

        /// <summary>
        /// Handle when in the custom data grid edited cell(s). Save the new value to the m_CustomCellValues array.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void excelCustomDataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.Row.Item is DataRowView rowView)
            {
                if (e.Column.DisplayIndex == 0)
                {
                    HandleFirstColumnEdit(e, rowView);
                }
                else
                {
                    HandleOtherColumnsEdit(e, rowView);
                }
            }
        }

        private void HandleFirstColumnEdit(DataGridCellEditEndingEventArgs e, DataRowView rowView)
        {
            ((TextBox)e.EditingElement).Text = rowView.Row[e.Column.DisplayIndex].ToString();
            MessageBox.Show("Invalid value!", "Invalid value error!", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        private void HandleOtherColumnsEdit(DataGridCellEditEndingEventArgs e, DataRowView rowView)
        {
            object modifiedValue = ((TextBox)e.EditingElement).Text;

            if (modifiedValue != DBNull.Value)
            {
                if (double.TryParse(modifiedValue.ToString(), out double parsedValue) ||
                    (modifiedValue.ToString().EndsWith("Event") && e.Column.DisplayIndex == 1))
                {
                    UpdateRowViewAndCustomExcelData(e, rowView, modifiedValue, parsedValue);
                }
                else
                {
                    HandleInvalidValue(e, rowView);
                }
            }
        }

        private void UpdateRowViewAndCustomExcelData(DataGridCellEditEndingEventArgs e, DataRowView rowView, object modifiedValue, double parsedValue)
        {
            rowView.Row[e.Column.DisplayIndex] = parsedValue != 0 ? parsedValue : modifiedValue;

            int rowIndex = e.Row.GetIndex() + 1;
            int columnIndex = e.Column.DisplayIndex;

            if (m_ImportedFile != null)
            {
                m_ImportedFile.CustomExcelData[rowIndex, columnIndex] = modifiedValue.ToString().EndsWith("Event") ? modifiedValue : parsedValue;
                m_ChartControl.UpdateCharts(m_ImportedFile, m_OriginalDataTable, m_CustomDataTable);
                SetChartColor();
            }
        }

        private void HandleInvalidValue(DataGridCellEditEndingEventArgs e, DataRowView rowView)
        {
            MessageBox.Show("Invalid value!", "Invalid value error!", MessageBoxButton.OK, MessageBoxImage.Error);
            ((TextBox)e.EditingElement).Text = rowView.Row[e.Column.DisplayIndex].ToString();
        }

        /// <summary>
        /// Exported custom data table to the new excel file.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void exporttoexcel_Click(object sender, RoutedEventArgs e)
        {
            if (m_ImportedFile == null)
            {
                return;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel file (*.xlsx)|*.xlsx",
                Title = "Save",
                FileName = Path.GetFileNameWithoutExtension(m_ImportedFile.FileName) + "_customtable.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                string outputPath = saveFileDialog.FileName;

                using (FileStream stream = File.Create(outputPath))
                {
                    using (ExcelPackage package = new ExcelPackage(stream))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

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
            MessageBoxResult result = MessageBox.Show("Are you sure?", "Close", MessageBoxButton.OKCancel, MessageBoxImage.Hand);
            if (result == MessageBoxResult.OK)
            {
                Application.Current.Shutdown();
            }
        }

        /// <summary>
        /// Export project to special '.EDF' file.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void exportproject_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "EDF file (*.edf)|*.edf";

            if (saveFileDialog.ShowDialog() == true)
            {
                m_ExportControl.ExportProject(saveFileDialog.FileName);
            }
        }

        /// <summary>
        /// Import project from special '.EDF' file.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void importProject_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "EDF file (*.edf)|*.edf";

            if (openFileDialog.ShowDialog() == true)
            {
                m_ImportControl.ImportProject(openFileDialog.FileName);
            }
        }

        /// <summary>
        /// Double mouse click to centered view settings.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void myPlot_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            m_ChartControl.AutoScaleAndRefreshChart(originalChart);
            m_ChartControl.AutoScaleAndRefreshChart(CustomChart);
        }

        /// <summary>
        /// Switch between dark and light mode.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DarkModeToggleButton_Click(object sender, RoutedEventArgs e)
        {
            if (DarkModeToggleButton.IsChecked == true)
            {
                m_UIColorUpdate.EnableDarkMode(m_UIHelper);
            }
            else if (DarkModeToggleButton.IsChecked == false)
            {
                m_UIColorUpdate.DisableDarkMode(m_UIHelper);
            }
        }

        public void RefreshCharts()
        {
            originalChart.Refresh();
            CustomChart.Refresh();
        }

        /// <summary>
        /// Color selection at combobox in the config page.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox comboBox)
            {
                DataGridCell cell = m_ChildParentHelper.FindVisualParent<DataGridCell>(comboBox);
                DataGridRow row = m_ChildParentHelper.FindVisualParent<DataGridRow>(cell);

                if (row != null)
                {
                    object selectedValue = comboBox.SelectedValue;
                    string colorInHex = new ColorConverter().ColorNameToHex(selectedValue);
                    object colorInHex2 = ColorConverterForBrushes.Instance.Convert(selectedValue, typeof(Brush), null, CultureInfo.CurrentCulture);

                    m_UIColorUpdate.UpdateCellBackground(row, colorInHex2 as Brush);
                    selectedValues[row.GetIndex()] = comboBox.SelectedValue;
                }
            }
        }

        /// <summary>
        /// Radio Button change between sample and time modes.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RadioButtonMode_Checked(object sender, RoutedEventArgs e)
        {
            bool isFileAndTableNotNull = IsDataAvailable() && timeUnitTextBox != null;

            if (m_ChartControl == null)
            {
                return;
            }

            if (sample_RadioButton.IsChecked == true)
            {
                m_ChartControl.Stamp = 1;
                if (isFileAndTableNotNull)
                {
                    timeUnitTextBox.IsReadOnly = true;
                }
            }
            else if (time_RadioButton.IsChecked == true && isFileAndTableNotNull)
            {
                timeUnitTextBox.IsReadOnly = false;
            }
        }

        /// <summary>
        /// Save button between sample and time mode.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void saveTimeUnitButton_Click(object sender, RoutedEventArgs e)
        {
            if (IsDataAvailable())
            {
                string getTimeStamp = sample_RadioButton.IsChecked == true ? "1" : timeUnitTextBox.Text;

                if (double.TryParse(getTimeStamp, out double result))
                {
                    m_ChartControl.Stamp = result;
                    invalidvalueTextBlock.Visibility = Visibility.Hidden;
                    timeUnitTextBox.Background = new SolidColorBrush(Color.FromRgb(255, 255, 255));
                    timeUnitTextBox.Foreground = new SolidColorBrush(Color.FromRgb(0, 0, 0));
                }
                else
                {
                    invalidvalueTextBlock.Visibility = Visibility.Visible;
                    invalidvalueTextBlock.Foreground = new SolidColorBrush(Color.FromRgb(255, 0, 0));
                    timeUnitTextBox.Background = new SolidColorBrush(Color.FromRgb(255, 0, 0));
                    timeUnitTextBox.Foreground = new SolidColorBrush(Color.FromRgb(255, 255, 255));
                }

                m_ChartControl.UpdateCharts(m_ImportedFile, m_OriginalDataTable, m_CustomDataTable);

                SetChartColor();
            }
        }

        /// <summary>
        /// I dont know its should be here.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void enabledCustomEventLine_Checked(object sender, RoutedEventArgs e)
        {
        }

        /// <summary>
        /// Save button between show and hide extra lines on custom chart.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void saveCustomEventLines_Click(object sender, RoutedEventArgs e)
        {
            if (IsDataAvailable())
            {
                m_ChartControl.UpdateCharts(m_ImportedFile, m_OriginalDataTable, m_CustomDataTable);
                SetChartColor();
            }
        }

        private void SetChartColor()
        {
            for (int i = 0; i < selectedValues.Count; i++)
            {
                var selectedValue = selectedValues[i];
                if (selectedValue != null)
                {
                    string colorInHex = new ColorConverter().ColorNameToHex(selectedValue);
                    m_ChartControl.UpdateChartColor(i, colorInHex);
                }
            }
        }

        /// <summary>
        /// I dont know its should be here.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void enabledOriginalEventLine_Checked(object sender, RoutedEventArgs e)
        {
        }

        /// <summary>
        /// Save button between show and hide extra lines on original chart.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void saveOriginalEventLines_Click(object sender, RoutedEventArgs e)
        {
            if (IsDataAvailable())
            {
                m_ChartControl.UpdateCharts(m_ImportedFile, m_OriginalDataTable, m_CustomDataTable);
                SetChartColor();
            }
        }

        private bool IsDataAvailable()
        {
            return m_ImportedFile != null && m_OriginalDataTable != null && m_CustomDataTable != null;
        }

        /// <summary>
        /// Handles the event when the user clicks the Clean project button.
        /// This method resets the imported file, data tables, cell values, and charts.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cleanButton_Click(object sender, RoutedEventArgs e)
        {
            m_ImportedFile = null;
            m_OriginalDataTable = null;
            m_CustomDataTable = null;
            m_CellValues = null;
            m_CustomCellValues = null;
            m_SelectedFiles.Clear();
            filesListing.Items.Clear();
            excelDataGrid.ItemsSource = null;
            excelCustomDataGrid.ItemsSource = null;
            configGrid.ItemsSource = null;
            configGrid.Columns.Clear();
            namedValuesDataGrid.ItemsSource = null;
            namedValuesDataGrid.Columns.Clear();
            originalChart.Plot.Title(string.Empty);
            originalChart.Plot.Clear();
            CustomChart.Plot.Title(string.Empty);
            CustomChart.Plot.Clear();
            RefreshCharts();
            tabControlBorder.BorderBrush = new SolidColorBrush(Color.FromRgb(128, 128, 128));
        }

        private void namedValuesDataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.Row.Item is DataRowView rowView)
            {
                if (e.Column.DisplayIndex == 0)
                {
                    HandleFirstColumnEdit(e, rowView);
                }
                else
                {
                    HandleOtherNamedValuesColumnsEdit(e, rowView);
                }
            }
        }

        private void HandleOtherNamedValuesColumnsEdit(DataGridCellEditEndingEventArgs e, DataRowView rowView)
        {
            object modifiedValue = ((TextBox)e.EditingElement).Text;

            if (modifiedValue != DBNull.Value)
            {
                if (e.Column.DisplayIndex == 1)
                {
                    UpdateRowViewAndCustomExcelData2(e, rowView, modifiedValue);
                }
                else
                {
                    HandleInvalidValue(e, rowView);
                }
            }
        }
        
        private void UpdateRowViewAndCustomExcelData2(DataGridCellEditEndingEventArgs e, DataRowView rowView, object modifiedValue)
        {
            rowView.Row[e.Column.DisplayIndex] = modifiedValue;
            object value = rowView.Row[0];

            if (m_ImportedFile != null)
            {
                m_ImportedFile.NamedValues[double.Parse(value.ToString())] = modifiedValue.ToString();
            }
        }

        private void saveNamedValues_Click(object sender, RoutedEventArgs e)
        {
            if (IsDataAvailable())
            {
                m_ChartControl.UpdateCharts(m_ImportedFile, m_OriginalDataTable, m_CustomDataTable);
                SetChartColor();
            }
        }

        private void enabledNamedValues_Checked(object sender, RoutedEventArgs e)
        {
        }
    }
}