using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows;
using ScottPlot;
using ScottPlot.Plottables;
using ScottPlot.WPF;
using Szakdolgozat.Interfaces;
using Szakdolgozat.Model;

namespace Szakdolgozat.ViewModel
{
    /// <summary>
    /// Manages and updates charts in the MainWindow.
    /// </summary>
    internal class ChartControl : IChartControl
    {
        /// <summary>
        /// The MainWindow instance whose charts are to be updated.
        /// This instance is used to access the charts that need to be updated.
        /// </summary>
        private MainWindow m_MainWindow;

        /// <summary>
        /// The stamp value used to scale the x values in the charts.
        /// </summary>
        private double m_Stamp = 1.0;

        /// <inheritdoc cref="IChartControl.Stamp"/>
        public double Stamp
        {
            get => m_Stamp;
            set => m_Stamp = value;
        }

        /// <summary>
        /// Public constructor that initializes a new instance of the ChartControl class.
        /// </summary>
        public ChartControl(MainWindow mainWindow)
        {
            m_MainWindow = mainWindow;
        }

        /// <inheritdoc cref="IChartControl.UpdateCharts"/>
        public void UpdateCharts(ImportedFile importedFile, DataTable originalDataTable, DataTable customDataTable)
        {
            UpdateChart(importedFile, originalDataTable);
            UpdateCustomChart(importedFile, customDataTable);
        }

        /// <inheritdoc cref="IChartControl.UpdateChartColor"/>
        public void UpdateChartColor(int seriesIndex, string colorInHex)
        {
            if (m_MainWindow.originalChart != null)
            {
                List<Scatter> scatterSeriesList = GetScatterSeriesList(m_MainWindow.originalChart);
                List<Scatter> customScatterSeriesList = GetScatterSeriesList(m_MainWindow.CustomChart);

                if (seriesIndex >= 0 && seriesIndex < scatterSeriesList.Count && seriesIndex < customScatterSeriesList.Count)
                {
                    scatterSeriesList[seriesIndex].LineStyle.Color = Color.FromHex(colorInHex);
                    customScatterSeriesList[seriesIndex].LineStyle.Color = Color.FromHex(colorInHex);
                    m_MainWindow.originalChart.Refresh();
                    m_MainWindow.CustomChart.Refresh();
                }
            }
        }

        /// <inheritdoc cref="IChartControl.AutoScaleAndRefreshChart"/>
        public void AutoScaleAndRefreshChart(WpfPlot chart)
        {
            chart.Plot.Axes.AutoScaleX();
            chart.Plot.Axes.AutoScaleY();
            chart.Plot.ShowLegend();
            chart.Refresh();
        }

        /// <summary>
        /// Updates the original chart with data from an ImportedFile and a DataTable.
        /// </summary>
        /// <param name="importedFile">The ImportedFile containing the data to be displayed in the chart.</param>
        /// <param name="dataTable">The DataTable containing the data to be displayed in the chart.</param>
        private void UpdateChart(ImportedFile importedFile, DataTable dataTable)
        {
            //Clear for other call
            m_MainWindow.originalChart.Plot.Clear();

            //Chart Titles
            SetChartTitles(m_MainWindow.originalChart, importedFile);

            bool invalidValues = false;

            if (dataTable.Rows.Count > 0)
            {
                DataColumn xColumn = dataTable.Columns[0];
                DataColumn eventColumn = dataTable.Columns[1];
                DataColumn[] yColumns = dataTable.Columns.Cast<DataColumn>().Skip(2).ToArray();

                double[] x = GetXValues(dataTable, xColumn);
                object[] events = GetEventValues(dataTable, eventColumn);

                foreach (DataColumn yColumn in yColumns)
                {
                    object[] y = GetYValues(dataTable, yColumn, ref invalidValues);
                    AddScatterSeriesToChart(m_MainWindow.originalChart, x, y, yColumn.ColumnName);
                }

                AddEventLineToChart(m_MainWindow.originalChart, x, events, false);
            }

            AutoScaleAndRefreshChart(m_MainWindow.originalChart);
        }

        /// <summary>
        /// Updates the custom chart with data from an ImportedFile and a DataTable.
        /// </summary>
        /// <param name="importedFile">The ImportedFile containing the data to be displayed in the chart.</param>
        /// <param name="dataTable">The DataTable containing the data to be displayed in the chart.</param>
        private void UpdateCustomChart(ImportedFile importedFile, DataTable dataTable)
        {
            //Clear for other call
            m_MainWindow.CustomChart.Plot.Clear();

            //Chart Titles
            SetChartTitles(m_MainWindow.CustomChart, importedFile);

            bool invalidValues = false;

            if (dataTable.Rows.Count > 0)
            {
                DataColumn xColumn = dataTable.Columns[0];
                DataColumn eventColumn = dataTable.Columns[1];
                DataColumn[] yColumns = dataTable.Columns.Cast<DataColumn>().Skip(2).ToArray();

                double[] x = GetXValues(dataTable, xColumn);
                object[] events = GetEventValues(dataTable, eventColumn);

                foreach (DataColumn yColumn in yColumns)
                {
                    object[] y = GetYValues(dataTable, yColumn, ref invalidValues);
                    AddScatterSeriesToChart(m_MainWindow.CustomChart, x, y, yColumn.ColumnName);
                }

                AddEventLineToChart(m_MainWindow.CustomChart, x, events, true);
            }

            if (invalidValues)
            {
                MessageBox.Show("Excel contains invalid values! Invalid values are set to 0.", "Import warning!", MessageBoxButton.OK, MessageBoxImage.Warning);
            }

            AutoScaleAndRefreshChart(m_MainWindow.CustomChart);
        }

        /// <summary>
        /// Sets the title and labels of the axes for the given chart.
        /// </summary>
        /// <param name="chart">The chart for which to set the title and labels.</param>
        /// <param name="importedFile">The ImportedFile whose file name is used as the chart title.</param>
        private void SetChartTitles(WpfPlot chart, ImportedFile importedFile)
        {
            chart.Plot.Title($"{importedFile.FileName}");

            if (m_MainWindow.sample_RadioButton.IsChecked == true)
            {
                chart.Plot.XLabel("Samples");
            }
            else if (m_MainWindow.time_RadioButton.IsChecked == true)
            {
                chart.Plot.XLabel("Time (s)");
            }

            chart.Plot.YLabel("Values");
        }

        /// <summary>
        /// Gets the x values from a DataTable.
        /// </summary>
        /// <param name="dataTable">The DataTable from which to get the x values.</param>
        /// <param name="xColumn">The DataColumn that contains the x values.</param>
        /// <returns>An array of the x values.</returns>
        private double[] GetXValues(DataTable dataTable, DataColumn xColumn)
        {
            return dataTable.AsEnumerable().Select(row =>
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
        }

        /// <summary>
        /// Gets the event values from a DataTable.
        /// </summary>
        /// <param name="dataTable">The DataTable from which to get the event values.</param>
        /// <param name="eventColumn">The DataColumn that contains the event values.</param>
        /// <returns>An array of the event values.</returns>
        private object[] GetEventValues(DataTable dataTable, DataColumn eventColumn)
        {
            return dataTable.AsEnumerable().Select(row =>
            {
                object eventValue = row[eventColumn];
                return eventValue != DBNull.Value ? eventValue : 0.0;
            }).ToArray();
        }

        /// <summary>
        /// Gets the y values from a DataTable.
        /// </summary>
        /// <param name="dataTable">The DataTable from which to get the y values.</param>
        /// <param name="yColumn">The DataColumn that contains the y values.</param>
        /// <param name="invalidValues">A reference to a boolean that is set to true if any invalid values are encountered.</param>
        /// <returns>An array of the y values.</returns>
        private object[] GetYValues(DataTable dataTable, DataColumn yColumn, ref bool invalidValues)
        {
            bool localInvalidValues = false;

            object[] yValues = dataTable.AsEnumerable().Select(row =>
            {
                object yValue = row[yColumn];
                if (yValue != DBNull.Value)
                {
                    if (double.TryParse(yValue.ToString(), out double parsedValue))
                    {
                        return (object)parsedValue;
                    }
                    else if (yValue.ToString().EndsWith("Event"))
                    {
                        return yValue;
                    }
                    else
                    {
                        // Logic for invalid values
                        localInvalidValues = true;
                        row[yColumn] = 0.0;
                        return (object)0.0;
                    }
                }
                else
                {
                    // Logic for empty values
                    localInvalidValues = true;
                    row[yColumn] = 0.0;
                    return (object)0.0;
                }
            }).ToArray();

            invalidValues = localInvalidValues;
            return yValues;
        }

        /// <summary>
        /// Adds a scatter series to the given chart with the specified x and y values and column name.
        /// </summary>
        /// <param name="chart">The chart to which to add the scatter series.</param>
        /// <param name="x">The x values of the scatter series.</param>
        /// <param name="y">The y values of the scatter series.</param>
        /// <param name="columnName">The name of the column that contains the y values.
        ///     This is used as the label of the scatter series.</param>
        private void AddScatterSeriesToChart(WpfPlot chart, double[] x, object[] y, string columnName)
        {
            Scatter currentLine = chart.Plot.Add.Scatter(x, y);
            currentLine.Label = columnName;

            if (m_MainWindow.DarkModeToggleButton.IsChecked == true)
            {
                currentLine.Color = Color.FromHex("#ffffff");
            }
            else if (m_MainWindow.DarkModeToggleButton.IsChecked == false)
            {
                currentLine.Color = Color.FromHex("#000000");
            }

            currentLine.MarkerStyle.IsVisible = false;
        }

        /// <summary>
        /// Adds an event line to the given chart with the specified x values and event values.
        /// </summary>
        /// <param name="chart">The chart to which to add the event line.</param>
        /// <param name="x">The x values of the event line.</param>
        /// <param name="events">The events to be added to the chart.</param>
        /// <param name="isCustomChart">A boolean indicating whether the chart is a custom chart.</param>
        private void AddEventLineToChart(WpfPlot chart, double[] x, object[] events, bool isCustomChart)
        {
            Scatter eventLine = CreateEventLine(chart, x);

            for (int i = 0; i < events.Length; i++)
            {
                Marker marker = chart.Plot.Add.Marker(x[i], 0);
                if (events[i].ToString() == "Alarm_Event")
                {
                    AddAlarmEvent(chart, x[i], marker, isCustomChart);
                }
                else if (events[i].ToString() == "Error_Event")
                {
                    AddErrorEvent(chart, x[i], marker, isCustomChart);
                }
                else
                {
                    marker.MarkerStyle.IsVisible = false;
                }
            }
        }

        /// <summary>
        /// Adds an alarm event to the given chart at the specified x index.
        /// </summary>
        /// <param name="chart">The chart to which to add the alarm event.</param>
        /// <param name="xIndex">The x index at which to add the alarm event.</param>
        /// <param name="marker">The marker that represents the alarm event.</param>
        /// <param name="isCustomChart">A boolean indicating whether the chart is a custom chart.</param>
        private void AddAlarmEvent(WpfPlot chart, double xIndex, Marker marker, bool isCustomChart)
        {
            if ((isCustomChart && m_MainWindow.enabledCustomEventLine.IsChecked == true) ||
                (!isCustomChart && m_MainWindow.enabledOriginalEventLine.IsChecked == true))
            {
                VerticalLine alarmEventLine = chart.Plot.Add.VerticalLine(xIndex);
                alarmEventLine.Text = "Alarm Event";
                alarmEventLine.LabelOppositeAxis = true;
                alarmEventLine.LineWidth = 1;
                alarmEventLine.Color = Color.FromARGB(4278190219);
            }

            marker.MarkerStyle.Fill.Color = Color.FromARGB(4278190219); // Dark Blue
        }

        /// <summary>
        /// Adds an error event to the given chart at the specified x index.
        /// </summary>
        /// <param name="chart">The chart to which to add the error event.</param>
        /// <param name="xIndex">The x index at which to add the error event.</param>
        /// <param name="marker">The marker that represents the error event.</param>
        /// <param name="isCustomChart">A boolean indicating whether the chart is a custom chart.</param>
        private void AddErrorEvent(WpfPlot chart, double xIndex, Marker marker, bool isCustomChart)
        {
            if ((isCustomChart && m_MainWindow.enabledCustomEventLine.IsChecked == true) ||
                (!isCustomChart && m_MainWindow.enabledOriginalEventLine.IsChecked == true))
            {
                VerticalLine errorEventLine = chart.Plot.Add.VerticalLine(xIndex);
                errorEventLine.Text = "Error Event";
                errorEventLine.LabelOppositeAxis = true;
                errorEventLine.LineWidth = 1;
                errorEventLine.Color = Color.FromARGB(4294901760);
            }

            marker.MarkerStyle.Shape = MarkerShape.FilledTriangleUp;
            marker.MarkerStyle.Fill.Color = Color.FromARGB(4294901760); // Red
        }

        /// <summary>
        /// Creates an event line for the given chart with the specified x values.
        /// </summary>
        /// <param name="chart">The chart on which to create the event line.</param>
        /// <param name="x">The x values of the event line.</param>
        /// <returns>The created event line as a Scatter object.</returns>
        private Scatter CreateEventLine(WpfPlot chart, double[] x)
        {
            double[] y = new double[x.Length];
            Scatter eventLine = chart.Plot.Add.Scatter(x, y);
            eventLine.Label = "Event Line";
            eventLine.MarkerStyle.IsVisible = false;
            eventLine.Color = m_MainWindow.DarkModeToggleButton.IsChecked == true ? Color.FromHex("#ffffff") : Color.FromHex("#000000");
            return eventLine;
        }

        /// <summary>
        /// Gets the scatter series from the specified chart.
        /// </summary>
        /// <param name="chart">The chart from which to get the scatter series.</param>
        /// <returns>A list of the scatter series in the chart.</returns>
        private List<Scatter> GetScatterSeriesList(WpfPlot chart)
        {
            return chart.Plot.GetPlottables()
                .Where(p => p is Scatter scatter)
                .Cast<Scatter>()
                .ToList();
        }
    }
}
