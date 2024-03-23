using ScottPlot.WPF;
using System.Data;
using Szakdolgozat.Model;

namespace Szakdolgozat.Interfaces
{
    internal interface IChartControl
    {
        /// <summary>
        /// Gets or sets the stamp value used to scale the x values in the charts.
        /// </summary>
        double Stamp { get; set; }

        /// <summary>
        /// Updates the charts in the MainWindow with data from an ImportedFile and DataTables.
        /// </summary>
        /// <param name="importedFile">The ImportedFile containing the data to be displayed in the charts.</param>
        /// <param name="originalDataTable">The DataTable containing the original data to be displayed in the original chart.</param>
        /// <param name="customDataTable">The DataTable containing the custom data to be displayed in the custom chart.</param>
        void UpdateCharts(ImportedFile importedFile, DataTable originalDataTable, DataTable customDataTable);

        /// <summary>
        /// Updates the color of the specified series in the original and custom charts.
        /// </summary>
        /// <param name="seriesIndex">The index of the scatter series whose color is to be updated.</param>
        /// <param name="colorInHex">The new color for the scatter series, specified as a hexadecimal string.</param>
        void UpdateChartColor(int seriesIndex, string colorInHex);

        /// <summary>
        /// Auto-scales the axes of the given chart, shows the legend, and refreshes the chart.
        /// </summary>
        /// <param name="chart">The chart to auto-scale and refresh.</param>
        void AutoScaleAndRefreshChart(WpfPlot chart);
    }
}
