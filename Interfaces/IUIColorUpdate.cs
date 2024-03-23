using ScottPlot.WPF;
using System.Windows;
using Szakdolgozat.View;

namespace Szakdolgozat.Interfaces
{
    internal interface IUIColorUpdate
    {
        /// <summary>
        /// Sets the style of the original and custom charts for light mode.
        /// The style includes the colors of the axes, grid, background, and legend.
        /// </summary>
        /// <param name="originalChart">The original chart to be styled.</param>
        /// <param name="customChart">The custom chart to be styled.</param>
        void LightModeForCharts(WpfPlot originalChart, WpfPlot customChart);

        /// <summary>
        /// Sets the style of the given chart for dark mode.
        /// The style includes the colors of the axes, grid, background, and legend.
        /// </summary>
        /// <param name="chart">The chart to be styled.</param>
        void SetChartStyleForDarkMode(WpfPlot chart);

        /// <summary>
        /// Sets the background colors of various UI elements for dark mode.
        /// It also updates the icon images for dark mode.
        /// </summary>
        /// <param name="uiHelper">The UIHelper object used to set the icon images.</param>
        /// <param name="elements">The UI elements whose foreground color is to be set.</param>
        void SetDarkModeBackground(UIHelper uiHelper, params FrameworkElement[] elements);

        /// <summary>
        /// Sets the background colors of various UI elements for light mode.
        /// It also updates the icon images for light mode.
        /// </summary>
        /// <param name="uiHelper">The UIHelper object used to set the icon images.</param>
        /// <param name="elements">The UI elements whose foreground color is to be set.</param>
        void SetLightModeBackground(UIHelper uiHelper, params FrameworkElement[] elements);
    }
}
