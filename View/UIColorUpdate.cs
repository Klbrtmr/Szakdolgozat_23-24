using System.Windows.Media;
using System.Windows;
using System.Windows.Controls;
using ScottPlot.WPF;
using Szakdolgozat.Interfaces;

namespace Szakdolgozat.View
{
    /// <summary>
    /// A class that updates the UI colors of the MainWindow.
    /// It includes methods to set the light mode and dark mode for charts and backgrounds.
    /// It also includes methods to create SolidColorBrush objects and set colors for various UI elements.
    /// </summary>
    internal class UIColorUpdate : IUIColorUpdate
    {
        /// <summary>
        /// The MainWindow whose UI colors are to be updated.
        /// </summary>
        private MainWindow m_MainWindow;

        /// <summary>
        /// Initializes a new instance of the UIColorUpdate class.
        /// </summary>
        /// <param name="mainWindow">The MainWindow whose UI colors are to be updated.</param>
        public UIColorUpdate(MainWindow mainWindow)
        {
            m_MainWindow = mainWindow;
        }

        /// <inheritdoc cref="IUIColorUpdate.LightModeForCharts"/>
        public void LightModeForCharts(WpfPlot originalChart, WpfPlot customChart)
        {
            SetChartStyle(originalChart, "#000000", "#e0e0e0", "#ffffff", "#000000");
            SetChartStyle(customChart, "#000000", "#e0e0e0", "#ffffff", "#000000");
        }

        /// <inheritdoc cref="IUIColorUpdate.SetChartStyleForDarkMode"/>
        public void SetChartStyleForDarkMode(WpfPlot chart)
        {
            chart.Plot.Style.DarkMode();
        }

        /// <inheritdoc cref="IUIColorUpdate.SetDarkModeBackground"/>
        public void SetDarkModeBackground(UIHelper uiHelper, params FrameworkElement[] elements)
        {
            SolidColorBrush darkBackground = CreateSolidColorBrush(40, 40, 40); // dark gray
            SolidColorBrush solidDarkBackground = CreateSolidColorBrush(50, 50, 50); // dark gray
            SolidColorBrush whiteColor = CreateSolidColorBrush(255, 255, 255);

            SetColors(darkBackground, solidDarkBackground, solidDarkBackground, whiteColor, elements);
            uiHelper.SetIconImagesForDarkMode();
        }

        /// <inheritdoc cref="IUIColorUpdate.SetLightModeBackground"/>
        public void SetLightModeBackground(UIHelper uiHelper, params FrameworkElement[] elements)
        {
            SolidColorBrush lightBackground = CreateSolidColorBrush(128, 128, 128); // gray
            SolidColorBrush lighterBackground = CreateSolidColorBrush(169, 169, 169);
            SolidColorBrush blackColor = CreateSolidColorBrush(0, 0, 0);
            SolidColorBrush whiteColor = CreateSolidColorBrush(255, 255, 255);

            SetColors(lightBackground, lighterBackground, whiteColor, blackColor, elements);
            uiHelper.SetIconImagesForLightMode();
        }

        public void EnableDarkMode(UIHelper uiHelper)
        {
            m_MainWindow.DarkModeToggleButton.Content = uiHelper.CreateImage("Assets/brightness.png");

            SetChartStyleForDarkMode(m_MainWindow.originalChart);
            SetChartStyleForDarkMode(m_MainWindow.CustomChart);

            SetDarkModeBackground(uiHelper,
                m_MainWindow.myNameText,
                m_MainWindow.homeButton, m_MainWindow.importExcel, m_MainWindow.importProject,
                m_MainWindow.exporttoexcel, m_MainWindow.exportproject, m_MainWindow.closeProject,
                m_MainWindow.filesListing, m_MainWindow.sample_RadioButton, m_MainWindow.time_RadioButton,
                m_MainWindow.customEventLineTextBlock, m_MainWindow.enabledCustomEventLine, m_MainWindow.disabledCustomEventLine,
                m_MainWindow.originalEventLineTextBlock, m_MainWindow.enabledOriginalEventLine, m_MainWindow.disabledOriginalEventLine);

            m_MainWindow.RefreshCharts();
        }

        public void DisableDarkMode(UIHelper uiHelper)
        {
            m_MainWindow.DarkModeToggleButton.Content = uiHelper.CreateImage("Assets/moon.png");

            LightModeForCharts(m_MainWindow.originalChart, m_MainWindow.CustomChart);
            SetLightModeBackground(uiHelper,
                m_MainWindow.myNameText,
                m_MainWindow.homeButton, m_MainWindow.importExcel, m_MainWindow.importProject,
                m_MainWindow.exporttoexcel, m_MainWindow.exportproject, m_MainWindow.closeProject,
                m_MainWindow.filesListing, m_MainWindow.sample_RadioButton, m_MainWindow.time_RadioButton,
                m_MainWindow.customEventLineTextBlock, m_MainWindow.enabledCustomEventLine, m_MainWindow.disabledCustomEventLine,
                m_MainWindow.originalEventLineTextBlock, m_MainWindow.enabledOriginalEventLine, m_MainWindow.disabledOriginalEventLine);

            m_MainWindow.RefreshCharts();
        }

        /// <summary>
        /// Updates the foreground color of the specified UI elements.
        /// </summary>
        /// <param name="solidColorBrush">The new foreground color as a SolidColorBrush object.</param>
        /// <param name="elements">The UI elements whose foreground color is to be updated.</param>
        private void UpdateForegroundColors(SolidColorBrush solidColorBrush, params FrameworkElement[] elements)
        {
            foreach (var element in elements)
            {
                if (element is Control control)
                {
                    control.Foreground = solidColorBrush;
                }
                else if (element is TextBlock textBlock)
                {
                    textBlock.Foreground = solidColorBrush;
                }
            }
        }

        /// <summary>
        /// Sets the style of the given chart.
        /// </summary>
        /// <param name="chart">The chart to be styled.</param>
        /// <param name="axesColor">The color of the axes as a hexadecimal color code.</param>
        /// <param name="gridColor">The color of the grid as a hexadecimal color code.</param>
        /// <param name="backgroundColor">The color of the background as a hexadecimal color code.</param>
        /// <param name="legendColor">The color of the legend as a hexadecimal color code.</param>
        private void SetChartStyle(WpfPlot chart, string axesColor, string gridColor, string backgroundColor, string legendColor)
        {
            var axesColorObj = ScottPlot.Color.FromHex(axesColor);
            var gridColorObj = ScottPlot.Color.FromHex(gridColor);
            var backgroundColorObj = ScottPlot.Color.FromHex(backgroundColor);
            var legendColorObj = ScottPlot.Color.FromHex(legendColor);

            chart.Plot.Style.ColorAxes(axesColorObj);
            chart.Plot.Style.ColorGrids(gridColorObj);
            chart.Plot.Style.Background(figure: backgroundColorObj, data: backgroundColorObj);
            chart.Plot.Style.ColorLegend(background: backgroundColorObj, foreground: legendColorObj, border: legendColorObj);
        }

        /// <summary>
        /// Creates a SolidColorBrush object from the specified RGB color values.
        /// </summary>
        /// <param name="red">The red component of the color.</param>
        /// <param name="green">The green component of the color.</param>
        /// <param name="blue">The blue component of the color.</param>
        /// <returns>A SolidColorBrush object with the specified color.</returns>
        private SolidColorBrush CreateSolidColorBrush(byte red, byte green, byte blue)
        {
            return new SolidColorBrush(System.Windows.Media.Color.FromRgb(red, green, blue));
        }

        /// <summary>
        /// Sets the background colors of various UI elements in the MainWindow.
        /// </summary>
        /// <param name="background1">The background color for the BackgroundBehindTabs element.</param>
        /// <param name="background2">The background color for the mainTabControl and controllerStackPanel elements.</param>
        /// <param name="background3">The background color for the filesListing element.</param>
        /// <param name="foreground">The foreground color for the specified UI elements.</param>
        /// <param name="elements">The UI elements whose foreground color is to be updated.</param>
        private void SetColors(SolidColorBrush background1, SolidColorBrush background2, SolidColorBrush background3, SolidColorBrush foreground, params FrameworkElement[] elements)
        {
            //m_MainWindow.BackgroundBehindTabs.Background = background1;
            m_MainWindow.BackgroundBehindTabs.Background = background1;
            m_MainWindow.mainTabControl.Background = background2;
            m_MainWindow.controllerStackPanel.Background = background2;
            m_MainWindow.filesListing.Background = background3;

            UpdateForegroundColors(foreground, elements);
        }
    }
}
