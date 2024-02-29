using OfficeOpenXml;
using System;
using System.Windows;
using System.Windows.Controls;

namespace Szakdolgozat
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public App()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            ToolTipService.ShowDurationProperty.OverrideMetadata(
                typeof(DependencyObject),
                new FrameworkPropertyMetadata(Int32.MaxValue) );
        }
    }
}
