using OfficeOpenXml;
using System.Windows;

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
    }
}
