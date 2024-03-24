using System;
using System.IO;
using System.Windows;
using ICSharpCode.SharpZipLib.Zip;
using OfficeOpenXml;
using Szakdolgozat.Interfaces;
using Szakdolgozat.Model;
using Szakdolgozat.Properties;

namespace Szakdolgozat.ViewModel
{
    internal class ExportControl : IExportControl
    {
        /// <summary>
        /// The main window of the application.
        /// </summary>
        private readonly MainWindow m_MainWindow;

        /// <summary>
        /// The ImportControl instance used for importing projects from EDF files.
        /// </summary>
        private readonly ImportControl m_ImportControl;

        /// <summary>
        /// Initializes a new instance of the ExportControl class.
        /// </summary>
        /// <param name="mainWindow">The main window of the application.</param>
        public ExportControl(MainWindow mainWindow, ImportControl importControl)
        {
            m_MainWindow = mainWindow;
            m_ImportControl = importControl;
        }

        /// <inheritdoc cref="IExportControl.ExportProject(string)"/>
        public void ExportProject(string edfFilePath)
        {
            try
            {
                string tempDirectory = m_ImportControl.CreateTemporaryDirectory();

                ExportFilesToTemporaryDirectory(tempDirectory);

                CreateEdfFile(edfFilePath, tempDirectory);

                Directory.Delete(tempDirectory, true);

                MessageBox.Show(Resources.SaveEDF, Resources.Success, MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while exporting files: {ex.Message}", Resources.Error, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Exports the selected files to a temporary directory.
        /// Each selected file is exported to an Excel file in the temporary directory using the ExportToExcel method.
        /// </summary>
        /// <param name="tempDirectory">The path of the temporary directory to which to export the files.</param>
        private void ExportFilesToTemporaryDirectory(string tempDirectory)
        {
            foreach (ImportedFile selectedFile in m_MainWindow.SelectedFiles)
            {
                ExportToExcel(selectedFile, tempDirectory);
            }
        }

        /// <summary>
        /// Creates an EDF file from the files in a temporary directory.
        /// </summary>
        /// <param name="edfFileName">The name of the EDF file to create.</param>
        /// <param name="tempDirectory">The path of the temporary directory containing the files to include in the EDF file.</param>
        private void CreateEdfFile(string edfFileName, string tempDirectory)
        {
            FastZip fastZip = new FastZip();
            fastZip.CreateZip(edfFileName, tempDirectory, true, "");
        }

        /// <summary>
        /// Exports an ImportedFile to an Excel file in a specified directory.
        /// </summary>
        /// <param name="importedFile">The ImportedFile to export.</param>
        /// <param name="outputDirectory">The directory in which to create the Excel file.</param>
        private void ExportToExcel(ImportedFile importedFile, string outputDirectory)
        {
            string outputPath = Path.Combine(outputDirectory, importedFile.FileName + "_customtable.xlsx");

            using (FileStream stream = File.Create(outputPath))
            {
                using (ExcelPackage package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

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
    }
}
