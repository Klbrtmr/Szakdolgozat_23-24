using System;
using System.Data;
using System.IO;
using System.Windows;
using ExcelDataReader;
using ICSharpCode.SharpZipLib.Zip;
using Szakdolgozat.Converters;
using Szakdolgozat.Helper;
using Szakdolgozat.Interfaces;
using Szakdolgozat.Model;
using Szakdolgozat.Properties;

namespace Szakdolgozat.ViewModel
{
    /// <summary>
    /// Handles the import of projects from EDF files.
    /// </summary>
    internal class ImportControl : IImportControl
    {
        /// <summary>
        /// The main window of the application.
        /// </summary>
        private MainWindow m_MainWindow;

        private IFileHandler m_FileHandler;

        private ColorGeneratorHelper m_ColorGenerator;

        private ITemporaryDirectoryHelper m_TemporaryDirectoryHelper = new TemporaryDirectoryHelper();

        /// <summary>
        /// The number of files imported from the current EDF file.
        /// </summary>
        private int m_importedFileNumber;

        /// <summary>
        /// Initializes a new instance of the ImportControl class.
        /// </summary>
        /// <param name="mainWindow">The main window of the application.</param>
        public ImportControl(MainWindow mainWindow, IFileHandler fileHandler, ColorGeneratorHelper colorGenerator)
        {
            m_MainWindow = mainWindow;
            m_FileHandler = fileHandler;
            m_ColorGenerator = colorGenerator;
        }

        /// <inheritdoc cref="IImportControl.ImportFileNumber"/>
        public int ImportFileNumber
        {
            get => m_importedFileNumber;
            set => m_importedFileNumber = value;
        }

        /// <inheritdoc cref="IImportControl.ImportProject(string)"/>
        public void ImportProject(string edfFilePath)
        {
            try
            {
                string tempDirectory = m_TemporaryDirectoryHelper.CreateTemporaryDirectory();
                m_importedFileNumber = 0;

                ExtractFilesFromEdfToTemporaryDirectory(edfFilePath, tempDirectory);

                ImportExcelFilesFromDirectory(tempDirectory);

                Directory.Delete(tempDirectory, true);

                ShowImportResultMessage();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while importing files: {ex.Message}", Resources.Error, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        /*
        /// <inheritdoc cref="IImportControl.CreateTemporaryDirectory"/>
        public string CreateTemporaryDirectory()
        {
            string tempDirectory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempDirectory);
            return tempDirectory;
        }*/

        /// <summary>
        /// Extracts files from an EDF file to a temporary directory.
        /// </summary>
        /// <param name="edfFilePath">The path of the EDF file.</param>
        /// <param name="tempDirectory">The path of the temporary directory.</param>
        private void ExtractFilesFromEdfToTemporaryDirectory(string edfFilePath, string tempDirectory)
        {
            using (ZipInputStream zipInputStream = new ZipInputStream(File.OpenRead(edfFilePath)))
            {
                ZipEntry entry;
                while ((entry = zipInputStream.GetNextEntry()) != null)
                {
                    string entryPath = Path.Combine(tempDirectory, entry.Name);

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
        }

        /// <summary>
        /// Imports all valid Excel files from a directory.
        /// </summary>
        /// <param name="directoryPath">The path of the directory from which to import the Excel files.</param>
        private void ImportExcelFilesFromDirectory(string directoryPath)
        {
            foreach (string excelFile in Directory.GetFiles(directoryPath, "*.xlsx"))
            {
                if (IsValidExcelFile(excelFile))
                {
                    ImportExcelFile(excelFile);
                }
                else
                {
                    MessageBox.Show($"Error: {excelFile} is not a valid Excel file.", Resources.Error, MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        /// <summary>
        /// Checks if a file is a valid Excel file.
        /// </summary>
        /// <param name="filePath">The path of the file to check.</param>
        /// <returns>True if the file is a valid Excel file, false otherwise.</returns>
        private bool IsValidExcelFile(string filePath)
        {
            try
            {
                using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
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

        /// <summary>
        /// Shows a message box with the result of the import operation.
        /// </summary>
        private void ShowImportResultMessage()
        {
            if (m_importedFileNumber == 0)
            {
                MessageBox.Show("EDF file is not valid. 0 file imported.", Resources.Error, MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else if (m_importedFileNumber == 1)
            {
                MessageBox.Show("File imported successfully.", Resources.Success, MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else if (m_importedFileNumber > 1)
            {
                MessageBox.Show("Files imported successfully.", Resources.Success, MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        /// <summary>
        /// Import excel file from edf package file.
        /// </summary>
        /// <param name="excelFilePath"></param>
        private void ImportExcelFile(string excelFilePath)
        {
            try
            {
                using (FileStream streamval = File.Open(excelFilePath, FileMode.Open, FileAccess.Read))
                {
                    using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(streamval))
                    {
                        DataSet dataSet = m_FileHandler.GetDataSetFromReader(reader);

                        if (!m_FileHandler.IsValidDataSet(dataSet))
                        {
                            MessageBox.Show("Error: Wrong file structure or file is corrupt.", Resources.Error, MessageBoxButton.OK, MessageBoxImage.Error);
                            return;
                        }

                        m_MainWindow.m_CellValues = m_FileHandler.GetCellValuesFromDataSet(dataSet);
                        m_MainWindow.m_CustomCellValues = (object[,])m_MainWindow.m_CellValues.Clone();

                        ImportedFile importedFile = m_FileHandler.CreateImportedFile(
                            excelFilePath,
                            m_MainWindow.m_CellValues,
                            m_MainWindow.m_CustomCellValues,
                            m_FileHandler.GenerateNewID(),
                            m_ColorGenerator.GetDisplayColorForFile(excelFilePath),
                            m_MainWindow.SelectedFiles);

                        m_MainWindow.SelectedFiles.Add(importedFile);
                        ImportFileNumber++;
                    }
                }
                m_MainWindow.ListFiles();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while importing the file: {ex.Message}", Resources.Error, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

    }
}
