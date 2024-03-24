using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ExcelDataReader;
using Szakdolgozat.Interfaces;
using Szakdolgozat.Model;
using Szakdolgozat.Properties;

namespace Szakdolgozat.ViewModel
{
    internal class FileHandler : IFileHandler
    {
        private int currentID = 0;

        /// <inheritdoc cref="IFileHandler.GenerateUniqueFileName"/>
        public string GenerateUniqueFileName(string excelFilePath, List<ImportedFile> selectedFiles)
        {
            string fileName = Path.GetFileNameWithoutExtension(excelFilePath);
            string newFileName = fileName;
            int counter = 1;

            while (selectedFiles.Any(file => file.FileName == newFileName))
            {
                newFileName = $"{fileName}_{counter}";
                counter++;
            }

            return newFileName;
        }

        /// <inheritdoc cref="IFileHandler.ReadExcelFile"/>
        public async Task<object[,]> ReadExcelFile(string excelFilePath)
        {
            object[,] cellValues;

            using (FileStream streamval = File.Open(excelFilePath, FileMode.Open, FileAccess.Read))
            {
                using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(streamval))
                {
                    ExcelDataSetConfiguration configuration = new ExcelDataSetConfiguration
                    {
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration
                        {
                            UseHeaderRow = false
                        }
                    };
                    DataSet dataSet = reader.AsDataSet(configuration);

                    ValidateDataSet(dataSet);

                    cellValues = new object[dataSet.Tables[0].Rows.Count, dataSet.Tables[0].Columns.Count];

                    await Task.Run(() =>
                    {
                        PopulateCellValues(dataSet, cellValues);
                        System.Threading.Thread.Sleep(2000);
                    });
                }
            }

            return cellValues;
        }

        /// <inheritdoc cref="IFileHandler.CreateSingleImportedFile"/>
        public ImportedFile CreateSingleImportedFile(string newFileName, string excelFilePath, object[,] cellValues, object[,] customCellValues, int newID, System.Windows.Media.Color displayColor)
        {
            return new ImportedFile
            {
                ID = newID,
                FileName = newFileName,
                FilePath = excelFilePath,
                DisplayColor = displayColor,
                ExcelData = cellValues,
                CustomExcelData = customCellValues,
            };
        }

        /// <inheritdoc cref="IFileHandler.GenerateNewID"/>
        public int GenerateNewID()
        {
            return currentID++;
        }

        /// <summary>
        /// Validates the structure of a DataSet obtained from an Excel file.
        /// </summary>
        /// <param name="dataSet">The DataSet to validate.</param>
        /// <exception cref="Exception">Exception when file is not correct.</exception>
        private void ValidateDataSet(DataSet dataSet)
        {
            if (dataSet.Tables.Count > 0)
            {
                object firstCellValue = dataSet.Tables[0].Rows[0].ItemArray[0];
                object secondCellValue = dataSet.Tables[0].Rows[0].ItemArray[1];
                if (firstCellValue == null || !firstCellValue.ToString().Equals(Resources.Sample, StringComparison.OrdinalIgnoreCase) &&
                    secondCellValue == null || !secondCellValue.ToString().Equals(Resources.Events, StringComparison.OrdinalIgnoreCase))
                {
                    throw new Exception("Wrong file structure or file is corrupt.");
                }
            }
        }

        /// <summary>
        /// Populates a 2D object array with the cell values from a DataSet.
        /// </summary>
        /// <param name="dataSet">The DataSet from which to get the cell values.</param>
        /// <param name="cellValues">The 2D object array to populate with the cell values.</param>
        private void PopulateCellValues(DataSet dataSet, object[,] cellValues)
        {
            for (int i = 0; i < dataSet.Tables[0].Columns.Count; i++)
            {
                cellValues[0, i] = dataSet.Tables[0].Rows[0].ItemArray[i];
            }

            for (int i = 1; i < dataSet.Tables[0].Rows.Count; i++)
            {
                for (int j = 0; j < dataSet.Tables[0].Columns.Count; j++)
                {
                    object actualValue = dataSet.Tables[0].Rows[i].ItemArray[j];

                    if (double.TryParse(actualValue.ToString(), out double parsedValue))
                    {
                        cellValues[i, j] = parsedValue;
                    }
                    else if (actualValue.ToString().EndsWith(Resources.Event))
                    {
                        cellValues[i, j] = actualValue.ToString();
                    }
                    else
                    {
                        // Logic for invalid values
                        cellValues[i, j] = 0.0;
                    }
                }
            }
        }

        //################################################################################
        //--------------------------------------------------------------------------------
        //################################################################################

        /// <inheritdoc cref="IFileHandler.GetDataSetFromReader"/>
        public DataSet GetDataSetFromReader(IExcelDataReader reader)
        {
            ExcelDataSetConfiguration configuration = new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = false
                }
            };

            return reader.AsDataSet(configuration);
        }

        /// <inheritdoc cref="IFileHandler.IsValidDataSet"/>
        public bool IsValidDataSet(DataSet dataSet)
        {
            if (dataSet.Tables.Count > 0)
            {
                object firstCellValue = dataSet.Tables[0].Rows[0].ItemArray[0];
                object secondCellValue = dataSet.Tables[0].Rows[0].ItemArray[1];
                if (firstCellValue == null || !firstCellValue.ToString().Equals(Resources.Sample, StringComparison.OrdinalIgnoreCase) &&
                    secondCellValue == null || !secondCellValue.ToString().Equals(Resources.Events, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }

            return true;
        }

        /// <inheritdoc cref="IFileHandler.GetCellValuesFromDataSet"/>
        public object[,] GetCellValuesFromDataSet(DataSet dataSet)
        {
            object[,] cellValues = new object[dataSet.Tables[0].Rows.Count, dataSet.Tables[0].Columns.Count];

            for (int i = 0; i < dataSet.Tables[0].Columns.Count; i++)
            {
                cellValues[0, i] = dataSet.Tables[0].Rows[0].ItemArray[i];
            }

            for (int i = 1; i < dataSet.Tables[0].Rows.Count; i++)
            {
                for (int j = 0; j < dataSet.Tables[0].Columns.Count; j++)
                {
                    object actualValue = dataSet.Tables[0].Rows[i].ItemArray[j];

                    if (double.TryParse(actualValue.ToString(), out double parsedValue))
                    {
                        cellValues[i, j] = parsedValue;
                    }
                    else if (actualValue.ToString().EndsWith(Resources.Event))
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

            return cellValues;
        }

        /// <inheritdoc cref="IFileHandler.CreateImportedFile"/>
        public ImportedFile CreateImportedFile(string excelFilePath, object[,] cellValues, object[,] customCellValues, int newID, System.Windows.Media.Color displayColor, List<ImportedFile> selectedFiles)
        {
            return new ImportedFile
            {
                ID = newID,
                FileName = GenerateUniqueFileName(excelFilePath, selectedFiles),
                FilePath = excelFilePath,
                DisplayColor = displayColor,
                ExcelData = cellValues,
                CustomExcelData = customCellValues,
            };
        }
    }
}
