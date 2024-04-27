using ExcelDataReader;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using System.Windows.Media;
using Szakdolgozat.Model;

namespace Szakdolgozat.Interfaces
{
    internal interface IFileHandler
    {
        /// <summary>
        /// Generates an unique file name based on the given excel file path and the list of selected files.
        /// </summary>
        /// <param name="excelFilePath">The path of the excel file.</param>
        /// <param name="selectedFiles">The list of selected files.</param>
        /// <returns>An unique file name.</returns>
        string GenerateUniqueFileName(string excelFilePath, List<ImportedFile> selectedFiles);

        /// <summary>
        /// Reads an excel file and returns its contents as a 2D object array.
        /// </summary>
        /// <param name="excelFilePath">The path of the excel file to read.</param>
        /// <returns>A 2D object array containing the contents of the excel file.</returns>
        Task<object[,]> ReadExcelFile(string excelFilePath);

        /// <summary>
        /// Creates a new ImportedFile object with the specified parameters.
        /// </summary>
        /// <param name="newFileName">The name of the new file.</param>
        /// <param name="excelFilePath">The path of the excel file.</param>
        /// <param name="cellValues">A 2D object array containing the cell values from the excel file.</param>
        /// <param name="customCellValues">A 2D object array containing the custom cell values.</param>
        /// <param name="newID">The ID of the new file.</param>
        /// <param name="displayColor">The color to display for a new file.</param>
        /// <param name="namedValues">Named values.</param>
        /// <returns>A new ImportedFile object.</returns>
        ImportedFile CreateSingleImportedFile(string newFileName,
            string excelFilePath, object[,] cellValues, object[,] customCellValues,
            int newID, Color displayColor, IDictionary<double, string> namedValues);

        /// <summary>
        /// Generates a new unique ID.
        /// The ID is incremented each time this method called.
        /// </summary>
        /// <returns>A new unique ID.</returns>
        int GenerateNewID();

        /// <summary>
        /// Converts the data read from an excel file into a DataSet.
        /// The ExcelDataReader is configured to not use the header row of the excel file.
        /// </summary>
        /// <param name="reader">The ExcelDataReader that has read the excel file.</param>
        /// <returns>A DataSet containing the data from the excel file.</returns>
        DataSet GetDataSetFromReader(IExcelDataReader reader);

        /// <summary>
        /// Validates the structure of a DataSet obtained from an excel file.
        /// </summary>
        /// <param name="dataSet">The DataSet to validate.</param>
        /// <returns>True if the DataSet is valid, false otherwise.</returns>
        bool IsValidDataSet(DataSet dataSet);

        /// <summary>
        /// Extracts cell values from a DataSet and populates them into a 2D object array.
        /// </summary>
        /// <param name="dataSet">The DataSet from which to get the cell values.</param>
        /// <returns>A 2D object array containing the cell values from the DataSet.</returns>
        object[,] GetCellValuesFromDataSet(DataSet dataSet);

        /// <summary>
        /// Creates a new ImportedFile object with the specified parameters.
        /// </summary>
        /// <param name="excelFilePath">The path of the excel file.</param>
        /// <param name="cellValues">A 2D object array containing the cell values from the excel file.</param>
        /// <param name="customCellValues">A 2D object array containing the custom cell values.</param>
        /// <param name="newID">The ID of the new file.</param>
        /// <param name="displayColor">The color to display for a new file.</param>
        /// <param name="namedValues">Named values.</param>
        /// <param name="selectedFiles">The list of selected files.</param>
        /// <returns>A new ImportedFile object.</returns>
        ImportedFile CreateImportedFile(string excelFilePath,
            object[,] cellValues, object[,] customCellValues,
            int newID, Color displayColor, IDictionary<double, string> namedValues, List<ImportedFile> selectedFiles);

        /// <summary>
        ///  Get values for named values from cellValues.
        /// </summary>
        /// <param name="cellValues">A 2D object array containing the cell values from the excel file.</param>
        /// <returns>Returns with the dictionary which contains double values and named values.</returns>
        IDictionary<double, string> GetValuesForNamedValues(object[,] cellValues);
    }
}
