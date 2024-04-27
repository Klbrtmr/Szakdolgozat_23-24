using System.Collections.Generic;

namespace Szakdolgozat.Model
{
    /// <summary>
    /// Represents an imported file in the application.
    /// It includes properties for the file's ID, name, path, display color, and excel data.
    /// </summary>
    public class ImportedFile
    {
        /// <summary>
        /// Gets or sets the ID of the imported file.
        /// </summary>
        public int ID { get; set; }

        /// <summary>
        /// Gets or sets the name of the imported file.
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// Gets or sets the path of the imported file.
        /// </summary>
        public string FilePath { get; set; }

        /// <summary>
        /// Gets or sets the display color of the imported file.
        /// </summary>
        public System.Windows.Media.Color DisplayColor { get; set; }

        /// <summary>
        /// Gets or sets the excel data of the imported file.
        /// </summary>
        public object[,] ExcelData { get; set; }

        /// <summary>
        /// Gets or sets the custom excel data of the imported file.
        /// </summary>
        public object[,] CustomExcelData { get; set; }

        /// <summary>
        /// Gets or sets the name of values of the imported file.
        /// </summary>
        public IDictionary<double, string> NamedValues { get; set; }
    }
}
