namespace Szakdolgozat.Interfaces
{
    internal interface IImportControl
    {
        /// <summary>
        /// Gets or sets the number of files imported from the current EDF file.
        /// </summary>
        int ImportFileNumber { get; set; }

        /// <summary>
        /// Imports a project from an EDF file.
        /// </summary>
        /// <param name="edfFilePath">The path of the EDF file to import.</param>
        void ImportProject(string edfFilePath);

        /// <summary>
        /// Creates a temporary directory in the system's temp path.
        /// </summary>
        /// <returns>The path of the created temporary directory.</returns>
        string CreateTemporaryDirectory();
    }
}
