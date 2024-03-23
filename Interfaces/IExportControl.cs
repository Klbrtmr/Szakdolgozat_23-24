namespace Szakdolgozat.Interfaces
{
    internal interface IExportControl
    {
        /// <summary>
        /// Exports a project to an EDF file.
        /// </summary>
        /// <param name="edfFilePath">The path of the EDF file to create.</param>
        void ExportProject(string edfFilePath);
    }
}
