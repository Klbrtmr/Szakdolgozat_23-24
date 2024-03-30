namespace Szakdolgozat.Interfaces
{
    internal interface ITemporaryDirectoryHelper
    {
        /// <summary>
        /// Creates a temporary directory in the system's temp path.
        /// </summary>
        /// <returns>The path of the created temporary directory.</returns>
        string CreateTemporaryDirectory();
    }
}
