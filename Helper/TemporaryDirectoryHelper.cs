using System;
using System.IO;
using Szakdolgozat.Interfaces;

namespace Szakdolgozat.Helper
{
    internal class TemporaryDirectoryHelper : ITemporaryDirectoryHelper
    {
        /// <inheritdoc cref="ITemporaryDirectoryHelper.CreateTemporaryDirectory"/>
        public string CreateTemporaryDirectory()
        {
            string tempDirectory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempDirectory);
            return tempDirectory;
        }
    }
}
