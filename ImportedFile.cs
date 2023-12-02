using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Szakdolgozat
{
    public class ImportedFile
    {
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public System.Windows.Media.Color DisplayColor { get; set; }
        public object[,] ExcelData { get; set; }
    }
}
