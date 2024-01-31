namespace Szakdolgozat
{
    public class ImportedFile
    {
        public int ID { get; set; }
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public System.Windows.Media.Color DisplayColor { get; set; }
        public object[,] ExcelData { get; set; }
        public object[,] CustomExcelData { get; set; }
    }
}
