namespace ExcelRead.Models
{
    public class ExcelData
    {
        public List<string> Headers { get; set; } = new List<string>();
        public List<Dictionary<string, string>> Rows { get; set; } = new List<Dictionary<string, string>>();
    }
}
