
namespace ExcelToJson
{
    public class ExcelToolConfig
    {
        public int StartHead { get; set; }

        public string InputExcelDir { get; set; }
        public string OutputJsonDir { get; set; }
        public string OutputCSDir { get; set; }

        public bool AutoParse { get; set; }
    }
}
