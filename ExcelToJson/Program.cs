using Newtonsoft.Json;

namespace ExcelToJson
{
    internal class Program
    {
        static void Main(string[] args)
        {
            JsonConvertHelper.ConfigureJsonInternal();

            ExcelToolConfig config = new ExcelToolConfig();

            string json = File.ReadAllText("./ExcelToolConfig.json");
            config = JsonConvert.DeserializeObject<ExcelToolConfig>(json);

            ExcelTool excelTool = new ExcelTool(config);

            string[] files = excelTool.ReadAllExcel();

            excelTool.ExportToJsonFile(files);

            excelTool.ExportToCSFile(files);

            Console.ReadKey();
        }
    }
}
