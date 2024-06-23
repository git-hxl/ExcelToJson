using Newtonsoft.Json;

namespace ExcelToJson
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
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
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }
        }
    }
}
