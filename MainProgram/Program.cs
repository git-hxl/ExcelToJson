
using Newtonsoft.Json;
using System;
using System.IO;

namespace Excel2Json
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //try
            //{
            JsonConvertHelper.ConfigureJsonInternal();

            string json = File.ReadAllText("./ExcelToolConfig.json");

            ExcelToolConfig config = JsonConvert.DeserializeObject<ExcelToolConfig>(json);

            string[] files = ExportTool.ReadExcelFiles(config.InputExcelDir);

            foreach (var file in files)
            {
                ExportTool.ExportToJson(file,config.OutputJsonDir,config.StartHead);
                ExportTool.ExportToCs(file,config.OutputCSDir);

                Console.WriteLine(file);
            }
            Console.WriteLine("导出成功");

            Console.ReadKey();
        }
    }
}
