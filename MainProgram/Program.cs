
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

            ExcelToolConfig excelToolConfig = JsonConvert.DeserializeObject<ExcelToolConfig>(json);


            ExcelTool excelTool = new ExcelTool(excelToolConfig.StartHead, excelToolConfig.OutputJsonDir, excelToolConfig.OutputCSDir);

            string[] files = excelTool.ReadExcelFiles(excelToolConfig.InputExcelDir);

            foreach (var file in files)
            {
                excelTool.ExportToJson(file);
                excelTool.ExportToCS(file);

                Console.WriteLine(file);
            }
            Console.WriteLine("导出成功");

            Console.ReadKey();
        }
    }
}
