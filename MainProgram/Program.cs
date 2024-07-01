
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

            ExcelTool excelTool = new ExcelTool(json);

            excelTool.ExportToJsonFile();

            excelTool.ExportToCSFile();

            Console.WriteLine("导出成功");

            Console.ReadKey();


            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine(ex.Message);
            //    Console.ReadKey();
            //}
        }
    }
}
