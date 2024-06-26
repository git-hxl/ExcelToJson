
using Newtonsoft.Json;
using System.Data;
using System.Text;

namespace ExcelToJson
{
    public class ExcelTool
    {
        private ExcelToolConfig _config;

        public ExcelTool(ExcelToolConfig config)
        {
            _config = config;

            if (!Directory.Exists(config.OutputCSDir))
            {
                Directory.CreateDirectory(config.OutputCSDir);
            }

            if (!Directory.Exists(config.OutputJsonDir))
            {
                Directory.CreateDirectory(config.OutputJsonDir);
            }

        }

        /// <summary>
        /// 读取所有的Excel
        /// </summary>
        public string[] ReadAllExcel()
        {
            string[] fileExtensions = new string[] { ".xls", ".xlsx" };

            string[] excelFiles = null;

            if (!string.IsNullOrEmpty(_config.InputExcelDir) && Directory.Exists(_config.InputExcelDir))
            {
                excelFiles = Directory.GetFiles(_config.InputExcelDir).Where(file => fileExtensions.Contains(Path.GetExtension(file)) && !file.Contains("~$"))
               .ToArray();
            }

            return excelFiles;
        }

        /// <summary>
        /// 导出成Json
        /// </summary>
        public void ExportToJsonFile(string[] excelFiles)
        {
            foreach (var file in excelFiles)
            {
                var tables = Utility.Excel.ReadExcelAllSheets(file);

                foreach (DataTable table in tables)
                {
                    if (table.Rows.Count > 0)
                    {
                        var newTable = Utility.Excel.SelectContent(table, _config.StartHead);

                        string json = JsonConvert.SerializeObject(newTable, Formatting.Indented);
                        if (!string.IsNullOrEmpty(json))
                        {
                            string fileName = table.TableName + ".json";
                            using (FileStream stream = new FileStream(_config.OutputJsonDir + "/" + fileName, FileMode.Create, FileAccess.ReadWrite))
                            {
                                byte[] data = Encoding.UTF8.GetBytes(json);
                                stream.Write(data, 0, data.Length);
                            }

                            Console.WriteLine($"导出文件：{file} {table.TableName}");
                        }
                    }
                }
            }
            Console.WriteLine("Json导出成功");
        }

        /// <summary>
        /// 导出成CS
        /// </summary>
        public void ExportToCSFile(string[] excelFiles)
        {
            foreach (var file in excelFiles)
            {
                var tables = Utility.Excel.ReadExcelAllSheets(file);

                foreach (DataTable table in tables)
                {
                    if (table.Rows.Count > 0)
                    {
                        ExcelToCS.Generate(table, _config.OutputCSDir);
                    }
                }
            }
            Console.WriteLine("CS导出成功");
        }
    }
}