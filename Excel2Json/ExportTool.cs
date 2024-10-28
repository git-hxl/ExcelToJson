
using Newtonsoft.Json;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace Excel2Json
{
    public class ExportTool
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="outPath"></param>
        /// <param name="startHead">startHead是主体内容 默认0是名称 1是类型 </param>
        public static void ExportToJson(string filePath, string outDir, int startHead = 5)
        {
            var tables = ExcelHelper.ReadExcelAllSheets(filePath);


            if (!Directory.Exists(outDir))
            {
                Directory.CreateDirectory(outDir);
            }

            foreach (DataTable table in tables)
            {
                if (table.Rows.Count > 0)
                {
                    var newTable = ExcelHelper.SelectContent(table, startHead);

                    string json = JsonConvert.SerializeObject(newTable, Formatting.Indented);
                    if (!string.IsNullOrEmpty(json))
                    {
                        string fileName = table.TableName + ".json";
                        using (FileStream stream = new FileStream(outDir + "/" + fileName, FileMode.Create, FileAccess.ReadWrite))
                        {
                            byte[] data = Encoding.UTF8.GetBytes(json);
                            stream.Write(data, 0, data.Length);
                        }
                    }
                }
            }
        }

        public static void ExportToCs(string filePath, string outDir)
        {
            var tables = ExcelHelper.ReadExcelAllSheets(filePath);

            if (!Directory.Exists(outDir))
            {
                Directory.CreateDirectory(outDir);
            }

            foreach (DataTable table in tables)
            {
                if (table.Rows.Count > 0)
                {
                    string fileName = table.TableName + ".cs";

                    System.IO.StreamWriter writer = new System.IO.StreamWriter(outDir + "/" + fileName);
                    writer.WriteLine("// This code was automatically generated");
                    writer.WriteLine();

                    writer.WriteLine($"public class {table.TableName}:IConfig");

                    writer.WriteLine("{");

                    for (int j = 0; j < table.Columns.Count; j++)
                    {
                        string columnName = table.Rows[0][j].ToString();

                        if (string.IsNullOrEmpty(columnName))
                            continue;

                        string typeStr = table.Rows[1][j].ToString();

                        Type type = CusTomType.GetTypeByString(typeStr);

                        writer.WriteLine($"     public {type} {columnName} {{ get; set;}}");
                    }

                    writer.WriteLine("}");
                    writer.Close();
                }
            }
        }


        /// <summary>
        /// 读取目录下所有Excel
        /// </summary>
        /// <param name="directory"></param>
        /// <returns></returns>
        public static string[] ReadExcelFiles(string directory)
        {
            if (Directory.Exists(directory))
            {
                string[] fileExtensions = new string[] { ".xls", ".xlsx" };

                string[] excelFiles = Directory.GetFiles(directory).Where(file => fileExtensions.Contains(Path.GetExtension(file)) && !file.Contains("~$"))
               .ToArray();

                return excelFiles;
            }

            return null;
        }
    }
}