using Newtonsoft.Json;
using OfficeOpenXml;
using System.Data;
using System.Drawing;
using System.Numerics;

namespace ExcelToJson
{
    public static partial class Utility
    {
        public static class Excel
        {
            /// <summary>
            /// 创建表
            /// </summary>
            /// <param name="filePath"></param>
            /// <param name="firstSheetName"></param>
            /// <param name="dataTable"></param>
            public static void CreateExcel(string filePath, string firstSheetName, DataTable dataTable = null)
            {
                FileInfo fileInfo = new FileInfo(filePath);
                using (var package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(firstSheetName);

                    if (dataTable != null)
                    {
                        for (int i = 0; i < dataTable.Rows.Count; i++)
                        {
                            for (int j = 0; j < dataTable.Columns.Count; j++)
                            {
                                var value = dataTable.Rows[i][j];
                                worksheet.Cells[i + 1, j + 1].Value = value;
                            }
                        }
                    }
                    package.Save();
                }
            }

            /// <summary>
            /// 写入
            /// </summary>
            /// <param name="filePath"></param>
            /// <param name="sheetIndex">默认初始index= 1</param>
            /// <param name="dataTable"></param>
            public static void WriteToExcel(string filePath, int sheetIndex, DataTable dataTable)
            {
                FileInfo fileInfo = new FileInfo(filePath);
                using (var package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetIndex];
                    int lastRowIndex = worksheet.Dimension.End.Row;
                    if (dataTable != null)
                    {
                        for (int i = 0; i < dataTable.Rows.Count; i++)
                        {
                            for (int j = 0; j < dataTable.Columns.Count; j++)
                            {
                                var value = dataTable.Rows[i][j];
                                worksheet.Cells[lastRowIndex + i + 1, j + 1].Value = value;
                            }
                        }
                    }
                    package.Save();
                }
            }

            /// <summary>
            /// 读取表
            /// </summary>
            /// <param name="filePath"></param>
            /// <returns></returns>
            public static List<DataTable> ReadExcelAllSheets(string filePath)
            {
                List<DataTable> tables = new List<DataTable>();

                FileInfo fileInfo = new FileInfo(filePath);
                using (var package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheets worksheets = package.Workbook.Worksheets;
                    foreach (var item in worksheets)
                    {
                        if (item.Dimension != null)
                        {
                            tables.Add(ConvertToDataTable(item));
                        }
                    }
                }
                return tables;
            }

            /// <summary>
            /// 读取表
            /// </summary>
            /// <param name="filePath"></param>
            /// <returns></returns>
            public static DataTable ReadExcelSheet(string filePath)
            {
                List<DataTable> tables = new List<DataTable>();

                FileInfo fileInfo = new FileInfo(filePath);
                using (var package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheets worksheets = package.Workbook.Worksheets;
                    return ConvertToDataTable(worksheets[1]);
                }
            }


            /// <summary>
            /// 转换数据格式
            /// </summary>
            /// <param name="worksheet"></param>
            /// <returns></returns>
            private static DataTable ConvertToDataTable(ExcelWorksheet worksheet)
            {
                int rows = worksheet.Dimension.End.Row;
                int cols = worksheet.Dimension.End.Column;

                DataTable dataTable = new DataTable(worksheet.Name);

                for (int i = 1; i <= rows; i++)
                {
                    DataRow row = dataTable.Rows.Add();

                    for (int j = 1; j <= cols; j++)
                    {
                        if (i == 1)
                        {
                            if (worksheet.Cells[i, j].Value != null)
                                dataTable.Columns.Add(worksheet.Cells[i, j].Value.ToString());
                            else
                                dataTable.Columns.Add("");
                        }
                        row[j - 1] = worksheet.Cells[i, j].Value;
                    }
                }
                return dataTable;
            }



            /// <summary>
            /// 删除行
            /// </summary>
            /// <param name="filePath"></param>
            /// <param name="sheetIndex">默认初始index= 1</param>
            /// <param name="rowIndex">大于等于１</param>
            public static void DeleteExcelRow(string filePath, int sheetIndex, int rowIndex)
            {
                FileInfo fileInfo = new FileInfo(filePath);
                using (var package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetIndex];
                    worksheet.DeleteRow(rowIndex);
                    package.Save();
                }
            }

            /// <summary>
            /// 获取表指定数据内容
            /// </summary>
            /// <param name="dataTable"></param>
            /// <param name="rowIndex"> 指定行 默认第3行</param>
            /// <param name="typeIndex"> 类型行 小于指定行 默认第1行</param>
            /// <param name="nameIndex"> 名称行 小于指定行 默认第0行 </param>
            /// <param name="isConvertType"> 是否根据第二行类型typeIndex进行转换 </param>
            /// <returns></returns>
            public static DataTable SelectContent(DataTable dataTable, int rowIndex = 3, int typeIndex = 1, int nameIndex = 0, bool isConvertType = true)
            {
                DataTable newDataTable = new DataTable();

                for (int i = rowIndex; i < dataTable.Rows.Count; i++)
                {
                    DataRow dataRow = newDataTable.Rows.Add();
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        if (dataTable.Rows[typeIndex][j] == null || string.IsNullOrEmpty(dataTable.Rows[typeIndex][j].ToString()))
                        {
                            break;
                        }

                        if (i == rowIndex)
                        {
                            string typeStr = dataTable.Rows[typeIndex][j].ToString();
                            string columnName = dataTable.Rows[nameIndex][j].ToString();

                            Type type = typeof(string);
                            if (isConvertType)
                            {
                                type = GetTypeByString(typeStr);
                            }

                            DataColumn dataColumn = new DataColumn(columnName, type);
                            newDataTable.Columns.Add(dataColumn);
                        }
                        dataRow[j] = ConvertDataTableData(dataTable.Rows[i][j].ToString(), newDataTable.Columns[j].DataType, dataTable.TableName);
                    }
                }
                return newDataTable;
            }

            /// <summary>
            /// 格式化数据
            /// </summary>
            /// <param name="data"></param>
            /// <param name="type"></param>
            /// <returns></returns>
            public static object ConvertDataTableData(string data, Type type, string tableName)
            {
                try
                {
                    if (string.IsNullOrEmpty(data))
                    {
                        return type.IsValueType ? Activator.CreateInstance(type) : null;
                    }
                    else if (data.StartsWith('[') && data.EndsWith(']'))
                    {
                        return JsonConvert.DeserializeObject(data, type);
                    }
                    else
                    {
                        return data;
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine($"表 {tableName} 数据格式化出错：{data} 目标类型 {type}");
                    Console.WriteLine(e.ToString());

                    return null;

                }
            }

            public static bool CanParse(Type type, string str)
            {
                var tryParse = type.GetMethod("TryParse", new Type[] { typeof(string) });
                if (tryParse == null)
                    return false;

                return true;
            }

            /// <summary>
            /// 根据名称返回类型
            /// </summary>
            /// <param name="typeName"></param>
            /// <returns></returns>
            public static Type GetTypeByString(string typeName)
            {
                switch (typeName.ToLower())
                {
                    case "uint":
                        return typeof(uint);
                    case "int":
                        return typeof(int);
                    case "long":
                        return typeof(long);
                    case "float":
                        return typeof(float);
                    case "double":
                        return typeof(double);
                    case "bool":
                        return typeof(bool);
                    case "string":
                        return typeof(string);
                    case "vector2":
                        return typeof(Vector2);
                    case "vector3":
                        return typeof(Vector3);
                    case "vector4":
                        return typeof(Vector4);
                    case "color":
                        return typeof(Color);

                    case "int[]":
                        return typeof(int[]);
                    case "long[]":
                        return typeof(long[]);
                    case "float[]":
                        return typeof(float[]);
                    case "double[]":
                        return typeof(double[]);
                    case "bool[]":
                        return typeof(bool[]);
                    case "string[]":
                        return typeof(string[]);
                    case "vector2[]":
                        return typeof(Vector2[]);
                    case "vector3[]":
                        return typeof(Vector3[]);
                    case "vector4[]":
                        return typeof(Vector4[]);
                    case "color[]":
                        return typeof(Color[]);

                    case "date":
                    case "datetime":
                        return typeof(DateTime);
                    default:
                        return typeof(string);
                }
            }
        }

    }

}
