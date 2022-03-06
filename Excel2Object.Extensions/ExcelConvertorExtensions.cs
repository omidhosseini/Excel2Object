using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Excel2Object.Extensions
{
    public static class ExcelConvertorExtensions
    {
        public static IList<TOut> ToList<TOut>(this Stream file) where TOut : new()
        {
            var outputProperties = typeof(TOut).GetProperties().ToList();
            IList<TOut> output = new List<TOut>();
            ConcurrentDictionary<string, int> columnIndex = new ConcurrentDictionary<string, int>();

            XSSFWorkbook xssWorkbook = new XSSFWorkbook(file);
            ISheet sheet = xssWorkbook.GetSheetAt(0);

            IRow columnTitles = sheet.GetRow(0);

            foreach (var title in columnTitles)
            {
                columnIndex.TryAdd(title.StringCellValue, title.ColumnIndex);
            }

            for (int row = 1; row <= sheet.LastRowNum; row++)
            {
                TOut newObj = new TOut();
                foreach (var item in outputProperties)
                {
                    if (!columnIndex.TryGetValue(item.Name, out var cellIndex)) continue;

                    if (sheet.GetRow(row) is null) continue;

                    if (sheet.GetRow(row).GetCell(cellIndex) is null) continue;

                    var objProp = newObj.GetType().GetProperty(item.Name);
                    if(!objProp.CanWrite) continue;

                    var cellValue = Convert.ChangeType(sheet.GetRow(row).Cells[cellIndex].ToString(), item.PropertyType);
                    objProp.SetValue(newObj, cellValue, null);
                }

                output.Add(newObj);
            }

            return output;
        }

        public static ObjectToExcelFileResult ToExcelFile<TIn>(this IList<TIn> dataList, string fileName = default) where TIn : new()
        {
            var inputTypeProperty = typeof(TIn).GetProperties();
            var dataProps = dataList.FirstOrDefault().GetType().GetProperties();

            int rowIndex = 0;
            int cellIndex = 0;

            XSSFWorkbook xssWorkbook = new XSSFWorkbook();
            ISheet sheet = xssWorkbook.CreateSheet();

            IRow row = sheet.CreateRow(rowIndex++);

            foreach (var item in inputTypeProperty)
            {
                ICell cell = row.CreateCell(cellIndex);
                cell.SetCellValue(item.Name);
                cellIndex++;
            }

            foreach (var data in dataList)
            {
                cellIndex = 0;
                row = sheet.CreateRow(rowIndex);
                foreach (var item in dataProps)
                {
                    ICell cell = row.CreateCell(cellIndex);
                    var celValue = item.GetValue(data)?.ToString() ?? "-";
                    cell.SetCellValue(celValue);
                    cellIndex++;
                }
                rowIndex++;
            }



            MemoryStream ms = new MemoryStream();
            xssWorkbook.Write(ms);
            ms.Close();
            xssWorkbook.Close();

            var result = new ObjectToExcelFileResult
            {
                ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                FileName = $"{fileName ?? DateTime.UtcNow.ToString("MM-dd-yyyy")}.xlsx",                
                File = ms.ToArray()
            };

            return result;
        }
    }
}
