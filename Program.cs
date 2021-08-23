using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace evecon
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            Console.WriteLine(ReadExcel());
            WriteExcel();
        }

        private static string ReadExcel()
        {
            var dtTable = new DataTable();
            var rowList = new List<string>();
            ISheet sheet;
            using (var stream = new FileStream("TestData.xlsx", FileMode.Open))
            {
                stream.Position = 0;
                var xssWorkbook = new XSSFWorkbook(stream);
                sheet = xssWorkbook.GetSheetAt(0);
                var headerRow = sheet.GetRow(0);
                int cellCount = headerRow.LastCellNum;
                for (var j = 0; j < cellCount; j++)
                {
                    var cell = headerRow.GetCell(j);
                    if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;
                    {
                        dtTable.Columns.Add(cell.ToString());
                    }
                }

                for (var i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                {
                    var row = sheet.GetRow(i);
                    if (row == null) continue;
                    if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;
                    for (int j = row.FirstCellNum; j < cellCount; j++)
                        if (row.GetCell(j) != null)
                            if (!string.IsNullOrEmpty(row.GetCell(j).ToString()) &&
                                !string.IsNullOrWhiteSpace(row.GetCell(j).ToString()))
                                rowList.Add(row.GetCell(j).ToString());
                    if (rowList.Count > 0)
                        dtTable.Rows.Add(rowList.ToArray());
                    rowList.Clear();
                }
            }

            return JsonConvert.SerializeObject(dtTable);
        }

        private static void WriteExcel()
        {
            var persons = new List<UserDetails>
            {
                new() {ID = "1001", Name = "ABCD", City = "City1", Country = "USA"},
                new() {ID = "1002", Name = "PQRS", City = "City2", Country = "INDIA"},
                new() {ID = "1003", Name = "XYZZ", City = "City3", Country = "CHINA"},
                new() {ID = "1004", Name = "LMNO", City = "City4", Country = "UK"}
            };

            // Lets converts our object data to Datatable for a simplified logic.
            // Datatable is most easy way to deal with complex datatypes for easy reading and formatting.

            var table = (DataTable) JsonConvert.DeserializeObject(JsonConvert.SerializeObject(persons),
                typeof(DataTable));
            var memoryStream = new MemoryStream();

            using (var fs = new FileStream("Result.xlsx", FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                var excelSheet = workbook.CreateSheet("Sheet1");

                var columns = new List<string>();
                var row = excelSheet.CreateRow(0);
                var columnIndex = 0;

                foreach (DataColumn column in table.Columns)
                {
                    columns.Add(column.ColumnName);
                    row.CreateCell(columnIndex).SetCellValue(column.ColumnName);
                    columnIndex++;
                }

                var rowIndex = 1;
                foreach (DataRow dsrow in table.Rows)
                {
                    row = excelSheet.CreateRow(rowIndex);
                    var cellIndex = 0;
                    foreach (var col in columns)
                    {
                        row.CreateCell(cellIndex).SetCellValue(dsrow[col].ToString());
                        cellIndex++;
                    }

                    rowIndex++;
                }

                workbook.Write(fs);
            }
        }
    }
}