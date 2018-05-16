using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using CsvHelper;

namespace ExcelToCsv
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName = "[input file name here]";
            FileInfo excelPath = new FileInfo($@"C:\ExcelToCsv\{fileName}.xlsx");

            if (excelPath.Directory != null)
            {
                var dt = GetDataTableFromExcel(excelPath.FullName);

                if (dt.Rows.Count > 0)
                {
                    var list = dt.ToList();
                    using (var fs = new FileStream($@"C:\ExcelToCsv\{fileName}.csv", FileMode.Create))
                    {
                        using (var tx = new StreamWriter(fs, Encoding.UTF8))
                        {
                            using (var csvWriter = new CsvWriter(tx))
                            {
                                csvWriter.WriteRecords(list);
                            }
                        }
                    }
                }
            }
        }

        public static DataTable GetDataTableFromExcel(string path, bool hasHeader = true)
        {
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                using (var stream = File.OpenRead(path))
                {
                    pck.Load(stream);
                }
                var ws = pck.Workbook.Worksheets.First();
                DataTable tbl = new DataTable();
                foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                {
                    tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                }
                var startRow = hasHeader ? 2 : 1;
                for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                {
                    var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                    DataRow row = tbl.Rows.Add();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                }
                return tbl;
            }
        }


    }

    public class APOGrupprum
    {
        public string Person { get; set; }
        public string EPost { get; set; }
        public string Personnummer { get; set; }
        public string Uppdragstyp { get; set; }
        public string Arbetsstalle { get; set; }
        public string Arbetsstallenummer { get; set; }
    }

    public static class Extensions
    {
        public static List<APOGrupprum> ToList(this DataTable dt)
        {
            List<APOGrupprum> returnList = new List<APOGrupprum>();

            foreach (DataRow dataRow in dt.Rows)
            {
                Console.WriteLine($"{dt.Rows.IndexOf(dataRow) + 1} / {dt.Rows.Count}");
                APOGrupprum item = new APOGrupprum()
                {
                    Person = dataRow.ItemArray[0].ToString(),
                    EPost = dataRow.ItemArray[1].ToString(),
                    Personnummer = dataRow.ItemArray[2].ToString(),
                    Uppdragstyp = dataRow.ItemArray[3].ToString(),
                    Arbetsstalle = dataRow.ItemArray[4].ToString(),
                    Arbetsstallenummer = dataRow.ItemArray[5].ToString()
                };

                returnList.Add(item);
            }

            return returnList;
        }
    }
}
