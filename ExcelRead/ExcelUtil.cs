using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelRead
{
    
    public static class ExcelRead
    {
        public static DataTable ToDataTable(System.IO.Stream stream, int columnNumber , bool IsFirstRowHeading)
        {
            int startAtRow = 2;
            if (!IsFirstRowHeading)
            {
                startAtRow = 1;
            }
            ExcelPackage package = new ExcelPackage(stream);
            ExcelWorksheet workSheet = package.Workbook.Worksheets.First();
            DataTable table = new DataTable();
            foreach (var firstRowCell in workSheet.Cells[1, 1, 1, columnNumber])
            {
                table.Columns.Add(firstRowCell.Text);
            }
            for (var rowNumber = startAtRow; rowNumber <= workSheet.Dimension.End.Row; rowNumber++)
            {
                var row = workSheet.Cells[rowNumber, 1, rowNumber, columnNumber];
                var newRow = table.NewRow();
                foreach (var cell in row)
                {
                    newRow[cell.Start.Column - 1] = cell.Text;
                }
                table.Rows.Add(newRow);
            }
            return table;
        }
        /// <summary>
        /// row 1 is assumes to be heading 
        /// </summary>
        /// <param name="package"></param>
        /// <param name="columnNumber"></param>
        /// <returns></returns>
        public static List<FileData> ToDynamicList(System.IO.Stream stream, int columnNumber, bool IsFirstRowHeading)
        {
            int startAtRow = 2;
            if (!IsFirstRowHeading)
            {
                startAtRow = 1;
            }
            try
            {
                ExcelPackage package = new ExcelPackage(stream);
                ExcelWorksheet workSheet = package.Workbook.Worksheets.FirstOrDefault();

                if (workSheet == null)
                {
                    throw new Exception($"no worksheet found in the excel file");
                }

                List<FileData> list = new List<FileData>();

                for (var rowNumber = startAtRow; rowNumber <= workSheet.Dimension.End.Row; rowNumber++)
                {
                    var row = workSheet.Cells[rowNumber, 1, rowNumber, columnNumber];

                    list.Add(ToObject(row, columnNumber));
                }
                return list;
            }
            catch (Exception er)
            {
                throw new Exception($"Error  reading stream: Mesage : {er.Message}");
            }
        }

        private static FileData ToObject(ExcelRange row, int columnNumber)
        {
            FileData rowObject = new FileData();
            if (row != null)
            {
                int columnIndex = 0;
                foreach (ExcelRangeBase cell in row)
                {
                    columnIndex++;
                    if (columnIndex == 1) { rowObject.A = cell?.Text; }
                    else if (columnIndex == 2) { rowObject.B = cell?.Text; }
                    else if (columnIndex == 3) { rowObject.C = cell?.Text; }
                    else if (columnIndex == 4) { rowObject.D = cell?.Text; }
                    else if (columnIndex == 5) { rowObject.E = cell?.Text; }
                    else if (columnIndex == 6) { rowObject.F = cell?.Text; }
                    else if (columnIndex == 7) { rowObject.G = cell?.Text; }
                    else if (columnIndex == 8) { rowObject.H = cell?.Text; }
                    else if (columnIndex == 9) { rowObject.I = cell?.Text; }
                    else if (columnIndex == 10) { rowObject.J = cell?.Text; }
                    else if (columnIndex == 11) { rowObject.K = cell?.Text; }
                    else if (columnIndex == 12) { rowObject.L = cell?.Text; }
                    else if (columnIndex == 13) { rowObject.M = cell?.Text; }
                    else if (columnIndex == 14) { rowObject.N = cell?.Text; }
                    else if (columnIndex == 15) { rowObject.O = cell?.Text; }
                    else if (columnIndex == 16) { rowObject.P = cell?.Text; }
                    else if (columnIndex == 17) { rowObject.Q = cell?.Text; }
                    else if (columnIndex == 18) { rowObject.R = cell?.Text; }
                    else if (columnIndex == 19) { rowObject.S = cell?.Text; }
                    else if (columnIndex == 20) { rowObject.T = cell?.Text; }
                    else if (columnIndex == 21) { rowObject.U = cell?.Text; }
                    else if (columnIndex == 22) { rowObject.V = cell?.Text; }
                    else if (columnIndex == 23) { rowObject.W = cell?.Text; }
                    else if (columnIndex == 24) { rowObject.X = cell?.Text; }
                    else if (columnIndex == 25) { rowObject.Y = cell?.Text; }
                    else if (columnIndex == 26) { rowObject.Z = cell?.Text; }
                }

            }
            return rowObject;
        }

        public class FileData
        {
            public string A { get; set; }
            public string B { get; set; }
            public string C { get; set; }
            public string D { get; set; }
            public string E { get; set; }
            public string F { get; set; }
            public string G { get; set; }
            public string H { get; set; }
            public string I { get; set; }
            public string J { get; set; }
            public string K { get; set; }
            public string L { get; set; }
            public string M { get; set; }
            public string N { get; set; }
            public string O { get; set; }
            public string P { get; set; }
            public string Q { get; set; }
            public string R { get; set; }
            public string S { get; set; }
            public string T { get; set; }
            public string U { get; set; }
            public string V { get; set; }
            public string W { get; set; }
            public string X { get; set; }
            public string Y { get; set; }
            public string Z { get; set; }
        }
    }

    
}

