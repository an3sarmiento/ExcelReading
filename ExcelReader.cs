using LumenWorks.Framework.IO.Csv;
using NPOI.HSSF.UserModel;
using NPOI.OpenXml4Net.OPC;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using System;
using System.IO;

namespace ExcelReader
{
    public class ExcelReader
    {
        private ISheet FileSheet;
        private CsvReader Csv;
        private int CurrentRow;
        private bool Trim;

        public static ExcelReader GetExcelReader(string Base64String, string FileName, string WSName, string Delimiter)
        {
            byte[] ByteArray = Convert.FromBase64String(Base64String);
            if (FileName.Contains(".xlsx"))
            {
                return new ExcelReader(new XSSFWorkbook(new ByteArrayInputStream(ByteArray)).GetSheet(WSName), null);
            }
            else if (FileName.Contains(".xls"))
            {
                return new ExcelReader(new HSSFWorkbook(new ByteArrayInputStream(ByteArray)).GetSheet(WSName), null);
            }
            else if (FileName.Contains(".csv"))
            {
                return new ExcelReader(null,new CsvReader(new StreamReader(new MemoryStream(ByteArray)), false, (Delimiter.Length >= 1)?Delimiter[0]:';'));
            }
            return null;
        }

        public void setTrim(bool Trim)
        {
            this.Trim = Trim;
        }

        private ExcelReader(ISheet FileSheet, CsvReader Csv)
        {
            this.FileSheet = FileSheet;
            this.Csv = Csv;
            Trim = false;
        }

        public static ISheet GetSheetFromBizagiFile(string Base64String, string FileName, string WSName)
        {
            byte[] ByteArray = Convert.FromBase64String(Base64String);
            ISheet MySheet = null;
            if (FileName.Contains(".xlsx"))
            {
                MySheet = new XSSFWorkbook(new ByteArrayInputStream(ByteArray)).GetSheet(WSName);
            }
            else
            {
                MySheet = new HSSFWorkbook(new ByteArrayInputStream(ByteArray)).GetSheet(WSName);
            }
            return MySheet;
        }

        public static void Main(string[] arg0)
        {
            ExcelReader ER = null;
            string FileLocation = @"C:\Temp\files\Claims Template.xlsx";
            //string SheetName = "PADTable";
            string SheetName = null;
            if (FileLocation.Contains(".xlsx"))
            {
                OPCPackage pkg = OPCPackage.OpenOrCreate(FileLocation);
                XSSFWorkbook hssfwb = new XSSFWorkbook(pkg);
                ISheet MySheet = null;
                if (SheetName == null || SheetName == "")
                    MySheet = hssfwb.GetSheetAt(0);
                else
                    MySheet = hssfwb.GetSheet(SheetName);
                ER = new ExcelReader(MySheet, null);
            }
            else if (FileLocation.Contains(".xls"))
            {
                XSSFWorkbook xssfwb = new XSSFWorkbook(new FileStream(FileLocation, FileMode.Open, FileAccess.Read));
                ISheet MySheet = null;
                if (SheetName == null || SheetName == "")
                    MySheet = xssfwb.GetSheetAt(0);
                else
                    MySheet = xssfwb.GetSheet(SheetName);
                ER = new ExcelReader(MySheet, null);
            }
            else if (FileLocation.Contains(".csv"))
            {
                ER = new ExcelReader(null, new CsvReader(new StreamReader(FileLocation), false, ','));
            }

            var row = 0;
            while (ER.HasNextRow()) //null is when the row only contains empty cells 
            {
                row++;
                Console.WriteLine("Row " + row + " = " + ER.GetRowField(0));
                Console.WriteLine("Row " + row + " = " + ER.GetRowField(1));
                Console.WriteLine("Row " + row + " = " + ER.GetRowField(2));
                Console.WriteLine("Row " + row + " = " + ER.GetRowField(3));
            }
            ER.Close();
        }

        public string Close()
        {
            try
            {
                if (FileSheet != null)
                {
                    FileSheet.Workbook.Close();
                }
                else
                {
                    Csv.Dispose();
                }
                return "Success";
            }
            catch(Exception e)
            {
                return "Problem closing file: " + e.Message;
            }
        }

        public bool HasNextRow()
        {
            if (Csv == null)
            {
                CurrentRow++;
                return FileSheet.GetRow(CurrentRow) != null;
            }
            else
                return Csv.ReadNextRecord();
        }

        public static string GetCellValue(ICell cell, IFormulaEvaluator eval = null)
        {
            if (cell != null)
            {
                switch (cell.CellType)
                {
                    case CellType.String:
                        return cell.StringCellValue;

                    case CellType.Numeric:
                        if (DateUtil.IsCellDateFormatted(cell))
                        {
                            DateTime date = cell.DateCellValue;
                            ICellStyle style = cell.CellStyle;
                            // Excel uses lowercase m for month whereas .Net uses uppercase
                            string format = style.GetDataFormatString().Replace('m', 'M');
                            return date.ToString(format);
                        }
                        else
                        {
                            return cell.NumericCellValue.ToString();
                        }

                    case CellType.Boolean:
                        return cell.BooleanCellValue ? "TRUE" : "FALSE";

                    case CellType.Formula:
                        if (eval != null)
                            return GetCellValue(eval.EvaluateInCell(cell));
                        else
                            return cell.CellFormula;

                    case CellType.Error:
                        return FormulaError.ForInt(cell.ErrorCellValue).String;
                }
            }
            // null or blank cell, or unknown cell type
            return string.Empty;
        }


        public string GetRowField(int column, IFormulaEvaluator eval = null)
        {
            string Value = GetTrimedRow(Trim, column, eval);
            if (Trim)
                Value = Value.Trim();
            return Value;
        }

        private string GetTrimedRow(bool trimed, int column, IFormulaEvaluator eval = null)
        {
            ICell cell = (Csv == null) ? FileSheet.GetRow(CurrentRow).GetCell(column) : null;
            if (Csv == null && cell != null)
            {
                switch (cell.CellType)
                {
                    case CellType.String:
                        return cell.StringCellValue;

                    case CellType.Numeric:
                        if (DateUtil.IsCellDateFormatted(cell))
                        {
                            DateTime date = cell.DateCellValue;
                            ICellStyle style = cell.CellStyle;
                            // Excel uses lowercase m for month whereas .Net uses uppercase
                            string format = style.GetDataFormatString().Replace('m', 'M');
                            return date.ToString(format);
                        }
                        else
                        {
                            return cell.NumericCellValue.ToString();
                        }

                    case CellType.Boolean:
                        return cell.BooleanCellValue ? "TRUE" : "FALSE";

                    case CellType.Formula:
                        if (eval != null)
                            return GetCellValue(eval.EvaluateInCell(cell));
                        else
                            return cell.CellFormula;

                    case CellType.Error:
                        return FormulaError.ForInt(cell.ErrorCellValue).String;
                }
            }
            else if (Csv != null)
            {
                return Csv[column];
            }
            // null or blank cell, or unknown cell type
            return string.Empty;
        }
    }
}
