using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace ReadLargeXLSX
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = "C:/Users/Admin/Downloads/test.xlsx";

            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);

            //1. Reading from a binary Excel file ('97-2003 format; *.xls)
            //IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            //...

            //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            //...
            //3. DataSet - The result of each spreadsheet will be created in the result.Tables
            //DataSet result = excelReader.AsDataSet();
            //...
            //4. DataSet - Create column names from first row
            //excelReader.IsFirstRowAsColumnNames = true;
            //excelReader.
            //DataSet result = excelReader.AsDataSet();

            //5. Data Reader methods
            while (excelReader.Read())
            {
                int colCnt = excelReader.FieldCount;
                Console.WriteLine(excelReader.GetString(colCnt-1));
                break;
                //excelReader.GetInt32(0);
            }

            //6. Free resources (IExcelDataReader is IDisposable)
            excelReader.Close();
        }
    }
}
