﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;




namespace _EXCEL_ExcelAllLanguagesCommonDenominator
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("State filepath to first bilingual Excel file:");
                //string firstExcelFile = @"C:\Users\Bernd\Downloads\Csharp\_EXCEL_ExcelAllLanguagesCommonDenominator\testfiles\en-DE.xlsx";   //debug
                //string firstExcelFile = @"C:\Users\oelll\Dropbox\_ME\III Professionella Expertis\C# Project\_EXCEL_ExcelAllLanguagesCommonDenominator\testfiles\en-DE.xlsx";   //debug
                string initialExcelFile = Console.ReadLine();
                Console.WriteLine();

                Console.WriteLine(@"State filepaths to other bilingual Excel files. Format: Filepaths, separated by semicolons, no linebreaks. Example: 'C:\otherExcelFiles\en-FR;C:\otherExcelFiles\en-NO;C:\otherExcelFiles\en-SV;C:\otherExcelFiles\en-FI'");
                 string toBeMergedExcelFiles = Console.ReadLine();
                 List<string> toBeMergedExcelFilesList = new List<string>(toBeMergedExcelFiles.Split(';'));   //notice the subtle difference? Split(";") doesnt work ...
                int amountFilesToBeMerged = toBeMergedExcelFilesList.Count();
                Console.WriteLine();

                string mergedFile = Path.GetDirectoryName(initialExcelFile) + @"\" + "merged_" + DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") + "." + Path.GetExtension(initialExcelFile);

                Dictionary<string, string> initialExcelData = ReadExcelData(initialExcelFile);
                Dictionary<string, string> toBeMergedExcelFilesData = ReadExcelData(toBeMergedExcelFiles);

                List<Tuple<string, string, string>> matchedData = GetMatchingRows(initialExcelData, toBeMergedExcelFilesData);

                CreateMergedExcel(mergedFile, matchedData);

                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Successfully created merged Excel (see directory of initial Excel file)");
                Console.ResetColor();
            }
            catch (Exception exception)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(">>> ERROR <<<: " + exception);
                Console.ResetColor();
            }
        }




        private static Dictionary<string, string> ReadExcelData(string filePath)
        {
            Console.WriteLine("Reading Excel files ...");

            Excel.Application excelApplication = null;
            Excel.Workbook excelWorkbook = null;

            try
            {
                excelApplication = new Excel.Application();
                excelApplication.Visible = false;
                excelWorkbook = excelApplication.Workbooks.Open(filePath);
                Excel.Worksheet excelWorksheet = excelWorkbook.Worksheets[1];   //assuming that FIRST worksheet is the one

                int lastRow = excelWorksheet.UsedRange.Rows.Count;
                int lastColumn = excelWorksheet.UsedRange.Columns.Count;

                Dictionary<string, string> rowData = new Dictionary<string, string>();
                for (int row = 2; row <= lastRow; row++)   //"row = 2" => assuming that first row is headings "English", "German", etc.
                {
                    //Only iterate over columns A and B (1 and 2 in Excel Interop):
                    string key = excelWorksheet.Cells[row, 1].Value2?.ToString();   //captures column A row 2 cell value, then r3, and so on ...
                    string value = excelWorksheet.Cells[row, 2].Value2?.ToString();   //captures column B row 2 cell value, then r3, and so on ...
                    Console.WriteLine("col A | B: " + key + " | " + value);   //debug

                    if (key != null && value != null)   //Add only if both key and value are not null
                    {
                        rowData.Add(key, value);
                    }
                }

                Console.WriteLine();

                excelWorkbook.Close();
                excelApplication.Quit();

                return rowData;
            }
            catch (Exception exception)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(">>> ERROR <<<: " + exception);
                Console.ResetColor();

                excelWorkbook.Close();
                excelApplication.Quit();

                return null;
            }
        }




        private static List<Tuple<string, string, string>> GetMatchingRows(Dictionary<string, string> initialExcelData, Dictionary<string, string> toBeMergedExcelFileslData)
        {
            Console.WriteLine("Matching rows ...");

            try
            {
                List<Tuple<string, string, string>> matchedData = new List<Tuple<string, string, string>>();   //this is how you store multiple values in a list, here as "3-tuples" (more than the merely 2's in a dictionary)

                foreach (KeyValuePair<string, string> sourceTextAndtranslationRow_ in initialExcelData)
                {
                    foreach (KeyValuePair<string, string> sourceTextAndtranslationRow__ in toBeMergedExcelFileslData)
                    {
                        if (sourceTextAndtranslationRow_.Key == sourceTextAndtranslationRow__.Key)
                        {
                            matchedData.Add(new Tuple<string, string, string>(sourceTextAndtranslationRow_.Key, sourceTextAndtranslationRow_.Value, sourceTextAndtranslationRow__.Value));   //this is how to add Tuples to that Tuple list; source text + translation first (e.g. German) + translation second (e.g. Swedish)!
                        }
                    }
                }

                foreach (Tuple<string, string, string> item in matchedData)
                {
                    Console.WriteLine("matchedDate:" + item.Item1 + " " + item.Item2 + " " + item.Item3);
                }
                Console.WriteLine();

                return matchedData;

            }
            catch (Exception exception)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(">>> ERROR <<<: " + exception);
                Console.ResetColor();

                return null;
            }
        }




        private static void CreateMergedExcel(string filePath, List<Tuple<string, string, string>> matchedData)
        {
            Console.WriteLine("Creating merged Excel ...");

            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;

            try
            {
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                workbook = excelApp.Workbooks.Add();
                Excel.Worksheet worksheet = workbook.Worksheets[1];

                //Add Headers:
                worksheet.Cells[1, 1] = "Source language";
                worksheet.Cells[1, 2] = "Target language 1";
                worksheet.Cells[1, 3] = "Target language 2";

                int row = 2;
                foreach (Tuple<string, string, string> item in matchedData)
                {
                    worksheet.Cells[row, 1] = item.Item1;
                    worksheet.Cells[row, 2] = item.Item2;
                    worksheet.Cells[row, 3] = item.Item3;
                    row++;
                }

                workbook.SaveAs(filePath);
                workbook.Close();
                excelApp.Quit();
            }
            catch (Exception exception)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(">>> ERROR <<<: " + exception);
                Console.ResetColor();
                workbook.Close();
                excelApp.Quit();
            }
        }
    }
}