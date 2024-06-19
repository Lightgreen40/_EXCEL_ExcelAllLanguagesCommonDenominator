using System;
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
            //Replace these paths with your actual file paths:
            Console.WriteLine("State filepath to first bilingual Excel file:");
            string firstExcelFile = @"C:\Users\oelll\Dropbox\_ME\III Professionella Expertis\C# Project\_EXCEL_ExcelAllLanguagesCommonDenominator\testfiles\en-DE.xlsx";   //debug
            //string firstExcelFile = Console.ReadLine();
            Console.WriteLine();

            Console.WriteLine("State filepath to second bilingual Excel file:");
            string secondExcelFile = @"C:\Users\oelll\Dropbox\_ME\III Professionella Expertis\C# Project\_EXCEL_ExcelAllLanguagesCommonDenominator\testfiles\en-SV.xlsx";   //debug
            //string secondExcelFile = Console.ReadLine();
            Console.WriteLine();

            string mergedFile = Path.GetDirectoryName(firstExcelFile) + @"\" + "merged.xlsx";

            var firstExcelData = ReadExcelData(firstExcelFile);
            var secondExcelData = ReadExcelData(secondExcelFile);

            Console.ReadLine();   //debug

            var matchedData = GetMatchingRows(firstExcelData, secondExcelData);

            CreateMergedExcel(mergedFile, matchedData);
        }




        private static Dictionary<string, string> ReadExcelData(string filePath)
        {
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.Visible = false;
            Excel.Workbook excelWorkbook = excelApplication.Workbooks.Open(filePath);
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

            excelWorkbook.Close();
            excelApplication.Quit();

            return rowData;
        }




        private static Dictionary<string, string> GetMatchingRows(Dictionary<string, string> firstExcelData, Dictionary<string, string> secondExcelData)
        {
            List<string, string, string> matchedData = new List<string, string, string>();   //new some collection able to hold three string values!

            foreach (var sourceTextAndtranslationRow_ in firstExcelData)
            {
                foreach (var sourceTextAndtranslationRow__ in secondExcelData)
                {
                    if (sourceTextAndtranslationRow_.Key == sourceTextAndtranslationRow__.Key)
                    {
                        matchedData.Add(sourceTextAndtranslationRow_.Key, sourceTextAndtranslationRow_.Value, sourceTextAndtranslationRow__.Value);   //source text + translation first (e.g. German) + translation second (e.g. Swedish)!
                    }
                }



                //var englishText = sourceTextAndtranslationRow["ENGLISH"];
                //var matchingRow = secondExcelData.FirstOrDefault(row => row["ENGLISH"] == englishText);
                //if (matchingRow != null)
                //{
                //    matchedData.Add(new Dictionary<string, string>()
                //{
                //    { "ENGLISH", englishText },
                //    { "GERMAN", sourceTextAndtranslationRow["GERMAN"] },
                //    { "SWEDISH", matchingRow["SWEDISH"] }
                //});
                //}
            }

            return matchedData;
        }




        private static void CreateMergedExcel(string filePath, Dictionary<string, string> data)
        {
            var excelApp = new Excel.Application();
            excelApp.Visible = false;
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.Worksheets[1];

            //Add Headers:
            worksheet.Cells[1, 1] = "Source language";
            worksheet.Cells[1, 2] = "Target language 1";
            worksheet.Cells[1, 3] = "Target language 2";

            int row = 2;
            foreach (var item in data)
            {
                worksheet.Cells[row, 1] = item["ENGLISH"];
                worksheet.Cells[row, 2] = item["GERMAN"];
                worksheet.Cells[row, 3] = item["SWED swedishData"];
                row++;
            }

            workbook.SaveAs(filePath);
            workbook.Close();
            excelApp.Quit();
        }
    }
}




/*
moin! i have different bilingual excel files (english-German excel and english-Swedish excel) and i would like to create a merged excel file with all languages, but only with those rows where the english text is the same in the german and the swedish excel. in other words i want to extract the "smallest common denominator". I want to accomplish this in C# with Excel.interop.

example: this is how the en-de excel looks like
en-de:
     A    				  B			
1   | ENGLISH         | GERMAN
2   | the cat is on the roof  | die katze is auf dem dach
3   | it went boom!      | es machte bumm!
4   | phatman in da room    | fetter mann im raum
5   | why is this so?     | warum ist das so?
6	  | tralala         | lalalala

this is the en-SV excel:
en-SV:
     A    				  B			
1   | ENGLISH         | SWEDISH
2   | the dog is on the roof  | hunden är på taket
3   | the woman is home    | fru är hemma
4   | why is this so?     | varför är det så?
5	  | this is different    | det är annat
6   | it went boom!      | det gjorde boom!

merged result should be:
merged:
         A       				    B			         C
1      | ENGLISH                  | GERMAN             | SWEDISH
2      | it went boom!            | es machte bumm!    | det gjorde boom!
3      | why is this so?          | warum ist das so?  | varför är det så?



Here's how you can achieve the desired outcome using C# with Excel.Interop:
[CODE]
This code defines three functions:

ReadExcelData: This function reads data from an Excel file into a List of Dictionaries. Each dictionary represents a row in the Excel sheet, with key-value pairs corresponding to column headers and their respective values.
GetMatchingRows: This function iterates through the English-German data and searches for matching English text in the English-Swedish data. If a match is found, it creates a new dictionary with English, German, and Swedish translations and adds it to a list of matched data.
CreateMergedExcel: This function creates a new Excel workbook, writes the headers ("ENGLISH", "GERMAN", "SWEDISH") to the first row, and
Källor
info
github.com/iwanowskipawel/CSR_ToExcel



share


more_vert

continue response
edit
Visa utkast
volume_up

(Continuation)

... then iterates through the list of matched data and writes each entry (English, German, Swedish) to a new row in the worksheet. Finally, it saves the workbook as the specified merged file and closes the Excel application.

Points to Consider:

Remember to replace the placeholder file paths with your actual file locations.
This code assumes the first row in both Excel sheets contains headers. You might need to adjust the code if this is not the case.
Error handling can be added to handle situations where files cannot be opened or data is missing.
This code provides a basic framework for merging bilingual Excel files based on matching English text. You can further customize it to suit your specific needs, such as:

Adding formatting to the merged Excel sheet.
Handling cases where there are no matching English entries.
Including additional languages by adding more columns to the merged sheet.
*/