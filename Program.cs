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
        static string mergedFile = "";
        static List<string> toBeMergedExcelFilesList = null;


        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Reading paths.txt ...");

                //Console.WriteLine("State filepaths to bilingual Excel files to be merged, format: filepaths, separated by semicolons, no linebreaks.");
                //Console.WriteLine(@"Example: 'C:\otherExcelFiles\en-FR;C:\otherExcelFiles\en-NO;C:\otherExcelFiles\en-SV;C:\otherExcelFiles\en-FI'");
                //string firstExcelFile = @"C:\Users\Bernd\Downloads\Csharp\_EXCEL_ExcelAllLanguagesCommonDenominator\testfiles\en-DE.xlsx";   //debug
                //string firstExcelFile = @"C:\Users\oelll\Dropbox\_ME\III Professionella Expertis\C# Project\_EXCEL_ExcelAllLanguagesCommonDenominator\testfiles\en-DE.xlsx";   //debug

                //step 1 saving all paths to single string var - DONE:
                string toBeMergedExcelFiles = File.ReadAllText(@"C:\Users\Bernd\Downloads\Csharp\_EXCEL_ExcelAllLanguagesCommonDenominator\testfiles\paths.txt");
                //string toBeMergedExcelFiles = Console.ReadLine();

                //step 2 splitting var into list of single paths - DONE:
                toBeMergedExcelFilesList = new List<string>(toBeMergedExcelFiles.Split(';'));   //notice the subtle difference? Split(";") doesnt work ...
                foreach (var item in toBeMergedExcelFilesList)   //debug
                {
                    Console.WriteLine("toBeMergedExcelFilesList: " + item);
                }
                Console.WriteLine();

                //step 3 extract bilingual data of each Excel into SEPARATE dictionary = 'dictionary ID, col A, col B' - PENDING:
                //see far below in comment section for possible solution to this tricky problem!
                int dictID = 0;
                List<Tuple<int, string, string>> collatedData = new List<Tuple<int, string, string>>();
                foreach (string toBeMergedExcelFilesListSinglePath in toBeMergedExcelFilesList)
                {
                    List<Tuple<int, string, string>> dataFromExcel = ReadExcelData(toBeMergedExcelFilesListSinglePath, dictID);   //have to store LIST of tuples in temporary additional list "dataFromExcel" as you cannot add a list to a list, but only the tuples ot the list of tuples!!

                    foreach (Tuple<int, string, string> tuple in dataFromExcel)   //each tuple represents one row of current Excel
                    {
                        collatedData.Add(tuple);
                    }

                    dictID++;
                }
                Console.WriteLine();

                //debug:
                foreach (var item in collatedData)
                {
                    Console.WriteLine(item.Item1 + " " + item.Item2 + " " + item.Item3);
                }

                //comparison algorithm:
                //step 1: determine number of distinctive dictIDs, e.g. 0, 1, 2 or 0, 1, 2, 3, 4, 5 etc.
                //step 2: number determine number of loop cycles


                Console.ReadLine();




                //List<Tuple<string, string, string>> matchedData = GetMatchingRows(toBeMergedExcelFilesData);

                //CreateMergedExcel(mergedFile, matchedData);


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




        private static List<Tuple<int, string, string>> ReadExcelData(string toBeMergedExcelFilesListSinglePath, int dictID)
        {
            Console.WriteLine("Reading Excel files ...");

            string mergedFile = Path.GetDirectoryName(toBeMergedExcelFilesList[0]) + @"\" + "merged_" + DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") + "." + Path.GetExtension(toBeMergedExcelFilesList[0]);

            Excel.Application excelApplication = null;
            Excel.Workbook excelWorkbook = null;

            try
            {
                excelApplication = new Excel.Application();
                excelApplication.Visible = false;
                excelWorkbook = excelApplication.Workbooks.Open(toBeMergedExcelFilesListSinglePath);
                Excel.Worksheet excelWorksheet = excelWorkbook.Worksheets[1];   //assuming that FIRST worksheet is the one

                int lastRow = excelWorksheet.UsedRange.Rows.Count;
                int lastColumn = excelWorksheet.UsedRange.Columns.Count;

                List<Tuple<int, string, string>> rowData = new List<Tuple<int, string, string>>();
                for (int row = 2; row <= lastRow; row++)   //"row = 2" => assuming that first row is headings "English", "German", etc.
                {
                    //Only iterate over columns A and B (1 and 2 in Excel Interop):
                    string key = excelWorksheet.Cells[row, 1].Value2?.ToString();   //captures column A row 2 cell value, then r3, and so on ...
                    string value = excelWorksheet.Cells[row, 2].Value2?.ToString();   //captures column B row 2 cell value, then r3, and so on ...
                    Console.WriteLine("col A | B: " + key + " | " + value);   //debug

                    if (key != null && value != null)   //Add only if both key and value are not null
                    {
                        rowData.Add(new Tuple<int, string, string>(dictID, key, value));   //have to state "new Tuple<..."
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




        //private static List<Tuple<string, string, string>> GetMatchingRows(Dictionary<string, string> toBeMergedExcelFileslData)
        //{
        //    Console.WriteLine("Matching rows ...");

        //    try   //try LINQ instead to avoid unknown number of nested foreach loops and make code scalable
        //    {
        //        List<Tuple<string, string, string>> matchedData = new List<Tuple<string, string, string>>();   //this is how you store multiple values in a list, here as "3-tuples" (more than the merely 2's in a dictionary)

        //        foreach (KeyValuePair<string, string> sourceTextAndtranslationRow_ in initialExcelData)
        //        {
        //            foreach (KeyValuePair<string, string> sourceTextAndtranslationRow__ in toBeMergedExcelFileslData)
        //            {
        //                if (sourceTextAndtranslationRow_.Key == sourceTextAndtranslationRow__.Key)
        //                {
        //                    matchedData.Add(new Tuple<string, string, string>(sourceTextAndtranslationRow_.Key, sourceTextAndtranslationRow_.Value, sourceTextAndtranslationRow__.Value));   //this is how to add Tuples to that Tuple list; source text + translation first (e.g. German) + translation second (e.g. Swedish)!
        //                }
        //            }
        //        }

        //        foreach (Tuple<string, string, string> item in matchedData)
        //        {
        //            Console.WriteLine("matchedDate:" + item.Item1 + " " + item.Item2 + " " + item.Item3);
        //        }
        //        Console.WriteLine();

        //        return matchedData;

        //    }
        //    catch (Exception exception)
        //    {
        //        Console.ForegroundColor = ConsoleColor.Red;
        //        Console.WriteLine(">>> ERROR <<<: " + exception);
        //        Console.ResetColor();

        //        return null;
        //    }
        //}




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




/*
using System;
using System.Collections.Generic;
using System.Linq;

class Program
{
    static void Main(string[] args)
    {
        // Sample data
        List<Tuple<int, string, string>> collatedData = new List<Tuple<int, string, string>>()
        {
            new Tuple<int, string, string>(0, "This is the story of a dog who got lost in the fog", "Dies ist die Geschichte eines Hundes, der sich im Nebel verirrte."),
            new Tuple<int, string, string>(0, "The cat is on the roof", "Die Katze ist auf dem Dach."),
            new Tuple<int, string, string>(0, "Signal improvement", "Signalverbesserung"),
            new Tuple<int, string, string>(0, "Love in times of science", "Liebe in Zeiten der Wissenschaft."),
            new Tuple<int, string, string>(1, "I go to the cinema every Thursday", "wertwert"),
            new Tuple<int, string, string>(1, "Signal improvement", "trwetrw"),
            new Tuple<int, string, string>(1, "Who are you?", "tretre"),
            new Tuple<int, string, string>(1, "The cat is on the roof", "ertertert"),
            new Tuple<int, string, string>(2, "I go to the cinema every Thursday", "Jag gå på bio varje Torsdag."),
            new Tuple<int, string, string>(2, "Signal improvement", "Signalförbättring"),
            new Tuple<int, string, string>(2, "Who are you?", "Vem är du?"),
            new Tuple<int, string, string>(2, "The cat is on the roof", "Katten är på taket.")
        };

        // Dictionaries to store colA strings from each dictID
        Dictionary<int, List<string>> colADict = new Dictionary<int, List<string>>();

        // Extract colA strings for each dictID
        foreach (var tuple in collatedData)
        {
            int dictID = tuple.Item1;
            string colA = tuple.Item2;

            if (!colADict.ContainsKey(dictID))
            {
                colADict[dictID] = new List<string>();
            }

            colADict[dictID].Add(colA);
        }

        // Find common strings in colA across all dictID
        var commonStrings = colADict.Values.Aggregate((prevList, nextList) => prevList.Intersect(nextList).ToList());

        // List to store matched strings along with their respective colB values
        List<Tuple<int, string, string>> matchedStrings = new List<Tuple<int, string, string>>();

        // Add common strings with their colB values to matchedStrings
        foreach (var commonString in commonStrings)
        {
            var matchedTuples = collatedData.Where(tuple => tuple.Item2 == commonString).ToList();
            matchedStrings.AddRange(matchedTuples);
        }

        // Output matched strings
        foreach (var matchedTuple in matchedStrings)
        {
            Console.WriteLine($"{matchedTuple.Item1} | {matchedTuple.Item2} | {matchedTuple.Item3}");
        }
    }
}

*/