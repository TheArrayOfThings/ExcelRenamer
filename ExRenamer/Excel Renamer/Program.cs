using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_Renamer
{
    class Program
    {
        static void Main(string[] args)
        {
            [DllImport("user32.dll")]
            static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);
            Excel.Application _excel = new Excel.Application();
            Excel.Workbooks _workbooks = null;
            Excel.Workbook _importBook = null;
            Excel.Worksheet _importSheet = null;
            Excel.Sheets _importSheets = null;
            Excel.Range _usedRange = null;
            string excelFileLocation;
            string attachmentLocation;
            string[] files;
            string[][] AllData;
            int sysIDColumnNumber = 0;
            int numberColumnNumber = 0;

            try
            {
                Console.WriteLine("Paste the full file path for the Excel file:");
                excelFileLocation = Console.ReadLine();
                //excelFileLocation = "C:\\Users\\Ryanf\\Downloads\\u_new_supplier_request.xls";
                Console.WriteLine("Paste location of attachments: ");
                attachmentLocation = Console.ReadLine();

                string start = "";
                while (!(start.Trim().ToLower().Equals("y")))
                {
                    Console.WriteLine("Start? (y/n)");
                    start = Console.ReadLine();
                }
                var startTime = System.DateTime.UtcNow;
                Console.WriteLine("Starting at: " + startTime);

                _workbooks = _excel.Workbooks;
                _importBook = _workbooks.Open(excelFileLocation);
                _importSheets = _importBook.Worksheets;
                _importSheet = _importSheets[1];

                _usedRange = _importSheet.UsedRange;
                AllData = new string[_usedRange.Rows.Count][];

                //Find Sys ID and Number columns
                for (int column = 0; column < _usedRange.Columns.Count; ++column)
                {
                    if (_importSheet.Cells[1, column + 1].Value.ToString().Trim().Equals("Number"))
                    {
                        Console.WriteLine("Number column = " + (column + 1));
                        numberColumnNumber = column + 1;
                        continue;
                    }
                    if (_importSheet.Cells[1, column + 1].Value.ToString().Trim().Equals("Sys ID"))
                    {
                        Console.WriteLine("Sys ID column = " + (column + 1));
                        sysIDColumnNumber = column + 1;
                        continue;
                    }
                    if (sysIDColumnNumber != 0 && numberColumnNumber != 0)
                    {
                        break;
                    }
                }


                //Get all sheet data
                for (int row = 0; row < _usedRange.Rows.Count; ++row)
                {
                    AllData[row] = new string[2];
                    AllData[row][0] = _importSheet.Cells[row + 1, numberColumnNumber].Value.ToString().Trim();
                    AllData[row][1] = _importSheet.Cells[row + 1, sysIDColumnNumber].Value.ToString().Trim();
                }

                files = Directory.GetFiles(attachmentLocation);
                for (int i = 0; i < files.Length; ++i)
                {
                    
                    //For each row (other than headers) find the Sys ID and see if I can match it to the attachments
                    for (int row = 1; row < AllData.Length; ++row)
                    {
                        if (files[i].Contains(AllData[row][1])) {
                            //Attachment and excel match found!
                            Console.WriteLine("Attachment found at row: " + row);
                            var attachPath = files[i].Substring(0, files[i].LastIndexOf("\\") + 1);
                            var newFileName = AllData[row][0] + " (" + AllData[row][1] + ").zip";
                            var oldFileName = files[i].Substring(files[i].LastIndexOf("\\") + 1);
                            if (newFileName.Equals(oldFileName))
                            {
                                continue;
                            }
                            Console.WriteLine("Renaming: " + newFileName + " to " + newFileName);
                            System.IO.File.Move(files[i], attachPath + newFileName);
                        }
                    }
                }
                var endTime = System.DateTime.UtcNow;
                var totalTime = endTime - startTime;
                Console.WriteLine("Finished at " + endTime + ". Program took " + totalTime.Seconds + " seconds.");

            } catch (Exception e)
            {
                Console.WriteLine("Error: " + e.Message);
            }
            Console.WriteLine("Trying to close excel...");
            int id;
            GetWindowThreadProcessId(_excel.Hwnd, out id);
            Process process = Process.GetProcessById(id);
            if (process != null)
            {
                process.Kill();
            }
        }
    }
}
