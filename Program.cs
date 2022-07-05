using System;
using System.Collections.Generic;
using System.Configuration;
using System.Text;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;
using System.Data;
using System.IO;

namespace excelMultipleSheetDeneme
{   
    class Program
    {
        static ExcelControl excelControl = new ExcelControl();
        static Random random = new Random();
        static string folderName = "D:\\TestResults";
        public static void Main(string[] ar)
        {
            bool folderExists = System.IO.Directory.Exists(folderName);
            if (!folderExists)
            {
                System.IO.Directory.CreateDirectory(folderName);
            }

            string dateTime = System.DateTime.UtcNow.ToString();
            dateTime = dateTime.Replace(".", "_");
            dateTime = dateTime.Replace(":", "_");
            dateTime = dateTime.Replace(" ", "-");
            string excelFileName = folderName + "\\TestResults_" + dateTime + ".xlsx";
            bool exists = File.Exists(excelFileName);
            if (exists)
            {
                File.Delete(excelFileName);
            }            

            Application ExcelApp = new Application();
            Workbook ExcelWorkBook = null;
            Worksheet ExcelWorkSheet = null;
            ExcelApp.Visible = true;
            ExcelWorkBook = ExcelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Font excelFont;
            Microsoft.Office.Interop.Excel.Range excelRange;
            Microsoft.Office.Interop.Excel.Borders borders;

            excelControl.CreateExcelDataSet();

            List<string> SheetNames = new List<string>();
            SheetNames.Add("Mode1");
            SheetNames.Add("Mode2");
            SheetNames.Add("Mode3");
            SheetNames.Add("Mode4");
            SheetNames.Add("Mode5");
            SheetNames.Add("Mode6");

            excelControl.CreateNewRowSheet1(1, "evpNum1", "discInp", 1, 5, 2, "Volt", "Pass");
            excelControl.CreateNewRowSheet1(2, "evpNum2", "discInpBuff", 4, 15, 8, "Ampere", "Fail");
            excelControl.CreateNewRowSheet1(3, "evpNum3", "discOut", 3, 4, 8, "Res", "Pass");

            excelControl.CreateNewRowSheet2(1, "evpNum6", "discInp", 1, 5, 2, "Volt", "Pass");
            excelControl.CreateNewRowSheet2(2, "evpNum7", "discInpBuff", 4, 15, 8, "Ampere", "Fail");
            excelControl.CreateNewRowSheet2(3, "evpNum8", "discOut", 3, 4, 8, "Res", "Pass");

            excelControl.CreateNewRowSheet3(1, "evpNum9", "discInp", 1, 5, 2, "Volt", "Pass");
            excelControl.CreateNewRowSheet3(2, "evpNum10", "discInpBuff", 4, 15, 8, "Ampere", "Fail");
            excelControl.CreateNewRowSheet3(3, "evpNum11", "discOut", 3, 4, 8, "Res", "Pass");

            excelControl.CreateNewRowSheet4(1, "evpNum12", "discInp", 1, 5, 2, "Volt", "Pass");
            excelControl.CreateNewRowSheet4(2, "evpNum13", "discInpBuff", 4, 15, 8, "Ampere", "Fail");
            excelControl.CreateNewRowSheet4(3, "evpNum14", "discOut", 3, 4, 8, "Res", "Pass");

            excelControl.CreateNewRowSheet5(1, "evpNum15", "discInp", 1, 5, 2, "Volt", "Pass");
            excelControl.CreateNewRowSheet5(2, "evpNum16", "discInpBuff", 4, 15, 8, "Ampere", "Fail");
            excelControl.CreateNewRowSheet5(3, "evpNum17", "discOut", 3, 4, 8, "Res", "Pass");

            excelControl.CreateNewRowSheet6(1, "evpNum18", "discInp", 1, 5, 2, "Volt", "Pass");
            excelControl.CreateNewRowSheet6(2, "evpNum19", "discInpBuff", 4, 15, 8, "Ampere", "Fail");
            excelControl.CreateNewRowSheet6(3, "evpNum20", "discOut", 3, 4, 8, "Res", "Pass");

            try
            {

                for (int i = 1; i < SheetNames.Count; i++)
                    ExcelWorkBook.Worksheets.Add(); //Adding New sheet in Excel Workbook

                for (int i = 0; i < excelControl.dataSet.Tables.Count; i++)
                {
                    ExcelWorkSheet = ExcelWorkBook.Worksheets[i + 1];
                    
                    //Writing Columns Name in Excel Sheet
                    for (int col = 1; col <= excelControl.dataSet.Tables[i].Columns.Count; col++)
                        ExcelWorkSheet.Cells[excelControl.rowCounterExcel, col] = excelControl.dataSet.Tables[i].Columns[col - 1].ColumnName;

                    excelControl.rowCounterExcel++;

                    //Writing Rows into Excel Sheet
                    for (int row = 0; row < excelControl.dataSet.Tables[i].Rows.Count; row++) //r stands for ExcelRow and col for ExcelColumn
                    {
                        // Excel row and column start positions for writing Row=1 and Col=1
                        for (int col = 1; col <= excelControl.dataSet.Tables[i].Columns.Count; col++)
                        {
                            ExcelWorkSheet.Cells[excelControl.rowCounterExcel, col] = excelControl.dataSet.Tables[i].Rows[row][col - 1].ToString();
                        }                           
                            
                        excelControl.rowCounterExcel++;
                    }

                    ExcelWorkSheet.Name = SheetNames[i];//Renaming the ExcelSheets  

                    excelRange = ExcelWorkSheet.Range[ExcelWorkSheet.Cells[1, 1], ExcelWorkSheet.Cells[1, 8]];
                    excelRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    excelRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    excelRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    excelRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    excelRange.Borders.Color = System.Drawing.Color.Black;
                    excelRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);

                    excelFont = excelRange.Font;
                    excelFont.Size = 14;
                    excelFont.Bold = true;
                    excelFont.Color = System.Drawing.Color.White;

                    excelRange = ExcelWorkSheet.Range[ExcelWorkSheet.Cells[2, 1], ExcelWorkSheet.Cells[excelControl.rowCounterExcel + 1, 7]];
                    excelRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkSalmon);

                    for (int k = 2; k < excelControl.rowCounterExcel; k++)
                    {
                        excelRange = ExcelWorkSheet.Range[ExcelWorkSheet.Cells[k, 8], ExcelWorkSheet.Cells[k, 8]];
                        if (ExcelWorkSheet.Cells[k, 8].Value == "Pass")
                        {
                            excelRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.SeaGreen);
                        }
                        else
                        {
                            excelRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.IndianRed);
                        }
                        
                    }

                    excelRange = ExcelWorkSheet.Range[ExcelWorkSheet.Cells[2, 1], ExcelWorkSheet.Cells[excelControl.rowCounterExcel + 1, 8]];
                    excelRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    excelRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    excelRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    excelRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    excelRange.Borders.Color = System.Drawing.Color.Black;

                    excelFont = excelRange.Font;
                    excelFont.Size = 12;
                    excelFont.Bold = true;
                    excelFont.Color = System.Drawing.Color.White;                    

                    ExcelWorkSheet.Columns.AutoFit();
                    //Change all cells' alignment to center
                    ExcelWorkSheet.Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    
                    excelControl.rowCounterExcel = 1;
                }

                ExcelWorkBook.SaveAs(excelFileName);
                ExcelWorkBook.Close();
                ExcelApp.Quit();

                Marshal.ReleaseComObject(ExcelWorkSheet);
                Marshal.ReleaseComObject(ExcelWorkBook);
                Marshal.ReleaseComObject(ExcelApp);
            }
            catch (Exception exHandle)
            {
                Console.WriteLine("Exception: " + exHandle.Message);
                Console.ReadLine();
            }
            finally
            {
                foreach (Process process in Process.GetProcessesByName("Excel"))
                    process.Kill();
            }
        }
    }
}