using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.IO;
using System.Globalization;
using System.Threading;
using System.Drawing.Printing;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using ExcelDrawing = OfficeOpenXml.Drawing.ExcelDrawing;

namespace ExcelCharts
{    
    internal class DixelData
    {
        readonly bool printNeeded = false;
        string saveFileDir = null;
        ExcelWorkbook workBook;
        FileInfo fi;
        FileInfo copyOfFi;
        ExcelPackage ep;


        public DixelData(string filePath, bool print)
        {
            try
            {
                fi = new FileInfo(filePath);                
                copyOfFi = fi.CopyTo(Path.GetDirectoryName(filePath) + "\\temp.xlsx");
                ep = new ExcelPackage(copyOfFi);
                workBook = ep.Workbook;
                SetSaveDirectory(filePath);
                printNeeded = print;
            }
            catch (ArgumentException)
            {
                MessageBox.Show("Invalid file path!");               
                return;
            }
            catch (NullReferenceException)
            {
                return;
            }
            catch (Exception e)
            {
                throw new Exception("Object was not created..: " + 
                    Environment.NewLine + 
                    e.ToString());
            }
        }
        private void SetSaveDirectory(string path)
        {
            try
            {
                saveFileDir = Path.GetDirectoryName(path);
            }
            catch (ArgumentException)
            {
                return;
            }
            catch (PathTooLongException)
            {
                MessageBox.Show("File path too long!");
            }
        }        
        /*
        public void CheckChartsTest()
        {
            Sheets xlWSheets;
            try
            {
                xlWSheets = xlWBook.Worksheets;
            }
            catch (Exception)
            {
                throw;
            }
            MainForm.LabelText("Printing....");
            if (xlWSheets != null)
            {
                ChartObjects chObjs;
                int iterations = 0;
                
                foreach(Worksheet ws in xlWSheets)
                {
                    chObjs = ws.ChartObjects();
                    if (chObjs == null)
                    {
                        continue;
                    }
                    
                    foreach(ChartObject chObj in chObjs)
                    {
                        iterations++;
                        Chart ch = chObj.Chart;
                        
                        if (iterations >= 3)
                        {
                            iterations = 0;
                            Thread.Sleep(5000);
                        }
                        
                        if (MainForm.PrintCanceled)
                        {
                            MainForm.LabelText("Print stopped!");
                            return;
                        }
                        /*
                        PrintDocument prDoc = new PrintDocument();
                        string tempPath = Path.GetTempPath() + ch.Name + ".png";
//                        xlApp.Goto(chObj, true);
                        chObj.Select();
                        ch.Export(tempPath, "PNG", false);
                        prDoc.PrintPage += (sender, args) =>
                        {
                            Image i = Image.FromFile(tempPath);
                            System.Drawing.Point p = new System.Drawing.Point(100, 100);
                            args.Graphics.DrawImage(i, 10, 10, i.Width, i.Height);
                        };
                        prDoc.DefaultPageSettings.Landscape = true;
                        prDoc.Print();
                        //*/
                        /*
                        ch.PrintOutEx();
                    }                    
                }
            }
            MainForm.LabelText("Print Finished!");
        }//*/
        public void LoadData()
        {
            List<Thread> treadsCharts = new List<Thread>();
            List<Thread> treadsConv = new List<Thread>();
            Thread load = new Thread(() =>
            {
            //DateTime start = DateTime.Now;
            ExcelWorksheets xlWSheets;
                try
                {
                    xlWSheets = workBook.Worksheets;
                }
                catch (Exception)
                {
                    throw;
                }
                int sheetCount = xlWSheets.Count;
                int sheetNumber = 1;
                
                foreach (ExcelWorksheet xlWSheet in xlWSheets)
                {
                    if (xlWSheet.Cells[xlWSheet.Dimension.Address] == null)
                        continue;
                    if (MainForm.isCancellationRequested)
                    {
                        MessageBox.Show("Stopped!", "!");                        
                        return;
                    }
                    MainForm.ConvProgBar(1, true, sheetNumber, sheetCount);
                    ConvertDateCellsToText(xlWSheet, sheetNumber, sheetCount);
                    Thread trCharts = new Thread(() =>
                    {
                        //if (MainForm.TempCharts)
                          //TempChartRanges(xlWSheet);
                        if (MainForm.HumidCharts)
                          HumidChartRanges(xlWSheet);

                    });
                    treadsCharts.Add(trCharts);
                    sheetNumber++;
                }
                foreach (Thread t in treadsCharts)
                {
                    t.Start();
                    t.Join();
                }
            });
            load.Start();
            load.Join();            
        }
        
        private void HumidChartRanges(ExcelWorksheet xlWSheet)
        {
            if (MainForm.isCancellationRequested)
            {
                //Dispose();
                return;
            }
            int startChartPositionLeft = 100;
            int startChartPositionTop = 100;
            List<ExcelChart> xlCharts = new List<ExcelChart>();
            
            ChartRange ChRange = null;
            ExcelAddressBase usedRange = null;
            try
            {
                usedRange = xlWSheet.Dimension;
            }
            catch(Exception e)
            {
                MainForm.LabelText(e.ToString());
                return;
            }            
            /*
            Range combinedAreas = null;
            if (usedRange.Areas.Count > 1)
            {
                combinedAreas = usedRange.Areas[1];
                for (int i = 2; i <= usedRange.Areas.Count; ++i)
                {
                    combinedAreas = xlApp.Union(combinedAreas, usedRange.Areas[i]);
                }
            }
            if (combinedAreas != null)
                usedRange = combinedAreas;
            //*/
            
            try
            {
                ChRange = new ChartRange('H', xlWSheet, usedRange, printNeeded, MainForm.SpecialCase);
            }
            catch (ArgumentException ae)
            {
                MessageBox.Show(ae.Message);
                return;
            }
            int usedRows = usedRange.Rows - 1;
            MainForm.ProgressBar(usedRows, true);
            
            bool firstDateOFRange = true;
            CultureInfo cInfo = new CultureInfo("bg-BG");
            cInfo.DateTimeFormat.ShortDatePattern = "dd/MM/yyyy";
            cInfo.DateTimeFormat.ShortTimePattern = "hh.mm";
            cInfo.DateTimeFormat.DateSeparator = "/";
            string currDateCell;
            object[,] xlRangeArr = (object[,])xlWSheet.Cells[usedRange.ToString()].Value;
            
            for (int i = 1; i <= usedRows; ++i)
            {
                if (MainForm.isCancellationRequested)
                {
                    return;
                }
                MainForm.ProgressBar(i, false);
                if (xlRangeArr[i, 1] == null)
                {
                    if (!(i == usedRows))
                    {
                        continue;
                    }
                    ChRange.CreateChart(xlWSheet, xlCharts, xlWSheet.Name, startChartPositionLeft, startChartPositionTop);
                    continue;
                }
                //MainForm.WriteIntoLabel("Chart " + ChRange.ChartNumber + " ->  Row: " + ChRange.RowOfRange, 1);
                currDateCell = Convert.ToString(xlRangeArr[i, 1]).Split(new char[0], StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                if (currDateCell.Contains("\'"))
                    currDateCell = currDateCell.Remove(currDateCell.IndexOf('\''), 1);
                DateTime date;
                if (DateTime.TryParse(currDateCell, cInfo, DateTimeStyles.None, out date))
                {                    
                    if (IsFirstDayOfMonth(currDateCell, cInfo))
                    {
                        if (firstDateOFRange && i != usedRows)
                        {
                            ChRange.ExpandRange(i);
                        }
                        else
                        {
                            ChRange.CreateChart(xlWSheet, xlCharts, xlWSheet.Name, startChartPositionLeft, startChartPositionTop);
                            
                            startChartPositionTop += 600;
                            ChRange.StartNewRange(i);
                            firstDateOFRange = true;
                        }
                    }
                    else
                    {
                        ChRange.ExpandRange(i);
                        firstDateOFRange = false;

                        if (i == usedRows)
                        {
                            ChRange.CreateChart(xlWSheet, xlCharts, xlWSheet.Name, startChartPositionLeft, startChartPositionTop);
                            
                            startChartPositionTop += 600;
                            ChRange.StartNewRange(i);
                        }
                        else
                        {
                            string nextCell = Convert.ToString(xlRangeArr[i+1, 1]);
                            if (IsFirstDayOfMonth(nextCell, cInfo))
                            {
                                ChRange.CreateChart(xlWSheet, xlCharts, xlWSheet.Name, startChartPositionLeft, startChartPositionTop);
                                startChartPositionTop += 600;
                                ChRange.StartNewRange(i + 1);
                                firstDateOFRange = true;
                            }
                        }
                    }
                }
                else
                {
                    if (ChRange.EnoughDataForChart())
                    {
                        ChRange.CreateChart(xlWSheet, xlCharts, xlWSheet.Name, startChartPositionLeft, startChartPositionTop);
                        startChartPositionTop += 600;
                        ChRange.StartNewRange(i);
                        firstDateOFRange = true;
                    }
                    ChRange.StartNewRange(i + 1);
                }
                ep.Save();
            }
            ep.Save();//*/
        }
        /*
        private void TempChartRanges(Worksheet xlWSheet)
        {
            if (MainForm.isCancellationRequested)
            {
                Dispose();
                return;
            }
            int startChartPositionLeft = 100;
            int startChartPositionTop = 100;
            ChartObjects xlChartObjs;
            Range usedRange;
            Range firstCol;
            try
            {
                xlChartObjs = xlWSheet.ChartObjects();
                usedRange = xlWSheet.UsedRange;
                firstCol = usedRange.Columns[1];
            }
            catch (Exception)
            {
                Dispose();
                return;
            }
            ChartRange ChRange = null;            

            try
            {
                ChRange = new ChartRange('T', usedRange, printNeeded, MainForm.SpecialCase);
            }
            catch (ArgumentException ae)
            {
                MessageBox.Show(ae.Message);
                return;
            }
            int usedRows = usedRange.Rows.Count;
            
            MainForm.ProgressBar(usedRows, true);
            
            bool firstDateOFRange = true;
            CultureInfo cInfo = new CultureInfo("bg-BG");
            cInfo.DateTimeFormat.ShortDatePattern = "dd/MM/yyyy";
            cInfo.DateTimeFormat.ShortTimePattern = "hh.mm.ss";
            cInfo.DateTimeFormat.DateSeparator = "/";
            string currDateCell;
            
            object[,] xlRangeArr = usedRange.Value;
            
            for (int i = 1; i <= usedRows; ++i)
            {
                if (MainForm.isCancellationRequested)
                {
                    return;
                }
                MainForm.ProgressBar(i, false);
                if(xlRangeArr[i, 1] == null)
                {
                    if (!(i == usedRows))
                    {
                        continue;
                    }
                    ChRange.CreateChart(xlChartObjs, xlWSheet.Name, startChartPositionLeft, startChartPositionTop);
                    continue;
                }    
                currDateCell = Convert.ToString(xlRangeArr[i,1]).Split(new char[0], StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                if(currDateCell.Contains("\'"))
                    currDateCell = currDateCell.Remove(currDateCell.IndexOf('\''),1);
                DateTime date;
                if (DateTime.TryParse(currDateCell, cInfo, DateTimeStyles.None, out date))
                {                    
                    if (date.DayOfWeek == DayOfWeek.Monday)
                    {
                        if (firstDateOFRange && i != usedRows)
                        {
                            ChRange.ExpandRange(i);
                        }
                        else
                        {
                            ChRange.CreateChart(xlChartObjs, xlWSheet.Name, startChartPositionLeft, startChartPositionTop);
                            startChartPositionTop += 600;
                            ChRange.StartNewRange(i);
                            firstDateOFRange = true;
                        }
                    }
                    else
                    {
                        ChRange.ExpandRange(i);
                        firstDateOFRange = false;
                        
                        if (i == usedRows)
                        {
                            ChRange.CreateChart(xlChartObjs, xlWSheet.Name, startChartPositionLeft, startChartPositionTop);
                            startChartPositionTop += 600;
                            ChRange.StartNewRange(i);
                        }
                        else
                        {
                            string nextCell = Convert.ToString(xlRangeArr[i+1, 1]);
                            DateTime nextDate;
                            if (DateTime.TryParse(nextCell, out nextDate) && nextDate.DayOfWeek == DayOfWeek.Monday)
                            {
                                ChRange.CreateChart(xlChartObjs, xlWSheet.Name, startChartPositionLeft, startChartPositionTop);
                                startChartPositionTop += 600;
                                ChRange.StartNewRange(i + 1);
                                firstDateOFRange = true;
                            }
                        }
                    }
                }
                else
                {
                    if (ChRange.EnoughDataForChart())
                    {
                        ChRange.CreateChart(xlChartObjs, xlWSheet.Name, startChartPositionLeft, startChartPositionTop);
                        
                        startChartPositionTop += 600;
                        ChRange.StartNewRange(i);
                        firstDateOFRange = true;
                    }
                    ChRange.StartNewRange(i + 1);
                }
            }
        }
        //*/
        private void ConvertDateCellsToText(ExcelWorksheet sheet, int sheetNumber, int sheetCount)
        {
            ExcelRange usedRange = sheet.Cells[sheet.Dimension.Address];
            MainForm.ConvProgBar(1, true, sheetNumber, sheetCount);
            MainForm.ConvProgBar(usedRange.Rows, true, sheetNumber, sheetCount);
            MainForm.ConvProgBar(0, false, sheetNumber, sheetCount);

            object xlNewRange = null;
            //object[,] test = null;
            try
            {
                xlNewRange = (object[,])usedRange.Value;
                //test = (object[,])xlNewRange;
            }catch(Exception e)
            {
                MessageBox.Show(e.ToString());
                return;
            }
            
            for (int i = 1; i <= usedRange.Rows; ++i)
            {
                if (MainForm.isCancellationRequested)
                {
                    return;
                }
                MainForm.ConvProgBar(i, false, sheetNumber, sheetCount);
                
                    /*DateTime d;
                    if (DateTime.TryParse(Convert.ToString(xlNewRange[i, 1]), out d))
                        xlNewRange[i, 1] = "\'" + xlNewRange[i, 1];//*/
            }
            usedRange.Value = xlNewRange;
            sheet.Cells[sheet.Dimension.Address].Value = usedRange.Value;
            ep.Save();
            //MainForm.ConvProgBar(0, true);
        }
        private bool IsFirstDayOfMonth(string date, CultureInfo cInfo)
        {
            DateTime d;
            if (DateTime.TryParse(date, cInfo, DateTimeStyles.None, out d) && d.Day == 1)
            {
                return true;
            }
            return false;
        }
        private bool IsMonday(string date, CultureInfo cInfo)
        {
            DateTime d;
            if (DateTime.TryParse(date, cInfo, DateTimeStyles.None, out d) && d.DayOfWeek == DayOfWeek.Monday)
            {
                return true;
            }
            return false;
        }
        private bool IsSunday(string date)
        {
            DateTime d;
            if (DateTime.TryParse(date, out d) && d.DayOfWeek == DayOfWeek.Sunday)
            {
                return true;
            }
            return false;
        }
        public void SaveFile()
        {
            MainForm.ProgressBar(0, false);
            MainForm.ConvProgBar(0, false, 1, 1);
            //xlApp.Visible = true;
            try
            {
                MainForm.SaveDialogBox(saveFileDir);
                if (string.IsNullOrEmpty(MainForm.SaveFilePath))
                {
                    MessageBox.Show("File was not saved!");

                    return;
                }
                else
                {
                    try
                    {
                        //FileInfo saveFi = new FileInfo(MainForm.SaveFilePath);
                        ep.Save();
                        //ep.SaveAs(saveFi);
                    }
                    catch(Exception e)
                    {
                        return;
                    }
                }
            }
            catch(Exception e)
            {
                return;
            }
            
        }
        /*
        public void SaveAndClose()
        {
            MainForm.ProgressBar(0, false);
            MainForm.ConvProgBar(0, false, 1, 1);
            //xlApp.Visible = true;
            try
            {
                MainForm.SaveDialogBox(saveFileDir);
                if(string.IsNullOrEmpty(MainForm.SaveFilePath))
                {
                    MessageBox.Show("File was not saved!");
                    
                    return;
                }
                else
                {
                    try
                    {
                        
                        xlWBook.SaveAs(MainForm.SaveFilePath,
                                          Type.Missing,
                                          Type.Missing,
                                          Type.Missing,
                                          false,
                                          false,
                                          XlSaveAsAccessMode.xlExclusive,
                                          false,
                                          false,
                                          Type.Missing,
                                          Type.Missing,
                                          Type.Missing);
                        //xlApp.Visible = false;
                        while (!xlWBook.Saved) { }

                        MessageBox.Show("File saved successfully in \"" + MainForm.SaveFilePath + "\"");
                        xlWBook.Close(false);
                        xlWBooks.Close();
                        //xlApp.Quit();
                        Dispose();
                    }
                    catch (COMException)
                    {
                        Dispose();
                    }
                    return;
                }                
            }
            catch (COMException comEx)
            {
                MessageBox.Show("An exception was thrown while saving the file:" +
                    Environment.NewLine +
                    comEx.ToString());
                xlApp.Quit();
                Dispose();
            }
            catch (Exception e)
            {
                MessageBox.Show("An exception was thrown while saving the file:" +
                    Environment.NewLine +
                    e.ToString());
                xlApp.Quit();
                Dispose();
            }
        }//*/
    }
}
