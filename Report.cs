using System;
//using ClosedXML.Excel;
using System.Drawing;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Windows.Forms.DataVisualization.Charting;
using SFB.LibraryFunctions;
namespace NGW_SharePoint.Utility
{ 
    class Reports
    {
        /// <summary>
        /// Summary of Exceptions report.
        /// </summary>
        /// <param name="_tname">Displays the test name where exception has occurred</param>
        /// <param name="_message">Displays the Exception message</param>
        /// <param name="_source">Displays the Exception Source</param>
        /// <param name="_stacktrace">Displays the stack trace of exception/param>
        public static void ExceptionReports(string _tname, string _message, string _source, string _stacktrace)
        {
            FileStream fs = null;
            try
            {
                string _exceptionPath = String.Concat(Constants.globalResultsPath + "\\" + "Exceptions");

                if (!Directory.Exists(_exceptionPath))
                {
                    Directory.CreateDirectory(_exceptionPath);

                }
                string path = Path.Combine(_exceptionPath, _tname + "Exception" + ".txt");

                if (!File.Exists(path))
                {
                    var _resultfile = File.Create(path);

                    _resultfile.Close();

                }
             
                //using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
                    fs = new FileStream(path, FileMode.OpenOrCreate);
                using (StreamWriter str = new StreamWriter(fs))
                {
                    fs = null;
                    str.BaseStream.Seek(0, SeekOrigin.End);
                    str.WriteLine("TestName:{0}" + Environment.NewLine + "Exception_Message:{1}" + Environment.NewLine + "Exception_Source:{2}" + Environment.NewLine + Environment.NewLine + "Exception.StackTrace:{3}", _tname, _message, _source, _stacktrace);
                    str.WriteLine(DateTime.Now.ToLongTimeString() + " " + DateTime.Now.ToLongDateString());
                    str.Flush();
                }
                DirectoryInfo d = new DirectoryInfo(_exceptionPath);//Assuming Test is your Folder
                FileInfo[] Files = d.GetFiles("*.txt"); //Getting Text files

                foreach (FileInfo file in Files)
                {
                    name = file.Name;
                }

                string textPath = Path.Combine(_exceptionPath, name);
                var uri = new System.Uri(textPath);
                var converted = uri.AbsoluteUri;
                Constants.ExceptionPath = converted;
            }

            finally
            {
                if (fs != null)
                    fs.Dispose();
            }
        }

        /// <summary>
        /// Exports Data from Gridview  to Excel 2007/2010/2013 format
        /// </summary>
        /// <param name="Title">Title to be shown on Top of Exported Excel File</param>
        /// <param name="HeaderBackgroundColor">Background Color of Title</param>
        /// <param name="HeaderForeColor">Fore Color of Title</param>
        /// <param name="HeaderFont">Font size of Title</param>
        /// <param name="DateRange">Specify if Date Range is to be shown or not.</param>
        /// <param name="FromDate">Value to be stored in From Date of Date Range</param>
        /// <param name="DateRangeBackgroundColor">Background Color of Date Range</param>
        /// <param name="DateRangeForeColor">Fore Color of Date Range</param>
        /// <param name="DateRangeFont">Font Size of Date Range</param>
        /// <param name="gv">GridView Containing Data. Should not be a templated Gridview</param>
        /// <param name="ColumnBackgroundColor">Background Color of Columns</param>
        /// <param name="ColumnForeColor">Fore Color of Columns</param>
        /// <param name="SheetName">Name of Excel WorkSheet</param>
        /// <param name="FileName">Name of Excel File to be Created</param>
        /// <returns>System.String.</returns>
        public static string ExportDataTable2Excel(string Title, Color HeaderBackgroundColor, Color HeaderForeColor, int HeaderFont,bool DateRange, string FromDate, Color DateRangeBackgroundColor, Color DateRangeForeColor, int DateRangeFont, System.Data.DataTable gv, Color ColumnBackgroundColor, Color ColumnForeColor, string SheetName, string FileName)
        {
            System.Data.DataTable _table = gv;
            if (gv != null)
            {
                //creating a new Workbook
                var wb = new ClosedXML.Excel.XLWorkbook();
                // adding a new sheet in workbook
                var ws = wb.Worksheets.Add(SheetName);

                //adding content
                //Title
                ws.Cell("A1").Value = Title;
                //  ws.Cell("A2").Value = "Date of Execution" + Utility.GetCurrentDateTime() ;

                // Date
                // ws.Cell("A2").Value = "Date :" + DateTime.Now.ToString("MM-dd-yyyy") + ' ' + "Total Run" + Constants.TOTALRUN + ' '+  "Total Pass:" + Constants.PASSCOUNT + ' ' + "Total Fail:" +' '+ Constants.FAILCOUNT + " Skipped:" + ' ' + Constants.OTHERSCOUNT+" Error/Abort:" + ' ' + Constants.ERRORCOUNT;
                ws.Cell("A2").Value = "Date of execution:" + DateTime.Now.ToString("MM-dd-yyyy")+ ":::::: Environment: " + Constants.Environment;             
                //add columns
                string[] cols = new string[_table.Columns.Count];
                for (int c = 0; c < _table.Columns.Count; c++)
                {
                    var a = _table.Columns[c].ToString();
                    cols[c] = _table.Columns[c].ToString().Replace('_', ' ');

                }

                int rCnt = (ws.LastRowUsed().RowNumber());

                char StartCharCols = 'A';
                int StartIndexCols = rCnt + 1;
                
                
                #region CreatingColumnHeaders
                for (int i = 1; i <= cols.Length; i++)
                {
                    if (i == cols.Length)
                    {
                        string DataCell = StartCharCols.ToString() + StartIndexCols.ToString();
                        ws.Cell(DataCell).Value = cols[i - 1];
                        ws.Cell(DataCell).WorksheetColumn().Width = cols[i - 1].ToString().Length + 10;
                        ws.Cell(DataCell).Style.Font.Bold = true;
                        ws.Cell(DataCell).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Orange;
                        ws.Cell(DataCell).Style.Font.FontColor = ClosedXML.Excel.XLColor.White;
                    }
                    else
                    {
                        string DataCell = StartCharCols.ToString() + StartIndexCols.ToString();
                        ws.Cell(DataCell).Value = cols[i - 1];
                        ws.Cell(DataCell).WorksheetColumn().Width = cols[i - 1].ToString().Length + 10;
                        ws.Cell(DataCell).Style.Font.Bold = true;
                        ws.Cell(DataCell).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Orange;
                        ws.Cell(DataCell).Style.Font.FontColor = ClosedXML.Excel.XLColor.White;
                        StartCharCols++;
                    }
                }
                #endregion           
                //Merging Header
                string Range = "A1:" + StartCharCols.ToString() + "1";

                ws.Range(Range).Merge();
                ws.Range(Range).Style.Font.FontSize = HeaderFont;
                ws.Range(Range).Style.Font.Bold = true;
                ws.Range(Range).Style.Alignment.SetVertical(ClosedXML.Excel.XLAlignmentVerticalValues.Center);
                ws.Range(Range).Style.Alignment.SetHorizontal(ClosedXML.Excel.XLAlignmentHorizontalValues.Center);
                if (HeaderBackgroundColor != null && HeaderForeColor != null)
                {
                    ws.Range(Range).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.White;
                    ws.Range(Range).Style.Font.FontColor = ClosedXML.Excel.XLColor.Maroon;
                }

                //Style definitions for Date range
                Range = "A2:" + StartCharCols.ToString() + "2";

                ws.Range(Range).Merge();
                ws.Range(Range).Style.Font.FontSize = 10;
                ws.Range(Range).Style.Font.Bold = true;
                ws.Range(Range).Style.Alignment.SetVertical(ClosedXML.Excel.XLAlignmentVerticalValues.Bottom);
                ws.Range(Range).Style.Alignment.SetHorizontal(ClosedXML.Excel.XLAlignmentHorizontalValues.Left);

                //border definitions for Columns
                Range = "A3:" + StartCharCols.ToString() + "3";
                ws.Range(Range).Style.Border.LeftBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                ws.Range(Range).Style.Border.RightBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                ws.Range(Range).Style.Border.TopBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                ws.Range(Range).Style.Border.BottomBorder = ClosedXML.Excel. XLBorderStyleValues.Thin;
                char StartCharData = 'A';
                int StartIndexData = 4;

                //char StartCharDataCol = char.MinValue;
                for (int i = 0; i < _table.Rows.Count; i++)
                {
                    for (int j = 0; j < _table.Columns.Count; j++)
                    {

                        string DataCell = StartCharData.ToString() + StartIndexData;
                        var a = _table.Rows[i][j].ToString();
                        a = a.Replace("&nbsp;", " ");
                        a = a.Replace("&amp;", "&");
                        //check if value is of integer type
                        int val = 0;
                        DateTime dt = DateTime.Now;
                        if (int.TryParse(a, out val))
                        {
                            ws.Cell(DataCell).Value = val;
                        }
                        //check if datetime type
                        else if (DateTime.TryParse(a, out dt))
                        {
                            ws.Cell(DataCell).Value = dt.ToShortDateString();
                        }

                        if(a==globalFunctions._fail)
                        {
                          
                            ws.Cell(DataCell).Style.Font.FontColor = ClosedXML.Excel.XLColor.Red;
                            goto add;
                        }
                        else
                        {

                            goto add;
                        }
                   
                    add:
                        ws.Cell(DataCell).Style.Font.FontSize = 9;
                        ws.Cell(DataCell).SetValue(a);
                        StartCharData++;
                    }
                    StartCharData = 'A';
                    StartIndexData++;
                }

                char LastChar = Convert.ToChar(StartCharData + _table.Columns.Count - 1);
                int TotalRows = _table.Rows.Count + 3;
                Range = "A4:" + LastChar + TotalRows;
                ws.Range(Range).Style.Border.LeftBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                ws.Range(Range).Style.Border.RightBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                ws.Range(Range).Style.Border.TopBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                ws.Range(Range).Style.Border.BottomBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                wb.SaveAs(FileName);
                return "Ok";
            }
            else
            {
                return "Invalid GridView. It is null";
            }
        }

        public static void WriteResultToNotepad()
        {

            //int[] Count= { Constants.passCount, Constants.failCount, Constants.othersCount };
            string[] Text = { "Passed", "Failed", "Skipped" };
            string statement = "";

             int[] Count = { Constants.passCount, Constants.failCount, Constants.othersCount };

            for (int i = 0; i <= 2; i++)
            {
                if (Count[i] > 0)
                {
                    int num = Count[i];
                    double percentage = (double)num / Constants.totalRun;

                     decimal round = Math.Round((decimal)percentage, 2);
                     decimal percent = round * 100;
                     int n = Convert.ToInt32(percent);
                    statement += " " + n + "%" + " " + Text[i] + ",";
                   // statement += " " +percentage * 100 + "%" + " " + Text[i] + ",";
                }

            }
            string FinalStatement = statement.TrimEnd(',');
            string filename = Constants.globalResultsPath + "\\ResultStatus.txt";
            if (!File.Exists(filename))
            {
                File.Create(filename).Dispose();
                if (File.Exists(filename))
                {
                    using (StreamWriter sw = new StreamWriter(filename))
                    {
                        sw.Write(Constants.ReportType + " Nightly Execution Report -" + FinalStatement);// Constants.passCount + " Passed, " + Constants.failCount + " Failed, "+Constants.othersCount + " Skipped");
                    }
                }

            }
        }

        public static void CreatePieChart(int pass, int fail, int skip)
        {
            double[] yValues = { pass, fail, skip };
            string[] xValues = { "Passed", "Failed", "Skipped" };

            Chart chart = new Chart();

            Series series = new Series("Default");
            series.ChartType = SeriesChartType.Pie;
            series["PieLabelStyle"] = "Enabled";
            series.Font = new Font("Arial", 9.0f, FontStyle.Regular);

            Title title = new Title()
            {
                Name = chart.Titles.NextUniqueName(),
                Text = "Test Run Statistics",
                Font = new Font("Trebuchet MS", 12F, FontStyle.Bold),
            };
            chart.Titles.Add(title);
            chart.Series.Add(series);

            ChartArea chartArea = new ChartArea();
            chart.Palette = ChartColorPalette.None;
            chart.PaletteCustomColors = new Color[] { Color.OliveDrab, Color.Crimson, Color.SteelBlue };

            chart.Series["Default"].IsValueShownAsLabel = true;

            chart.Series["Default"].Points.DataBindXY(xValues, yValues);

            foreach (System.Windows.Forms.DataVisualization.Charting.DataPoint point in chart.Series["Default"].Points)
            {
                point.Label = "\n#VALX \t \t#PERCENT{P0}";
            }
            foreach (System.Windows.Forms.DataVisualization.Charting.DataPoint point in chart.Series["Default"].Points)
            {
                if (point.YValues.Length > 0 && (double)point.YValues.GetValue(0) == 0)
                {
                    point.LegendText = point.AxisLabel;//In case you have legend
                    point.AxisLabel = string.Empty;
                    point.Label = string.Empty;
                    point.IsValueShownAsLabel = false;

                }
            }


            chart.ChartAreas.Add(chartArea);

            string filename = Constants.globalResultsPath + "\\Chart.png";
            chart.SaveImage(filename, ChartImageFormat.Png);
        }

        /// <summary>
        /// Gets or sets the test context.
        /// </summary>
        /// <value>The test context.</value>
        public TestContext TestContext
        {

            get
            {

                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }
        /// <summary>
        /// The test context instance
        /// </summary>
        private TestContext testContextInstance;
        private static string name;
    }
}
