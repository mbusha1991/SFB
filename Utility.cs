using System;
using System.Linq;
using System.Web;
using System.IO;
using System.Drawing;
using System.CodeDom;
using System.Security;
using System.Net.Mail;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.Windows.Forms;
using System.Collections.Generic;
using ExcelQ = Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.HtmlControls;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.Office.Interop.Excel;
using System.CodeDom.Compiler;
using VBIDE = Microsoft.Vbe.Interop;
using System.Data;
using System.Text;
using SFB.LibraryFunctions;
using GemBox.Spreadsheet;

namespace NGW_SharePoint.Utility
{
    /* This class is having some general functions which can be used directly at any time */

    /// <summary>
    /// Summary of Utility
    /// </summary>
    partial class Utility
    {     
        public static string browserWindowName;

        private static ExcelQ.Application xlApp;
        //private static int count;

        /// <summary>
        /// Displays the Date Format
        /// </summary>
        /// <param name="Date">Displays the date.</param>
        /// <returns>Current date</returns>
        public static string DateFormat(string Date)
        {
            return Convert.ToDateTime(Date).ToString("dd/MM/yyyy");

        }

        /// <summary>
        /// Displays the time format
        /// </summary>
        /// <param name="Time">Displays the time.</param>
        /// <returns>Current time</returns>
        public static string TimeFormat(string Time)
        {
            return Convert.ToDateTime(Time).ToString("hh:mm");

        }

        /// <summary>
        /// Gets the name of the current  user.
        /// </summary>
        /// <returns>Current user name</returns>
        public static string GetCurrentAdUserName()
        {
            return System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToLower();

        }

        /// <summary>
        /// Gets the past date.
        /// </summary>
        /// <returns>Past date</returns>
        public static string GetPastDate()
        {
            DateTime d1 = DateTime.Now;
            d1 = d1.AddDays(-1);
            return d1.ToString("dd/MM/yyyy");
        }

        /// <summary>
        /// Gets the current date.
        /// </summary>
        /// <returns>Current date.</returns>
        public static string GetCurrentDate()
        {
            return DateTime.Now.ToString("dd/MM/yyyy");

        }

        /// <summary>
        /// Gets the post current date.
        /// </summary>
        /// <returns>Post 2 days Current date</returns>
        public static string GetPostCurrentDate()
        {
            return DateTime.Now.AddDays(2).ToString("dd/MM/YYYY");

        }
        /// <summary>
        /// Gets the post 5 days of current date.
        /// </summary>
        /// <returns>Post 5 days of current date</returns>
        public static string GetPOstCurrentDate5()
        {
            return DateTime.Now.AddDays(5).ToString("dd/MM/yyyy");

        }

        /// <summary>
        /// Gets the current date time.
        /// </summary>
        /// <returns>Current Date Time</returns>
        public static string GetCurrentDateTime()
        {
            return DateTime.Now.ToString("dd/MM/yyyy_hh:mm:ss");

        }

        /// <summary>
        /// Gets the date time post 2 days of current date time
        /// </summary>
        /// <returns>Date time post 2 days of current date</returns>
        public static string GetCurrentDateTime2()
        {
            return DateTime.Now.AddDays(2).ToString("dd/MM/yyyyhh:mm");

        }

        /// <summary>
        /// Creates the document on desktop
        /// </summary>
        public static void CreateDesktopDoc()
        {
            string desktoppath = Environment.GetFolderPath(System.Environment.SpecialFolder.DesktopDirectory);
            string Docpath = "TestDoc.Doc";
            string path = Path.Combine(desktoppath, Docpath);
            if (!File.Exists(path))
            {
                using (FileStream f = new FileStream(path, FileMode.Create, FileAccess.ReadWrite))
                {
                    using (StreamWriter sw = new StreamWriter(path))
                    {
                        sw.WriteLine("TestDocument");
                        sw.Flush();
                        //sw.Dispose();
                        // sw.Close();

                    }
                }

            }
        }

        /// <summary>
        /// Searches the grid cell by value.
        /// </summary>
        /// <param name="table">The table.</param>
        /// <param name="searchvalue">The searchvalue.</param>
        /// <returns><c>true</c> if row contains search value, <c>false</c> otherwise.</returns>
        public static bool SearchGridCellByValue(HtmlTable table, string searchvalue)
        {
            foreach (HtmlRow row in table.Rows)
            {
                if (row.InnerText.Contains(searchvalue))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Gets the name of the function.
        /// </summary>
        /// <returns>Calling method name</returns>
        public static string GetFunctionName()
        {
            // Get call stack
            StackTrace stackTrace = new StackTrace();

            // Get calling method name
            return (stackTrace.GetFrame(1).GetMethod().Name);

        }

        /// <summary>
        /// Gets the caller method.
        /// </summary>
        /// <returns>Caller method name.</returns>
        public static string GetCallerMethod()
        {

            // Get call stack
            StackTrace stackTrace = new StackTrace();

            // Get calling method name
            return (stackTrace.GetFrame(0).GetMethod().Name);


        }

        /// <summary>
        /// Starts the default calculator.
        /// </summary>
        public static void StartCalculator()
        {

            Process.Start("calc.exe");
        }

        /// <summary>
        /// Creates the excel.
        /// </summary>
        /// <returns>Excel file</returns>
        //public static string CreateExcel()
        //{

        //    string excelReportPath;
        //    string excelReportFileName;
        //    string ProjectName = Assembly.GetCallingAssembly().GetName().Name;
        //    Excel.Application ExcelApp = new Excel.Application();


        //    Excel.Worksheet ExcelWorkSheet = null;

        //    //Excel.Name Contents_Table = null;
        //    ExcelApp.Visible = false;
        //    excelReportPath = string.Concat(@"C:\Users\mun7cob\Desktop\Manoj_Kumar_M\Results\", ProjectName + '\\' + DateTime.Today.ToString("MMddyyyy"), "\\", "Result", '\\', DateTime.Now.ToString("MMddyyyyHHmmss") + "TestReport");
        //    Excel.Workbook ExcelWorkBook = ExcelApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);

        //    try
        //    {

        //        ExcelWorkSheet = ExcelWorkBook.Worksheets[1];  // Compulsory Line in which sheet you want to write data
        //        ExcelWorkBook.Worksheets[1].Name = "TestReport";//Renaming the Sheet1 to TestReport



        //        if (!Directory.Exists(excelReportPath))
        //        {

        //            Directory.CreateDirectory(excelReportPath);


        //        }
        //        excelReportFileName = Path.Combine(excelReportPath, "TestLog" + DateTime.Now.ToString("ddmmyyyyhhmmss") + ".xlsx");

        //        File.Copy(@"C:\C:\Users\mun7cob\Desktop\Manoj_Kumar_M\TestLog.xlsx", excelReportFileName);
        //        ExcelWorkBook.SaveAs(excelReportFileName);
        //        ExcelWorkBook.Close();


        //        string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
        //            "Data Source={0};Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=0'", excelReportFileName);


        //        using (OleDbConnection cn = new OleDbConnection(connectionString))
        //        {
        //            cn.Open();



        //            using (OleDbCommand cmd1 = new OleDbCommand("Create table [TestReport$] (ProjectName String, TestCase Varchar,Action Varchar,Status varChar,TimeStamp string ,TestUser String,FailureSnapshot String)", cn))
        //            {



        //                cmd1.ExecuteNonQuery();
        //            }

        //        }

        //        ExcelApp.Quit();

        //        Marshal.ReleaseComObject(ExcelWorkSheet);

        //        Marshal.ReleaseComObject(ExcelWorkBook);

        //        Marshal.ReleaseComObject(ExcelApp);

        //        //BorderAround(Contents_Table.Cells, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(79, 129, 189)));
        //        return excelReportFileName;
        //    }

        //    catch (System.Exception exHandle)
        //    {

        //        Console.WriteLine("Exception: " + exHandle.Message);

        //        Console.ReadLine();
        //        throw;
        //    }
        //    finally
        //    {
        //        foreach (Process process in Process.GetProcessesByName("Excel"))

        //            process.Kill();
        //    }
        //}
        /// <summary>
        /// Captures the screen shot of failure
        /// </summary>
        /// <returns>Path of the screen shot.</returns>
        public static string FailureScreenShotCapture()
        {
            string failuresnapshotpath = String.Concat(NGW_SharePoint.Utility.Constants.globalResultsPath + "\\" + "Image");
            string name = null;

            if (!Directory.Exists(failuresnapshotpath))
            {

                Directory.CreateDirectory(failuresnapshotpath);

            }
            Image img = UITestControl.Desktop.CaptureImage();

            img.Save(Path.Combine(failuresnapshotpath, "Img" + DateTime.Now.ToString("dd_mm_yyyy_hh_mm_ss") + ".png"));

            DirectoryInfo d = new DirectoryInfo(failuresnapshotpath);//Assuming Test is your Folder
            FileInfo[] Files = d.GetFiles("*.png"); //Getting Text files

            foreach (FileInfo file in Files)
            {
                name = file.Name;
            }

            string path = Path.Combine(failuresnapshotpath, name);


            var uri = new Uri(path);
            var converted = uri.AbsoluteUri;
            return converted;

        }
        /// <summary>
        /// Reads the excel.
        /// </summary>
        /// <returns>List of test cases</returns>
        public static List<String> ReadExcel()
        {
            List<String> TestCaseList = new List<String>();
            string s;
            string fileName = string.Concat(@"C:\Users\mun7cob\Desktop\Manoj_Kumar_M\TestLog.xlsx");
            string conn = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                     "Data Source='{0}';Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=0'", fileName);
            using (OleDbConnection connection = new OleDbConnection(conn))
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("select * from [TestReport$]", connection);
                using (OleDbDataReader dr = command.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        var row1Col0 = dr[0];
                        s = row1Col0.ToString();
                        TestCaseList.Add(s);

                    }
                }
            }

            return TestCaseList;
        }
        /// <summary>
        /// Gets the name of the WebBrowser.
        /// </summary>
        /// <returns>Web Browser Name</returns>
        /// 
        public static bool IsOdd(int value)
        {
            return value % 2 != 0;
        }
        public static string GetWebBrowserName()
        {
            string WebBrowserName = string.Empty;

            WebBrowserName = HttpContext.Current.Request.Browser.Browser;
            return WebBrowserName;
        }
        /// <summary>
        /// Copy from one directory to other.
        /// </summary>
        /// <param name="sourceDirName">Name of the source dir.</param>
        /// <param name="destDirName">Name of the dest dir.</param>
        /// <param name="copySubDirs">if set to <c>true</c> [copy sub dirs].</param>
        /// <exception cref="DirectoryNotFoundException">Source directory does not exist or could not be found: 
        ///                     + sourceDirName</exception>
        public static void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs)
        {
            // Get the subdirectories for the specified directory.
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);
            DirectoryInfo[] dirs = dir.GetDirectories();

            if (!dir.Exists)
            {
                throw new DirectoryNotFoundException(
                    "Source directory does not exist or could not be found: "
                    + sourceDirName);
            }

            // If the destination directory doesn't exist, create it. 
            if (!Directory.Exists(destDirName))
            {
                Directory.CreateDirectory(destDirName);
            }

            // Get the files in the directory and copy them to the new location.
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                string temppath = Path.Combine(destDirName, file.Name);
                file.CopyTo(temppath, false);
            }

            // If copying subdirectories, copy them and their contents to new location. 
            if (copySubDirs)
            {
                foreach (DirectoryInfo subdir in dirs)
                {
                    string temppath = Path.Combine(destDirName, subdir.Name);
                    DirectoryCopy(subdir.FullName, temppath, copySubDirs);
                }
            }
        }
        /// <summary>
        /// Sends the attached email.
        /// </summary>
        /// <param name="EmailSender">The email sender.</param>
        /// <param name="to">To.</param>
        /// <param name="subject">The subject.</param>
        /// <param name="path">The path.</param>
        public static void SendAttachedEmail(string EmailSender, string to, string subject, string path)
        {
            //To connect the SMTP server for mailing.
            try
            {
                MailMessage objMailMsg = new MailMessage(EmailSender, to);

                objMailMsg.Subject = subject;
                System.Net.Mail.Attachment objAttachment = new System.Net.Mail.Attachment(path);
                objMailMsg.Attachments.Add(objAttachment);
                System.Net.Mail.SmtpClient objSmtpClient = new SmtpClient("192.168.10.1", 25);
                objSmtpClient.Send(objMailMsg);
            }
            catch (System.Exception)
            {
                throw;
            }
        }
        /// <summary>
        /// Screeshtots this instance.
        /// </summary>
        public static void Screeshot(string path)
        {
            Bitmap bitmap = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);

            Graphics graphics = Graphics.FromImage(bitmap as System.Drawing.Image);

            graphics.CopyFromScreen(0, 0, 0, 0, bitmap.Size);

            using (var m = new MemoryStream())
            {
                bitmap.Save(m, ImageFormat.Png);

                var img = System.Drawing.Image.FromStream(m);

                //TEST
                img.Save(path);

            }

        }
        /// <summary>
        /// Displays the first column.
        /// </summary>
        /// <param name="filename">The filename.</param>
        /// <returns>System.String[].</returns>
        public static string[] FirstColumn(string filename)
        {
            Microsoft.Office.Interop.Excel.Application xlsApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlsApp == null)
            {
                Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                return null;
            }

            //Displays Excel so you can see what is happening
            //xlsApp.Visible = true;

            Microsoft.Office.Interop.Excel.Workbook wb = xlsApp.Workbooks.Open(filename,
                0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true);
            Sheets sheets = wb.Worksheets;
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)sheets.get_Item(1);

            Range firstColumn = ws.UsedRange.Columns[1];
            System.Array myvalues = (System.Array)firstColumn.Cells.Value;
            string[] strArray = myvalues.OfType<object>().Select(o => o.ToString()).ToArray();
            return strArray;
        }
        /// <summary>
        /// Loops the through excel.
        /// </summary>
        /// <param name="_filename">The _filename.</param>
        /// <param name="_sheetName">Name of the _sheet.</param>
        /// <param name="_columnName">Name of the _column.</param>
        /// <param name="_comparisonValue">The _comparison value.</param>
        //public static void LoopThroughExcel(string _filename, string _sheetName, string _columnName, string _comparisonValue)
        //{

        //    Excel.Application xlApp;
        //    Excel.Workbook xlWorkBook;
        //    Excel.Worksheet xlWorkSheet;
        //    Excel.Range range;

        //    //string str;
        //    int rCnt = 0;
        //    int cCnt = 0;

        //    xlApp = new Excel.Application();
        //    xlWorkBook = xlApp.Workbooks.Open(_filename, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 0);
        //    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(_sheetName);

        //    range = xlWorkSheet.UsedRange;

        //    rCnt = range.Rows.Count;
        //    cCnt = range.Columns.Count;
        //    Console.WriteLine(rCnt);
        //    Console.WriteLine(cCnt);
        //    int m = 1;
        //    object[,] array = new object[rCnt, cCnt];
        //    array = range.Value;
        //    for (int j = 1; j <= cCnt; j++)
        //    {
        //        for (int i = 1; i <= rCnt; i++)
        //        {
        //            if (array[i, j].ToString() == _comparisonValue)
        //            {
        //                for (m = m + i; m < rCnt; m++)
        //                {
        //                    //if(Convert.ToInt32(array[m,j].ToString())>)
        //                    if ((array[m, j].ToString()) == _comparisonValue)
        //                    {
        //                        array[m, j + 1] = "NO";
        //                    }
        //                    range.Value = array;
        //                }
        //            }
        //        }
        //    }
        //    xlWorkBook.Save();
        //    xlWorkBook.Close();
        //    xlApp.Application.Quit();

        //    releaseObject(xlWorkSheet);
        //    releaseObject(xlWorkBook);
        //    releaseObject(xlApp);

        //    GC.Collect();
        //    GC.WaitForPendingFinalizers();
        //    GC.Collect();
        //    GC.WaitForPendingFinalizers();
        //}
        /// <summary>
        /// Releases the object.
        /// </summary>
        /// <param name="obj">The object.</param>
        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (System.Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        /// <summary>
        /// Lists to data table.
        /// </summary>
        /// <param name="testname">The testname.</param>
        /// <param name="status">The status.</param>
        /// <param name="link">The link.</param>
     
        //public static void ListToDataTable(string testname, string status, string link, string dataRow)
        //{
        //    if (Constants.reportFlag == false)
        //    {
        //        Constants.globalTable.Columns.Add("Module", typeof(string));
        //        Constants.globalTable.Columns.Add("TestCase", typeof(string));
        //        Constants.globalTable.Columns.Add("TestStatus", typeof(string));
        //        Constants.globalTable.Columns.Add("Step No", typeof(string));
        //        Constants.globalTable.Columns.Add("Steps", typeof(String));
        //        Constants.globalTable.Columns.Add("StepStatus", typeof(string));
        //        Constants.globalTable.Columns.Add("FailedScreenLink", typeof(string));
        //        Constants.globalTable.Columns.Add("ExceptionLink", typeof(string));
        //        Constants.globalTable.Columns.Add("TimeStamp", typeof(string));
        //        Constants.globalTable.Columns.Add("TestUser", typeof(string));
        //        Constants.reportFlag = true;
        //        goto add;
        //    }

        //    else
        //    {

        //        goto add;
        //    }

        //    add:
        //    // Here we add all the DataRows
        //    for (int j = 0; j < Library.listOfTuples.Count; j++)
        //    {
        //        Constants.globalTable.Rows.Add(Constants.ModuleName, testname, status, "DataRow:" + dataRow);
        //        //Constants.Constants.globalTable.Rows.Add(Constants.searchModuleName, testname, status, "DataRow:" + dataRow);

        //        int counter = 1;

        //        for (int i = 0; i < Library.listOfTuples.Count; i++)
        //        {

        //            if (Library.listOfTuples[i].Item4 == "Failed")
        //            {


        //                Constants.globalTable.Rows.Add(null, null, null, String.Concat("Step" + counter), Library.listOfTuples[i].Item3, Library.listOfTuples[i].Item4, link, Constants.ExceptionPath, Library.listOfTuples[i].Item5, GetCurrentAdUserName());
        //            }

        //            else
        //            {
        //                Constants.globalTable.Rows.Add(null, null, null, String.Concat("Step" + counter), Library.listOfTuples[i].Item3, Library.listOfTuples[i].Item4, null, null, Library.listOfTuples[i].Item5, GetCurrentAdUserName());

        //            }
        //            counter++;

        //        }
        //        int remove = Math.Max(0, Library.listOfTuples.Count);
        //        Library.listOfTuples.RemoveRange(0, remove);
        //        Constants.stepsCounter = counter - 1;
        //    }

        //}

        public static void ListToDataTable(string testname, string status, string link, string dataRow)
        {
            if (Constants.reportFlag == false)
            {
                Constants.globalTable.Columns.Add("Module", typeof(string));
                Constants.globalTable.Columns.Add("TestCase", typeof(string));

                Constants.globalTable.Columns.Add("TestStatus", typeof(string));
                Constants.globalTable.Columns.Add("Iteration", typeof(string));
                Constants.globalTable.Columns.Add("Step No", typeof(String));
                Constants.globalTable.Columns.Add("Steps", typeof(string));
                Constants.globalTable.Columns.Add("StepStatus", typeof(string));
                Constants.globalTable.Columns.Add("FailedScreenLink", typeof(string));
                Constants.globalTable.Columns.Add("ExceptionLink", typeof(string));
                Constants.globalTable.Columns.Add("TimeStamp", typeof(string));
                Constants.globalTable.Columns.Add("TestUser", typeof(string));
                Constants.reportFlag = true;
                goto add;
            }
            else
            {
                goto add;
            }
            add:
            // Here we add all the DataRows
            for (int j = 0; j < globalFunctions.listOfTuples.Count; j++)
            {
                Constants.globalTable.Rows.Add(Constants.ModuleName, testname, status, "" + dataRow);
                //Constants.Constants.globalTable.Rows.Add(Constants.searchModuleName, testname, status, "DataRow:" + dataRow);
               int counter = 1;

                for (int i = 0; i < globalFunctions.listOfTuples.Count; i++)
                {
                    if (i > 0)
                    {
                        int previous_iteration = Convert.ToInt32(globalFunctions.listOfTuples[i - 1].Item6);
                        int current_iteration = Convert.ToInt32(globalFunctions.listOfTuples[i].Item6);
                        if (previous_iteration < current_iteration)
                        {
                            Constants.globalTable.Rows.Add(null, null, null, null, null, null, null, null, null);
                            counter = 1;
                        }
                    }
                    if (globalFunctions.listOfTuples[i].Item4 == "Failed")
                        {

                            Constants.globalTable.Rows.Add(null, null, null, globalFunctions.listOfTuples[i].Item6, String.Concat("Step" + counter), globalFunctions.listOfTuples[i].Item3, globalFunctions.listOfTuples[i].Item4, link, Constants.ExceptionPath, globalFunctions.listOfTuples[i].Item5, GetCurrentAdUserName());
                        }
                        else
                        {
                            Constants.globalTable.Rows.Add(null, null, null, globalFunctions.listOfTuples[i].Item6, String.Concat("Step" + counter), globalFunctions.listOfTuples[i].Item3, globalFunctions.listOfTuples[i].Item4, null, null, globalFunctions.listOfTuples[i].Item5, GetCurrentAdUserName());

                        }
                    counter++;

                }
                int remove = Math.Max(0, globalFunctions.listOfTuples.Count);
                globalFunctions.listOfTuples.RemoveRange(0, remove);
                Constants.stepsCounter = counter - 1;
            }

        }
        /// <summary>
        /// Creates the directory.
        /// </summary>
        public static void createDirectory()
        {
            if (!Directory.Exists(Constants.globalResultsPath))
            {

                Directory.CreateDirectory(Constants.globalResultsPath);

            }
        }
        public static string ToLiteral(string input)
        {
            using (var writer = new StringWriter())
            {
                using (var provider = CodeDomProvider.CreateProvider("CSharp"))
                {
                    provider.GenerateCodeFromExpression(new CodePrimitiveExpression(input), writer, null);
                    return writer.ToString();
                }
            }
        }
        public static string EncodeTo64(string toEncode)
        {
            byte[] toEncodeAsBytes
                  = System.Text.ASCIIEncoding.ASCII.GetBytes(toEncode);
            string returnValue
                  = System.Convert.ToBase64String(toEncodeAsBytes);
            return returnValue;
        }
        /// <summary>
        /// Gets the elapsed time.
        /// </summary>
        public static void GetElapsedTime()
        {
            //   Stopwatch sw = new Stopwatch();
            //sw.Start();
            //Thread.Sleep(10000);
            //    double a=sw.Elapsed.Seconds;
            //    int b=(int)Math.Ceiling(a);
            //    Console.WriteLine(b);



            //    Stopwatch sw = new Stopwatch();
            //    sw.Start();
            //    Thread.Sleep(10000);
            //    double a = sw.Elapsed.Seconds;

            //    Console.WriteLine(Math.Round(a, 0, MidpointRounding.AwayFromZero));

            //    Stopwatch sw = new Stopwatch();
            //    sw.Start();
            //    Thread.Sleep(10000);
            //    double a = sw.Elapsed.Seconds;

            //    Console.WriteLine(Math.Round(a, 0, MidpointRounding.ToEven));
        }
        public static SecureString MakeSecureString(string text)
        {
            SecureString secure = new SecureString();
            foreach (char c in text)
            {
                secure.AppendChar(c);
            }

            return secure;
        }
        public static BrowserWindow RunAs(string path, string username, string password, string argument, string domain)
        {
            StackTrace stackTrace = new StackTrace();
            try
            {
                ProcessStartInfo myProcess = new ProcessStartInfo(path);
                myProcess.WindowStyle = ProcessWindowStyle.Minimized;
                FileInfo fileInfo = new FileInfo(path);

                myProcess.WorkingDirectory = fileInfo.DirectoryName;
                myProcess.UserName = username;
                myProcess.Password = MakeSecureString(password);
                myProcess.UseShellExecute = false;

                myProcess.LoadUserProfile = true;
                if (Constants.Environment != "d1")
                {
                    myProcess.Domain = domain;
                }
                else
                {
                    myProcess.Domain = "apac";
                }
                string userNameWithDomain = string.Concat(domain + "\\" + username);
                
                //Console.WriteLine(userNameWithDomain);
                //myProcess.Arguments = "https://sites.inside-share-d2.bosch.com/sites/000251/default.aspx  -extoff";
                myProcess.Arguments = String.Concat(argument + "   -extoff");
                Process.Start(myProcess);



                // Playback.PlaybackSettings.ContinueOnError = true;
                string currentUserName = GetCurrentAdUserName();

                // Console.WriteLine(currentUserName);
                if (currentUserName != userNameWithDomain && domain!="de")
                {
                    Playback.PlaybackSettings.MaximumRetryCount = 0;
                    Playback.PlaybackSettings.ContinueOnError = true;

                    WinWindow windowSecurity = new WinWindow();
                    windowSecurity.SearchProperties.Add(WinWindow.PropertyNames.ClassName, "#32770", WinWindow.PropertyNames.Name, "Windows Security");

                    UITestControl userNameEdit = new WinEdit(windowSecurity);
                    userNameEdit.SearchProperties.Add(WinEdit.PropertyNames.ClassName, "Edit");
                    userNameEdit.WaitForControlReady();
                    UITestControlCollection controlCol = userNameEdit.FindMatchingControls();
                    UITestControl passwordEdit = new WinEdit(windowSecurity);
                    //Select the Name property from Collection via LINQ
                    var controls = controlCol.Select(x => x.ClassName);
                    passwordEdit = controlCol.ElementAt(1);
                    userNameEdit.SetFocus();
                    Keyboard.SendKeys(username);
                    passwordEdit.SetFocus();
                    Keyboard.SendKeys(password);
                    WinButton okButton = new WinButton(windowSecurity);
                    okButton.SearchProperties.Add(WinButton.PropertyNames.Name, "OK");
                    Mouse.Click(okButton);

                    Playback.PlaybackSettings.ContinueOnError = false;
                    Playback.PlaybackSettings.MaximumRetryCount = 10;
                }
                WinWindow _window = new WinWindow();
                _window.SearchProperties.Add(WinWindow.PropertyNames.ClassName, "IEFrame", WinWindow.PropertyNames.ControlType, "Window");
                browserWindowName = _window.GetProperty("Name").ToString();
                var browserTab = BrowserWindow.Locate(browserWindowName);
                //browserTab.DrawHighlight();

                Constants.getCurrentMethod = stackTrace.GetFrame(0).GetMethod().Name;
                Constants.getCallingMethod = stackTrace.GetFrame(1).GetMethod().Name;
                Constants.logText = string.Concat("IE run with " + username + " and URL " + argument);
                globalFunctions.listOfTuples.Add(new Tuple<String, String, String,String, String, String>(NGW_SharePoint.Utility.Constants.getCurrentMethod, NGW_SharePoint.Utility.Constants.getCallingMethod, NGW_SharePoint.Utility.Constants.logText, "Passed", NGW_SharePoint.Utility.Utility.GetCurrentDateTime(),null));
                return browserTab;
            }

            catch (Exception)
            {

                Constants.getCurrentMethod = stackTrace.GetFrame(0).GetMethod().Name;
                Constants.getCallingMethod = stackTrace.GetFrame(1).GetMethod().Name;
                Constants.logText = string.Concat("Run as different user failed");
                globalFunctions.listOfTuples.Add(new Tuple<String, String, String,String, String, String>(NGW_SharePoint.Utility.Constants.getCurrentMethod, NGW_SharePoint.Utility.Constants.getCallingMethod, NGW_SharePoint.Utility.Constants.logText, "Failed", NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), null));
                throw;
            }

        }

        public static BrowserWindow RunAsWithoutArgument(string path, string username, string password, string argument, string domain)
        {
            StackTrace stackTrace = new StackTrace();
            try
            {

                ProcessStartInfo myProcess = new ProcessStartInfo(path);
                myProcess.WindowStyle = ProcessWindowStyle.Minimized;
                FileInfo fileInfo = new FileInfo(path);

                myProcess.WorkingDirectory = fileInfo.DirectoryName;
               // myProcess.UserName = username;
                myProcess.Password = MakeSecureString(password);
                myProcess.UseShellExecute = false;

                myProcess.LoadUserProfile = true;
                //myProcess.Domain = domain;
               // string userNameWithDomain = string.Concat(domain + "\\" + username);
                //  Console.WriteLine(userNameWithDomain);
                //  myProcess.Arguments = "https://sites.inside-share-d2.bosch.com/sites/000251/default.aspx  -extoff";
               // myProcess.Arguments = String.Concat(argument + "   -extoff");
                Process.Start(myProcess);

                // Playback.PlaybackSettings.ContinueOnError = true;
                string currentUserName = GetCurrentAdUserName();
                // Console.WriteLine(currentUserName);

               BrowserWindow browserTab = null;

                NGW_SharePoint.Utility.Constants.getCurrentMethod = stackTrace.GetFrame(0).GetMethod().Name;
                NGW_SharePoint.Utility.Constants.getCallingMethod = stackTrace.GetFrame(1).GetMethod().Name;
                NGW_SharePoint.Utility.Constants.logText = string.Concat("IE run with " + username + " and URL " + argument);
                globalFunctions.listOfTuples.Add(new Tuple<String, String, String, String, String, String>(NGW_SharePoint.Utility.Constants.getCurrentMethod, NGW_SharePoint.Utility.Constants.getCallingMethod, NGW_SharePoint.Utility.Constants.logText, "Passed", NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), null));
                return browserTab;
            }

            catch (Exception)
            {

                NGW_SharePoint.Utility.Constants.getCurrentMethod = stackTrace.GetFrame(0).GetMethod().Name;
                NGW_SharePoint.Utility.Constants.getCallingMethod = stackTrace.GetFrame(1).GetMethod().Name;
                NGW_SharePoint.Utility.Constants.logText = string.Concat("Run as different user failed");
                globalFunctions.listOfTuples.Add(new Tuple<String,String, String, String, String, String>(NGW_SharePoint.Utility.Constants.getCurrentMethod, NGW_SharePoint.Utility.Constants.getCallingMethod, NGW_SharePoint.Utility.Constants.logText, "Failed", NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), null));
                throw;
            }

        }
        public static void OutLookReference()
        {

            WinWindow ww = new WinWindow();
            ww.SearchProperties.Add(WinWindow.PropertyNames.ControlType, "Client");
            ww.SearchProperties.Add(WinWindow.PropertyNames.Name, "Inbox", PropertyExpressionOperator.Contains);
            ww.SetFocus();


        }
        public void GenerateGUID()
        {
            Guid g;
            // Create and display the value of two GUIDs.
            g = Guid.NewGuid();
            Console.WriteLine(g);
            Console.WriteLine(Guid.NewGuid());
        }
        public static void LaunchOutlook()
        {

            if (Process.GetProcessesByName("Outlook").Length > 0)
            {


                //Identifying the Outlook Window
                OutLookReference();


            }
            else
            {

                foreach (var process in Process.GetProcessesByName("outlook.exe"))
                {
                    process.Kill();
                }


                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = Path.Combine(@"C:\Program Files (x86)\Microsoft Office\Office15", "outlook.exe");
                startInfo.WorkingDirectory = @"C:\Program Files (x86)\Microsoft Office\Office15";
                Process p = Process.Start(startInfo);
                // Playback.Wait(10000);
                OutLookReference();
            }
        }
        public static void ListAddition(string testname, string status, string iteration, int stepsCounter)
        {
          
           Constants.resultsOverviewTuple.Add(new Tuple<String, String, String, int, String,String>(Constants.ModuleName, testname, iteration, stepsCounter, status, GetCurrentAdUserName())); 
        }
        public static void ListToDataTableConverter()
        {
            System.Data.DataTable dataTableShort = new System.Data.DataTable();
            dataTableShort.Columns.Add("ModuleName", typeof(string));
            dataTableShort.Columns.Add("TestCaseName", typeof(string));
            dataTableShort.Columns.Add("Iteration", typeof(string));
            dataTableShort.Columns.Add("No of steps from Script", typeof(int));
            dataTableShort.Columns.Add("TestStatus", typeof(string));
            dataTableShort.Columns.Add("TestUser", typeof(string));

            // Here we add all the DataRows
            for (int i = 0; i < Constants.resultsOverviewTuple.Count; i++)
            {
                dataTableShort.Rows.Add(Constants.resultsOverviewTuple[i].Item1, Constants.resultsOverviewTuple[i].Item2, Constants.resultsOverviewTuple[i].Item3, Constants.resultsOverviewTuple[i].Item4, Constants.resultsOverviewTuple[i].Item5, Constants.resultsOverviewTuple[i].Item6);
            }          
            string htmlText = toHTML_Table(dataTableShort);
            //Path is combined along with the file name 
            string shortReportPath = Path.Combine(Constants.globalResultsPath, "Result_Overview" + ".html");
            string RecentshortReportPath = Path.Combine(Constants.globalRecentResultsPath, "Result_Overview" + ".html");

            //Creating the directory to store Log files
            Utility.createDirectory();
            if (!Directory.Exists(Constants.globalRecentResultsPath))
            {
                Directory.CreateDirectory(Constants.globalRecentResultsPath);

            }

            if (!File.Exists(shortReportPath))
            {
                var resultfile = File.Create(shortReportPath);
                resultfile.Close();
                goto write;
            }
            else
            {
                goto write;
            }
            write:
            File.AppendAllText(@shortReportPath, htmlText);
            Constants.GlobalLogOverviewPath = shortReportPath;
            File.Copy(shortReportPath, RecentshortReportPath, true);
        }
        public static void excelAddButtonWithVBA()
        {
            //Excel.Application xlApp = new Excel.Application();
            //Excel.Workbook xlBook = xlApp.Workbooks.Open(Constants.mREPORTPATH);
            //Excel.Worksheet wrkSheet = xlBook.Worksheets[1];
            //Excel.Range range;

            // Playback.Initialize();
            xlApp = new ExcelQ.Application();
            //Excel.Workbook wb = xla.Workbooks.Add(Excel.XlSheetType.xlWorksheet);

            Workbook xlBook = xlApp.Workbooks.Open(Constants.mREPORTPATH, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Worksheet wrkSheet = (Worksheet)xlApp.ActiveSheet;
            Range range;
            try
            {
                //set range for insert cell
                range = wrkSheet.get_Range("A1:A1");
                //insert the dropdown into the cell
                Buttons xlButtons = wrkSheet.Buttons();
                //Excel.Button xlButton = xlButtons.Add((double)range.Left, (double)range.Top, (double)range.Width, (double)range.Height);
                ExcelQ.Button xlButton = xlButtons.Add(1150, 0, 102, 32);
                xlButton.Font.Color = Color.DarkBlue;
                xlButton.Font.Size = 16;
                xlButton.Font.Bold = true;
                xlButton.Font.Name = "Arial";
                //xlButton.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);            
                //set the name of the new button
                xlButton.Name = "btnDoSomething";
                xlButton.Text = "Log";
                xlButton.OnAction = "btnDoSomething_Click";
                buttonMacro(xlButton.Name, xlApp, xlBook, wrkSheet);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
            //xlApp.Visible = true;
            xlBook.Save();
            xlBook.Close(true, Type.Missing, Type.Missing);
            xlApp.Quit();
            releaseObject(wrkSheet);
            releaseObject(xlBook);
            releaseObject(xlApp);
        }
        private static void buttonMacro(string buttonName, ExcelQ.Application xlApp, ExcelQ.Workbook wrkBook, ExcelQ.Worksheet wrkSheet)
        {
            StringBuilder sb;
            VBIDE.VBComponent xlModule;
            VBIDE.VBProject prj;
            string path = @"C:\Windows\explorer.exe";

            prj = wrkBook.VBProject;
            sb = new StringBuilder();
            // build string with module code
            sb.Append("Sub " + buttonName + "_Click()" + "\n");
            sb.Append("\t" + " Dim x As Variant" + "\n");
            sb.Append("\t" + " Dim Path As String" + "\n");
            sb.Append("\t" + " Dim File As String" + "\n");
            sb.Append("\t" + " Path =" + '"' + path + '"' + "\n");
            sb.Append("\t" + " File=" + '"' + Constants.globalLogPath + '"' + "\n");
            sb.Append("\t" + "x = Shell(Path+" + '"' + ' ' + '"' + "+File" + ',' + "vbNormalFocus" + ")" + "\n");
            //sb.Append("\t" + "msgbox \"" + buttonName + "\"\n"); // add your custom vba code here
            sb.Append("End Sub");

            // set an object for the new module to create
            xlModule = wrkBook.VBProject.VBComponents.Add(VBIDE.vbext_ComponentType.vbext_ct_StdModule);

            // add the macro to the spreadsheet
            xlModule.CodeModule.AddFromString(sb.ToString());
        }
        public static bool DismissDialog(BrowserWindow browser, int waitMillis = 5 * 1000)
        {

            WinControl dialog = new WinControl();
            dialog.SearchProperties[UITestControl.PropertyNames.ControlType] = "Dialog";
            dialog.SearchProperties[UITestControl.PropertyNames.ClassName] = "#32770";
            //dialog.WindowTitles.Add("Title of the dialog");

            if (dialog.WaitForControlExist(waitMillis))
            {
                Playback.Wait(5 * 1000);
                browser.PerformDialogAction(Microsoft.VisualStudio.TestTools.UITest.Extension.BrowserDialogAction.Ok);
                return true;
            }
            return false;
        }
        //public static string toHTML_Table(System.Data.DataTable dt)
        //{
        //    if (dt.Rows.Count == 0) return ""; // enter code here

        //    StringBuilder builder = new StringBuilder();
        //    builder.Append("<!DOCTYPE html>");
        //    builder.Append("<html>");
        //    builder.Append("<head>");
        //    builder.Append("<title>");
        //    builder.Append("Page-");
        //    builder.Append(Guid.NewGuid());
        //    builder.Append("</title>");
        //    builder.Append("</head>");
        //    builder.Append("<body>");
        //    // builder.Append("<table cellpadding='5' cellspacing='0' style='border: solid 1px Silver; font-size:small;table-layout:fixed;font-family:arial,sans-serif;'> ");
        //    builder.Append("<table cellspacing='0' cellpadding='4' border= 'thin'  width='100' style =border-top-width:thin;font-size:small;table-layout:fixed;'>");
        //    builder.Append("<col width='250'>");
        //    builder.Append("<col width='400'>");
        //    builder.Append("<col width='100'>");
        //    builder.Append("<col width='100'>");
        //    builder.Append("<col width='100'>");
        //    builder.Append("<col width='100'>");
        //    builder.Append("<tr>");

        //    if (count == 0)
        //    {
        //        foreach (DataColumn c in dt.Columns)
        //        {
        //            builder.Append("<td><b>");
        //            builder.Append(c.ColumnName);
        //            builder.Append("</b></td>");
        //        }
        //        count = 1;
        //        goto tag;
        //    }
        //    else
        //    {
        //        goto tag;
        //    }
        //    tag:
        //    builder.Append("</tr>");
        //    foreach (DataRow r in dt.Rows)
        //    {
        //        builder.Append("<tr>");
        //        foreach (DataColumn c in dt.Columns)
        //        {
        //            //Below If condition is used to check whether the Status column value is Failed If so then Red colour will be given to that particular cell
        //            if (r[c.ColumnName].ToString() == "Failed")
        //            {
        //                builder.Append("<td>");
        //                builder.Append("<font color='red'>");
        //                builder.Append(r[c.ColumnName]);
        //                builder.Append("</font>");
        //                goto build;
        //            }
        //            else if (r[c.ColumnName].ToString() == "Inconclusive")
        //            {

        //                builder.Append("<td>");
        //                builder.Append("<font color='663300'>");
        //                builder.Append(r[c.ColumnName]);

        //                builder.Append("</font>");
        //                goto build;

        //            }
        //            else if (r[c.ColumnName].ToString() == "Passed")
        //            {

        //                builder.Append("<td>");
        //                builder.Append("<font color='00FF33'>");
        //                builder.Append(r[c.ColumnName]);

        //                builder.Append("</font>");
        //                goto build;

        //            }
        //            else
        //            {
        //                builder.Append("<td>");
        //                goto Passed;
        //            }

        //            Passed:
        //            builder.Append(r[c.ColumnName]);
        //            build:
        //            builder.Append("</td>");
        //        }
        //        builder.Append("</tr>");
        //    }
        //    builder.Append("</table>");
        //    builder.Append("</body>");
        //    builder.Append("</html>");
        //    return builder.ToString();

        //}
        public static string toHTML_Table(System.Data.DataTable dt)
        {
            if (dt.Rows.Count == 0) return ""; // enter code here

            StringBuilder builder = new StringBuilder();
            builder.Append("<!DOCTYPE html>");
            builder.Append("<html>");
            builder.Append("<head>");
            builder.Append("<title>");
            builder.Append("Page-");
            builder.Append(Guid.NewGuid());
            builder.Append("</title>");
            builder.Append("</head>");
            builder.Append("<body>");
            builder.Append("<img src = 'C:\\AutomationTesting\\AutomationTestScripts\\NGW_SharePoint\\NGW_SharePoint\\BoschLogo.png'/>");
            //builder.Append("<img src='data: image/png; base64,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' alt='BoschLogo.png'/>");
            //builder.Append("<div style='border - top:3px solid #22BCE5'>&nbsp;</div>");
            //builder.Append("<span style = 'font-family:Arial;font-size:10pt'>");
            builder.Append("<style>");
            builder.Append("hr {");
            builder.Append("display: block;");
            builder.Append("margin - top: 0.5em;");
            builder.Append("margin - bottom: 0.5em;");
            builder.Append("margin - left: auto;");
            builder.Append("margin - right: auto;");
            builder.Append("border - style: inset;");
            builder.Append("border - width: 1px;");
            builder.Append("}");
            builder.Append("</style>");
            builder.Append("</br>");
            builder.Append("Automation_TestRun_Results <br/><br/>");
            builder.Append("Environment: <b>{Environment}</b>,<br/><br/>");
            builder.Append("Run_Date: <b>{RunDate}</b>,<br/><br/>");
            builder.Append("</span>");
            builder.Append("<br/>");

            builder.Append("<table border = '1' id = 't01'>");
            builder.Append("<tr>");
            builder.Append("<th> Total Run </th>");
            builder.Append("<th> Total Pass </th>");
            builder.Append("<th> Total Fail </th>");
            builder.Append("</tr>");
            builder.Append("<tr>");
            builder.Append("<td>{TotalRun}</td>");
            builder.Append("<td>{TotalPass}</td>");
            builder.Append("<td>{TotalFail}</td>");
            builder.Append("</tr>");
            builder.Append("</table >");
            builder.Append("<br/>");
            // builder.Append("<table cellpadding='5' cellspacing='0' style='border: solid 1px Silver; font-size:small;table-layout:fixed;font-family:arial,sans-serif;'> ");
            builder.Append("<table style = 'width:100%' border = '1'>");                      
           // builder.Append("<tr>");
            builder.Append("<th>" + "ModuleName" + "</th>");
            builder.Append("<th>" + "TestCaseName" + "</th>");
            builder.Append("<th>" + "Iteration" + "</th>");
            builder.Append("<th>" + "No Of Steps From Script" + "</th>");
            builder.Append("<th>" + "Status" + "</th>");
            builder.Append("<th>" + "TestUser" + "</th>");

            foreach (DataRow r in dt.Rows)
            {
                builder.Append("<tr>");

                foreach (DataColumn c in dt.Columns)
                {
                    //Below If condition is used to check whether the Status column value is Failed If so then Red colour will be given to that particular cell
                    if (r[c.ColumnName].ToString() == "Failed")
                    {
                        builder.Append("<td style='word -break:break-all'>");
                        builder.Append("<font color='#660000'>");
                        builder.Append(r[c.ColumnName]);
                        builder.Append("</font>");
                        goto build;
                    }
                    else if (r[c.ColumnName].ToString() == "Inconclusive")
                    {

                        builder.Append("<td style='word -break:break-all'>");
                        builder.Append("<font color='663300'>");
                        builder.Append(r[c.ColumnName]);

                        builder.Append("</font>");
                        goto build;

                    }
                    else if (r[c.ColumnName].ToString() == "Passed")
                    {

                        builder.Append("<td style='word -break:break-all'>");
                        builder.Append("<font color='#006600'>");
                        builder.Append(r[c.ColumnName]);

                        builder.Append("</font>");
                        goto build;

                    }
                    else
                    {
                        builder.Append("<td style='word -break:break-all'>");
                        goto Passed;
                    }

                    Passed:
                    builder.Append(r[c.ColumnName]);
                    build:
                    builder.Append("</td>");
                }
                builder.Append("</tr>");
            }
            builder.Append("</table>");
            builder.Append("</body>");
            builder.Append("</html>");
            return builder.ToString();

        }

        public static void DirectoryCopyToNewDirectory(string sourceDirName, string destDirName, bool copySubDirs)
        {
            // Get the subdirectories for the specified directory.
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);
            DirectoryInfo[] dirs = dir.GetDirectories();

            if (!dir.Exists)
            {
                throw new DirectoryNotFoundException(
                    "Source directory does not exist or could not be found: "
                    + sourceDirName);
            }

            // If the destination directory doesn't exist, create it. 
            if (Directory.Exists(destDirName))
            {


                System.IO.DirectoryInfo di = new DirectoryInfo(destDirName);

                foreach (FileInfo file in di.GetFiles())
                {
                    file.Delete();
                }
                foreach (DirectoryInfo directory in di.GetDirectories())
                {
                    directory.Delete(true);
                }


                Directory.CreateDirectory(destDirName);
            }
            else
            {
                Directory.CreateDirectory(destDirName);
            }

            // Get the files in the directory and copy them to the new location.
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                string temppath = Path.Combine(destDirName, file.Name);
                file.CopyTo(temppath, true);
            }

            // If copying subdirectories, copy them and their contents to new location. 
            if (copySubDirs)
            {
                foreach (DirectoryInfo subdir in dirs)
                {
                    string temppath = Path.Combine(destDirName, subdir.Name);
                    DirectoryCopyToNewDirectory(subdir.FullName, temppath, copySubDirs);
                }
            }
        }

        public static void ConvertExceltoHtml()
        {
            try
            {
                SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

                ExcelFile ef = ExcelFile.Load(Constants.mREPORTPATH);

                var ws = ef.Worksheets["MyData"];

                int n = ws.Charts.Count;
                int m = ws.CalculateMaxUsedColumns();
                // Print area can be used to specify custom cell range which should be exported to HTML.
                ws.NamedRanges.SetPrintArea(ws.Cells.GetSubrange("L3", "M7"));

                HtmlSaveOptions options = new HtmlSaveOptions()
                {
                    HtmlType = HtmlType.Html,
                    SelectionType = SelectionType.ActiveSheet

                };
                options.EmbedImages = true;

                // ImageSaveOptions img = new ImageSaveOptions();

                ef.Save(Constants.globalResultsPath + "\\AutomationTestStatus.html", options);
                //  ef.Save(Constants.globalResultsPath + "\\LogStatus.png", img);
            }
            catch (Exception e)
            {

            }
        }

        public static void ConvertExceltoImage()
        {
            try
            {
                //Initialize a new Workbook object
                Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();

                //Open Template Excel file
                workbook.LoadFromFile(Constants.mREPORTPATH);
                 //Get the first wirksheet in Excel file
                Spire.Xls.Worksheet sheet = workbook.Worksheets[0];

                //Specify Cell Ranges and Save to certain Image formats
                sheet.SaveToImage(3, 12, 7, 13).Save(Constants.globalResultsPath + "\\TestStatus.png", ImageFormat.Png);

            }
            catch (Exception e)
            {

            }
        }
    }

}





