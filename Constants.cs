using System;
using System.IO;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using System.Data;
using System.Collections.Generic;

namespace NGW_SharePoint.Utility
{
    class Constants
    {
        #region Variable Declarations

        // The log text is used to store the 'Step' value of report
        public static string logText;
       
        // The report flag is a bool value that is used to check if the header is available in the report      
        public static bool reportFlag=false;

        // The null string is to assign 'null' value
        public static string nullString = null;
       
        // The get current method is used to store the 'TestMethod'     
        public static string getCurrentMethod;
     
        // The get calling method is used to store the 'CodedUiMethod'    
        public static string getCallingMethod;

        // The project name is to store the current project name      
        public static readonly string currentProjectName = Assembly.GetCallingAssembly().GetName().Name;

        // The global path is used to store the path of Result folder
        public static readonly  string globalResultsPath= string.Concat (@"C:\", currentProjectName + '\\' + DateTime.Today.ToString("MM_dd_yyyy"), "\\", "Result" + DateTime.Now.ToString("MM_dd_yyyy_HH_mm_ss"));

        //global log path is used to store the path of log file created
        public static  string globalLogPath = null;

        // The m report path is used to store the File Name
        public static  string mREPORTPATH = Path.Combine(globalResultsPath, "Automation_Test_Report.xlsm");

        // The global path is used to store the path of Result folder
        public static readonly string globalRecentResultsPath = string.Concat(@"C:\", currentProjectName + '\\' + "TestResults");

        //The recent report path is used to store the last execution report
        public static string RecentREPORTPATH = Path.Combine(globalRecentResultsPath, "Automation_Test_Report.xlsm");

        // The global table is to store values of Data Table
        public static DataTable globalTable =new DataTable();

        //Stores the steps in each and every testcase at run time
        public static int stepsCounter = 0;

        // The failcount
        public static int failCount=0;

        // The pass count is to store number of passed test cases
        public static int passCount = 0;

        // The otherscount
        public static int othersCount = 0;

        // The totalrun
        public static int totalRun = 0;

        //used to store errors count
        public static int errorCount = 0;
     
       //public static string dmsModuleName = "DMS";
        public static string ModuleName = null;

        public static string htmContent = null;

        public static string managementReportName = currentProjectName + "Management Report";

        internal static string GlobalLogOverviewPath;
       
        //internal static string reportTextInHtml;
        public static List<Tuple<String, String, String, String, String>> logTuple = new List<Tuple<String, String, String, String, String>>();

        public static  List<Tuple<String, String, String, int, String, String>> resultsOverviewTuple = new List<Tuple<String, String, String, int, String, String>>();

        internal static string ExceptionPath;

        internal static string Environment=string.Empty;


        /// <summary>
        /// SFB
        /// </summary>
        
        public static string ExePath = "C:\\Program Files (x86)\\Microsoft Office\\Office15\\lync.exe";

        public static string SFBName = "Skype for Business ";

        public static string SigninAddress = "Sign-in address:";

        public static string ProjectName = "SfbClientTesting";

        public static string ProjectDataFilePath = "SfbClientTesting\\DataFiles\\SampleTest\\";

        public static string DataFilePath = @"C:\SFB_Automation\DataFiles\SampleTest\";

        public static string Start = "Start ";

        public static string Start_menu = "Start menu";

        public static string Program_Manager = "Program Manager";

        public static string Conversation = "Conversation (3 Participants)";

        public static string Conversation2 = "Conversation(2 Participants)";

        public static string CreatePoll = "Create a Poll";

        public static string Outlook = "Untitled - Message (HTML) ";

        public static string Meeting = "Untitled - Meeting";

        public static string MicrosoftOutlook = "Microsoft Outlook";

        public static string ReportType = "SFB";
        public enum status
        {
            Available,
            Sign_Out,
            Off_Work,
            Appear_Away,
            Be_Right_Back,
            Do_Not_Disturb
            
        }

        #endregion

        #region Function

        /// <summary>
        /// Test startup.
        /// </summary>
        public static void StartTest()
        {
        // Configure the playback engine
        Playback.PlaybackSettings.WaitForReadyLevel = WaitForReadyLevel.Disabled;
        Playback.PlaybackSettings.MaximumRetryCount = 10;
        Playback.PlaybackSettings.ShouldSearchFailFast = false;
        Playback.PlaybackSettings.DelayBetweenActions = 500;
        Playback.PlaybackSettings.SearchTimeout = 1000;

        // Add the error handler
        Playback.PlaybackError -= Playback_PlaybackError; 
        Playback.PlaybackError += Playback_PlaybackError;
        }
        /// <summary>
        /// PlaybackError event handler.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="PlaybackErrorEventArgs"/> instance containing the event data.</param>
        private static void Playback_PlaybackError(object sender, PlaybackErrorEventArgs e)
        {
             // Wait a second
             System.Threading.Thread.Sleep(1000);

            // Retry the failed test operation
            e.Result = PlaybackErrorOptions.Retry;
        }
        #endregion
    }

}
