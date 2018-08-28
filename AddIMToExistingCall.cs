using System;
using System.Drawing;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SFB.LibraryFunctions;
using NGW_SharePoint.ObjectRepository.SampleTest;
using NGW_SharePoint.Utility;
using System.IO;
using System.Threading;
using System.Diagnostics;
using NGW_SharePoint.ObjectRepository.SampleTest2016;
using System.Collections.Generic;

namespace SFBTesting.DriverFunctions.SFBTesting
{
    /// <summary>
    /// Summary description for SkypeSigning
    /// </summary>
    [CodedUITest]
    public class SfB_Audio_AddIMtoExistingCall_1 : SampleTestObjectRepository
    {
        public SfB_Audio_AddIMtoExistingCall_1()
        {
            Constants.Environment = objEnvironment;
        }
    private static int iterationintialize = 1;
         private static string DataPath = Constants.DataFilePath;
        private  long data_rows = globalFunctions.CountLinesInFile(DataPath + "AddIMtoExistingCall.csv");
        private  static string sr = globalFunctions.GetLine(DataPath + "AddIMtoExistingCall.csv", iterationintialize);
        private  string[] values = sr.Split(',');

        private  long data_rows1 = globalFunctions.CountLinesInFile(DataPath + "Skypeversion.csv");
        private  static string sr1 = globalFunctions.GetLine(DataPath + "Skypeversion.csv", iterationintialize);
        private  string[] values1 = sr1.Split(',');
        private  string filePath = "C:\\SFB_Automation\\";

        [TestInitialize]
        public void MyTestInitialize()
        {
            try
            {
                //Minimize remote window
                //globalFunctions.SearchOnDesktopDoubleClick(filePath + "Images\\Remote_Minimize.png", 0.93f, 2, 1, "Remote Desktop Minimize");

                //Skype
                Process[] processes = System.Diagnostics.Process.GetProcessesByName("lync");
                if (processes.Length == 0)
                {
                    globalFunctions.StartLyncProcess(values1[1].ToString(), values1[0], values1[2], values1[3], values1[4], 1);
                }
                else
                {
                    WinWindow SkypeForBusiness = globalFunctions.GetWindowByName(objbuttonname, 1);
                    if (globalFunctions.WinTextNotExist(SkypeForBusiness, values1[1].ToString(), 1))
                    {
                        //Kill
                        globalFunctions.StopProcess("lync");
                        globalFunctions.StartLyncProcess(values1[1].ToString(), values1[0], values1[2], values1[3], values1[4], 1);
                    }
                }
                processes = null;
                //Remote login
                globalFunctions.RemoteConnection(values1[6].ToString(), values1[7].ToString(), values1[8].ToString(), values1[9].ToString(), values1[0].ToString(), values1[5].ToString(), 1);

                bool shortcutexist = false;
                //Start Skype- by clicking on the Skype shortcut                      
                bool shortcutexist1 = globalFunctions.SearchInCitrixWindowAndDoubleClick(values1[5].ToString(), filePath + "Images\\SkypeForBusinessShortcut.png", 0.90f, 0, 1, "SkypeForBusiness Shortcut");
                if (!shortcutexist)
                {
                    shortcutexist = globalFunctions.SearchInCitrixWindowAndDoubleClick(values1[5].ToString(), filePath + "Images\\SkypeForBusinessShortcut1.png", 0.80f, 1, 1, "Alternative SkypeForBusiness Shortcut");
                }

                //Minimize remote window
                System.Windows.Forms.SendKeys.SendWait("^%{BREAK}");
                globalFunctions.ShowDesktop();
                //globalFunctions.SearchOnDesktopDoubleClick(filePath + "Images\\Remote_Minimize.png", 0.93f, 1, 1, "Remote Desktop Minimize");

            }
            catch (Exception exception)
            {
                string testName = TestContext.TestName.ToString();
                string status = "Test Initialization Failed";

                LogDetails.Log(globalFunctions.listOfTuples, _dataRow);

                //Checking the current Test Status
                if (status == "Test Initialization Failed")
                {
                    Constants.failCount = Constants.failCount + 1;
                    _captureFailedScreenPath = Utility.FailureScreenShotCapture();
                    Utility.ListToDataTable(testName, status, _captureFailedScreenPath, _dataRow);
                }
                Constants.totalRun = Constants.passCount + Constants.failCount + Constants.othersCount + Constants.errorCount;
                Reports.ExportDataTable2Excel(Constants.managementReportName, Color.White, Color.Maroon, 15, true, "01/06/2015", Color.White,
                Color.Black, 10, Constants.globalTable, Color.Orange, Color.White, "MyData",
                Constants.mREPORTPATH);

                Utility.ListAddition(testName, status, _dataRow, Constants.stepsCounter);
                Utility.DirectoryCopyToNewDirectory(Constants.globalResultsPath, Constants.globalRecentResultsPath, true);


                Reports.ExceptionReports(TestContext.TestName, exception.Message, exception.Source, exception.StackTrace);
                Assert.Fail("An" + exception.GetType() + "Occured");
                throw;
            }
        }
        [TestCategory("Audio Test"), TestMethod(), Priority(1)]
        public void AddIMtoAnExistingCall_1()
        {
            //string testResultsOutFolerPath = System.IO.Directory.GetCurrentDirectory();
            //int posA = testResultsOutFolerPath.IndexOf("TestResults");
            //string filePath = testResultsOutFolerPath.Substring(0, posA);
            int catchiteration = 0;
            try
            {
                for (int iteration = 1; iteration <= data_rows; iteration++)
                {
                    catchiteration = iteration;
                    string sr = globalFunctions.GetLine(DataPath + "AddIMtoExistingCall.csv", iteration);
                    string[] values = sr.Split(',');
                    //Launch Skype
                    WinWindow SkypeForBusiness = globalFunctions.GetWindowByName(Constants.SFBName, iteration);

                    SkypeForBusiness.SetFocus();

                    // click on contacts
                    WinListItem winListItem = globalFunctions.WinListItemFindByName(SkypeForBusiness, objContactsTab, iteration);

                    //Search by name
                    globalFunctions.WinEditByName(SkypeForBusiness, objFindSomeone, values[0].ToString(), iteration, " Search by Name");

                    //Right click on searched name 
                    globalFunctions.WinListItemRightClickByName(SkypeForBusiness, values[0].ToString(), iteration);

                    //Click on Skype for Business Call
                    if (objSkypeversion == values1[0])
                    {
                        //Click on Skype for Business Call
                        globalFunctions.WinMenuItemClickByClassNameAndSubMenuItemClick(SkypeForBusiness, objStatusMenuItem, objMenuCall, objSubMenuSkyCall, iteration);
                    }
                    else
                    {
                        SampleTestObjectRepository2016 sample = new SampleTestObjectRepository2016();
                        //Click on Skype for Business Call
                        globalFunctions.WinMenuItemClickByClassNameAndSubMenuItemClick(SkypeForBusiness, objStatusMenuItem, objMenuCall, sample.objSubMenuSkypeCall, iteration);

                    }
                    //Minimize
                  //  globalFunctions.WinButtonClickByParentClassNameAndButtonName(SkypeForBusiness, objSeeContactCardMenuItem, objMinimize, iteration);

                   
                    //To Open and maximize remote window
                    //System.Windows.Forms.SendKeys.SendWait("%{TAB}");
                    ////   System.Windows.Forms.SendKeys.SendWait("^%{BREAK}");
                    //WinWindow maximizeWindow =  globalFunctions.GetWindowByNameAndClassName(values1[5].ToString(), objRemWindowMaxClassNAme, iteration);
                    //globalFunctions.ClickOnWindowCloseButton(maximizeWindow, objMaximize, iteration);
                   
                    //citrix window
                    #region Click on RDC Icon on toolbar by Control properties
                    WinWindow citrixwindow = globalFunctions.GetWindowByNameAndClassName(objtoolbarname, objwindowclassname, iteration);
                    if (objSkypeversion == values1[0])
                    {
                        // click on the icon 
                        globalFunctions.ClickOnToolBarButton(citrixwindow, objtoolbarname, values1[5].ToString(), iteration);
                    }
                    else
                    {
                        SampleTestObjectRepository2016 sample = new SampleTestObjectRepository2016();
                        globalFunctions.ClickOnToolBarButton(citrixwindow, objtoolbarname, sample.objRemoteDesktopConnection, iteration);
                    }
                    WinWindow maximizeWindow = globalFunctions.GetWindowByNameAndClassName(values1[5].ToString(), objRemWindowMaxClassNAme, iteration);
                    globalFunctions.ClickOnWindowButtonIfExist(maximizeWindow, objMaximize, iteration);

                    #endregion

                    // Open the Citrix Window
                   Thread.Sleep(2000);
                           bool callexist = globalFunctions.SearchBottomLeftInCitrixWindow(values1[5].ToString(),filePath + "Images\\call.png", 0.97f, 0, iteration, "Call");
                           if (!callexist)
                           {
                               globalFunctions.SearchBottomLeftInCitrixWindow(values1[5].ToString(),filePath + "Images\\AlternativeCall.png", 0.97f, 1, iteration, "AlternativeCall");
                           }
                           Thread.Sleep(2000);

                    //Minimize remote window
                    System.Windows.Forms.SendKeys.SendWait("^%{BREAK}");

                    globalFunctions.ShowDesktop();
                    //globalFunctions.SearchOnDesktopDoubleClick(filePath + "Images\\Remote_Minimize.png", 0.93f, 1, 1, "Remote Desktop Minimize");

                    //Check microphone button
                    WinWindow testwindow = globalFunctions.GetWindowByNameAndClassName(values[0].ToString(), LyncConversationWindowClassName, iteration, "Conversation Window");

                    globalFunctions.WinClientAndClickAlert(testwindow, objunmuted, objAlert, iteration, "Your microphone is unmuted. Press space to mute. Button");
                           
                    //IM Image
                    if (objSkypeversion == values1[0])
                    {
                        //IM Image
                        globalFunctions.WinControlClickByParentButtonNameAndWindowName(values[0].ToString(), LyncConversationWindowClassName, objIMButton, iteration);

                        System.Threading.Thread.Sleep(2000);

                    }
                    else
                    {
                        SampleTestObjectRepository2016 sample = new SampleTestObjectRepository2016();
                        //IM Image
                        globalFunctions.WinControlClickByParentButtonNameAndWindowName(values[0].ToString(), sample.LyncConversationWindowClassName, sample.objIMButton, iteration);

                        System.Threading.Thread.Sleep(2000);
                   }    
                    // enter data
                    Keyboard.SendKeys(values[1].ToString() + "{Enter}");
                    System.Threading.Thread.Sleep(1500);

                    //verify
                    //maximize the IM window
                    WinWindow maximizewindow = globalFunctions.GetWindowByNameAndClassNameAndTitle(values[0].ToString(), LyncConversationWindowClassName, values[0].ToString(), iteration, "Conversation Window");
                    globalFunctions.GetWindowClassNameDialogClickButton(maximizewindow, objSeeContactCardMenuItem, "Maximize", iteration);


                    System.Collections.Generic.List<string> capturedpath = globalFunctions.GetWindowByNameAndClassNameAndAllCaptureIM(values[0].ToString(), LyncConversationWindowClassName, iteration, "capture text area");

                    string text = globalFunctions.OCRPictrueInDisk(capturedpath[0].ToString(), iteration);

                    foreach (System.Diagnostics.Process proc in System.Diagnostics.Process.GetProcessesByName("ONENOTE"))
                    {
                        proc.Close();
                        proc.Refresh();
                    }

                    if (text.Contains(values[1].ToString()))
                    {
                        globalFunctions.Result("IM message is sent, message details: " + text, "pass", iteration);
                    }
                    else if (text == "Error")
                    {
                        globalFunctions.Result("Error in Intializing One Note", "fail", iteration);
                    }
                    else
                    {
                        string text1 = globalFunctions.OCRPictrueInDisk(capturedpath[1].ToString(), iteration);

                        foreach (System.Diagnostics.Process proc in System.Diagnostics.Process.GetProcessesByName("ONENOTE"))
                        {
                            proc.Close();
                        }
                        if (text1.Contains(values[1].ToString()))
                        {
                            globalFunctions.Result("IM message is sent, message details: " + text1, "pass", iteration);
                        }
                        else
                        {
                            globalFunctions.Result("IM message is not sent", "fail", iteration);
                        }
                    }

                    //End Call if exist
                    globalFunctions.WinButtonHoverByNameAndClickImageIfExist(maximizewindow, objstopaudio, values[0].ToString(), iteration);

                    //Fetching the system time and placing in data file
                    globalFunctions.InsertDateTimeToDataFile(DataPath + "AddIMtoExistingCall.csv", 3, iteration - 1);

                    ////close window if exist
                    if (globalFunctions.GetWindowByNameAndClassNameIfWindowExist(values[0].ToString(), LyncConversationWindowClassName, iteration, "Conversation Window"))
                    {
                        WinWindow closewindow1 = globalFunctions.GetWindowByNameAndClassName(values[0].ToString(), LyncConversationWindowClassName, iteration, "Conversation Window");
                        WinWindow closewindow = globalFunctions.GetWindowByParentClassName(closewindow1, values[0].ToString(), objSeeContactCardMenuItem, iteration, "Skype for Business");

                        globalFunctions.WinClickByControlTypeIfExist(closewindow, objDialog, objCloseButton, values[0].ToString(), iteration);
                        globalFunctions.GetWindowByAccessibleNameAndClassNameIfExist(objDoYouWantToClose, objSeeContactCardMenuItem, objDoYouWantToClose, objCloseAllTabs, iteration);

                    }

                    //Click on OK if exist
                    globalFunctions.GetWindowByNameAndClassNameIfExist(objbuttonname, objSkypeWindowClassName, objDialogName, objOKButton, iteration);

                    //Minimize
                    globalFunctions.WinButtonClickByParentClassNameAndButtonName(SkypeForBusiness, objSeeContactCardMenuItem, objMinimize, 1);

                    //To Open and maximize remote window
                    System.Windows.Forms.SendKeys.SendWait("%{TAB}");
                    globalFunctions.ClickOnWindowCloseButton(maximizeWindow, "Maximize", iteration);


                    // Open the Citrix Window
                    #region
                    //if (objSkypeversion == values1[0])
                    //{
                    //    // click on the icon 
                    //    globalFunctions.ClickOnToolBarButton(citrixwindow, objtoolbarname, values1[5].ToString(), iteration);
                    //}
                    //else
                    //{
                    //    SampleTestObjectRepository2016 sample = new SampleTestObjectRepository2016();
                    //    globalFunctions.ClickOnToolBarButton(citrixwindow, objtoolbarname, sample.objRemoteDesktopConnection, iteration);
                    //}
                    #endregion

                    // Minimize Skype
                    //bool exist1 = false;
                    //       bool exist = globalFunctions.SearchInCitrixWindow(values1[5].ToString(),filePath + "Images\\SkypeForBusiness.png", 0.97f, 0, iteration, "SkypeForBusiness");
                    //       if (!exist)
                    //       {
                    //           exist1 = globalFunctions.SearchInCitrixWindow(values1[5].ToString(),filePath + "Images\\AlternativeSkypeForBusiness.png", 0.97f, 2, iteration, "AlternativeSkypeForBusiness");
                    //       }
                    //       if (exist == true || exist1 == true)
                    //       {
                    //           bool exist3 = globalFunctions.SearchInCitrixWindow(values1[5].ToString(),filePath + "Images\\Minimize.png", 0.92f, 0, iteration, "Minimize");
                    //           if (!exist3)
                    //           {
                    //               bool exist4 = globalFunctions.SearchInCitrixWindow(values1[5].ToString(),filePath + "Images\\AlternativeMinimize.png", 0.92f, 2, iteration, "AlternativeMinimize");
                    //               if (!exist4)
                    //               {
                    //                   bool exist5 = globalFunctions.SearchInCitrixWindow(values1[5].ToString(),filePath + "Images\\Alternative1Minimize.png", 0.92f, 2, iteration, "Aletrnative1Minimize");
                    //                   if (!exist5)
                    //                   {
                    //                       globalFunctions.SearchInCitrixWindow(values1[5].ToString(),filePath + "Images\\AlternativeMinimize2.png", 0.92f, 2, iteration, "AletrnativeMinimize2");
                    //                   }
                    //               }
                    //           }
                    //       }

                    //       //Click on Close in Citrix window
                    //       Thread.Sleep(1000);
                    //       bool clicksuccess = globalFunctions.SearchInCitrixWindow(values1[5].ToString(),filePath + "Images\\Close.png", 0.93f, 0, iteration, "Close");
                    //       if (!clicksuccess)
                    //       {
                    //           globalFunctions.SearchInCitrixWindow(values1[5].ToString(),filePath + "Images\\AlternativeClose.png", 0.93f, 1, iteration, "AlternativeClose");
                    //       }
                    //       Thread.Sleep(2000);

                    //globalFunctions.SearchInCitrixWindow(values1[5].ToString(), filePath + "Images\\OK.png", 0.93f, 2, iteration, "OK");

                    //// Check for Close all tabs
                    //globalFunctions.SearchInCitrixWindow(values1[5].ToString(),filePath + "Images\\CloseAllTabs.png", 0.97f, 2, iteration, "Close All Tabs Button ");

                    //powershell details set up, two list order should match 
                    List<string> wordToBeReplaced = new List<string>();
                    wordToBeReplaced.Clear();
                    wordToBeReplaced.Add("$server");
                    wordToBeReplaced.Add("$Username");
                    wordToBeReplaced.Add("$Password");


                    List<string> newWords = new List<string>();
                    newWords.Clear();
                    newWords.Add(values1[9].ToString());
                    newWords.Add(values1[7].ToString());
                    newWords.Add(values1[8].ToString());

                    globalFunctions.ReplaceTheFirstOccurances(DataPath + "CleanupLync.ps1", wordToBeReplaced, newWords, 1);

                    //Kill remote skype and start process only in task manager
                    globalFunctions.RunPowershell(DataPath + "CleanupLync.ps1", 1);

                    //Minimize remote window
                    System.Windows.Forms.SendKeys.SendWait("^%{BREAK}");
                    globalFunctions.ShowDesktop();
                    //globalFunctions.SearchOnDesktopDoubleClick(filePath + "Images\\Remote_Minimize.png", 0.93f, 2, 1, "Remote Desktop Minimize");


                }
            }
           catch (Exception exception)
            {
                _captureFailedScreenPath = Utility.FailureScreenShotCapture();

                globalFunctions.ReturningToLocalDesktop(exception.Message, iterationintialize);
                globalFunctions.StopProcess("lync");

                //powershell details set up, two list order should match 
                List<string> wordToBeReplaced = new List<string>();
                wordToBeReplaced.Clear();
                wordToBeReplaced.Add("$server");
                wordToBeReplaced.Add("$Username");
                wordToBeReplaced.Add("$Password");


                List<string> newWords = new List<string>();
                newWords.Clear();
                newWords.Add(values1[9].ToString());
                newWords.Add(values1[7].ToString());
                newWords.Add(values1[8].ToString());

                globalFunctions.ReplaceTheFirstOccurances(DataPath + "CleanupLync.ps1", wordToBeReplaced, newWords, 1);

                //Kill remote skype and start process only in task manager
                globalFunctions.RunPowershell(DataPath + "CleanupLync.ps1", 1);

                Reports.ExceptionReports(TestContext.TestName, exception.Message, exception.Source, exception.StackTrace);
                Assert.Fail("An" + exception.GetType() + "Occured");
                throw;
               
            }
            finally
            {
                LogDetails.Log(globalFunctions.listOfTuples, _dataRow);
            }
        }


        [TestCleanup()]
        public void MyTestCleanUp()
        {
            string testName = TestContext.TestName.ToString();
            string status = TestContext.CurrentTestOutcome.ToString();
            try
            {
                //Checking the current Test Status
                if (TestContext.CurrentTestOutcome == UnitTestOutcome.Failed)
                {
                    Constants.failCount = Constants.failCount + 1;
                  //  _captureFailedScreenPath = Utility.FailureScreenShotCapture();
                    Utility.ListToDataTable(testName, status, _captureFailedScreenPath, _dataRow);
                }
                else if (TestContext.CurrentTestOutcome == UnitTestOutcome.Passed)
                {
                    Constants.passCount = Constants.passCount + 1;
                    Utility.ListToDataTable(testName, status, Constants.nullString, _dataRow);
                }
                else if (TestContext.CurrentTestOutcome == UnitTestOutcome.Inconclusive)
                {
                    Constants.othersCount = Constants.othersCount + 1;
                  //  _captureFailedScreenPath = Utility.FailureScreenShotCapture();
                    Utility.ListToDataTable(testName, status, _captureFailedScreenPath, _dataRow);
                }
                else if (TestContext.CurrentTestOutcome == UnitTestOutcome.Error || TestContext.CurrentTestOutcome == UnitTestOutcome.Aborted)
                {
                    Constants.errorCount = Constants.errorCount + 1;
                 //   _captureFailedScreenPath = Utility.FailureScreenShotCapture();
                    Utility.ListToDataTable(testName, status, _captureFailedScreenPath, _dataRow);
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                //Long report
                Constants.totalRun = Constants.passCount + Constants.failCount + Constants.othersCount + Constants.errorCount;
                Reports.ExportDataTable2Excel(Constants.managementReportName, Color.White, Color.Maroon, 15, true, "01/06/2015", Color.White,
                Color.Black, 10, Constants.globalTable, Color.Orange, Color.White, "MyData",
                Constants.mREPORTPATH);

                Utility.ListAddition(testName, status, _dataRow, Constants.stepsCounter);
            //    Utility.DirectoryCopyToNewDirectory(Constants.globalResultsPath, Constants.globalRecentResultsPath, true);
            }
        }

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
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

        private TestContext testContextInstance;
        private string _captureFailedScreenPath;
        private string _dataRow;
    }
}

