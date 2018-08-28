using System;
using System.Drawing;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SFB.LibraryFunctions;
using NGW_SharePoint.ObjectRepository.SampleTest;
using NGW_SharePoint.Utility;
using System.Diagnostics;
using System.Threading;
using System.IO;
using System.Collections.Generic;
using NGW_SharePoint.ObjectRepository.SampleTest2016;
using System.Globalization;

//If the below test case has to be executed in IST time zone, execution should be given after 11:30 am

namespace SFBTesting.DriverFunctions.SFBTesting
{
    /// <summary>
    /// Summary description for SkypeSigning
    /// </summary>
    [CodedUITest]
    public class SfB_Audio_JoinSkypeMeeting : SampleTestObjectRepository
    {
        public SfB_Audio_JoinSkypeMeeting()
        {
            NGW_SharePoint.Utility.Constants.Environment = objEnvironment;
        }
    private static int iterationintialize = 1;
         private static string DataPath = Constants.DataFilePath;
        private  long data_rows = globalFunctions.CountLinesInFile(DataPath + "JoinSkypeMeeting.csv");
        private static string sr = globalFunctions.GetLine(DataPath + "JoinSkypeMeeting.csv", iterationintialize);
        private  string[] values = sr.Split(',');

        private  long data_rows1 = globalFunctions.CountLinesInFile(DataPath + "Skypeversion.csv");
        private static string sr1 = globalFunctions.GetLine(DataPath + "Skypeversion.csv", iterationintialize);
        private  string[] values2 = sr1.Split(',');
        private  string filePath = "C:\\SFB_Automation\\";

        [TestInitialize]
        public void MyTestInitialize()
        {
            try
            {
           
                //Skype
                Process[] processes = System.Diagnostics.Process.GetProcessesByName("lync");
                if (processes.Length == 0)
                {
                    globalFunctions.StartLyncProcess(values2[1].ToString(), values2[0], values2[2], values2[3], values2[4], 1);
                }
                else
                {
                    WinWindow SkypeForBusiness = globalFunctions.GetWindowByName(objbuttonname, 1);
                    if (globalFunctions.WinTextNotExist(SkypeForBusiness, values2[1].ToString(), 1))
                    {
                        //Kill
                        globalFunctions.StopProcess("lync");
                        globalFunctions.StartLyncProcess(values2[1].ToString(), values2[0], values2[2], values2[3], values2[4], 1);
                    }
                }

                //Outlook
                Process[] process = System.Diagnostics.Process.GetProcessesByName("OUTLOOK");
                if (process.Length == 0)
                {
                    globalFunctions.StartOutlookProcess(values2[11].ToString(), 1, "Start Outlook Process");
                }
                else
                {

                    //  Launch Outlook
                    bool Outlook = globalFunctions.GetWindowByNameAndClassNameIfWindowExist(values2[11].ToString() + " - Outlook", objOutlookclassname, 1);

                    if (Outlook == false)
                    {
                        //Kill
                        globalFunctions.StopProcess("OUTLOOK");
                        globalFunctions.StartOutlookProcess(values2[11].ToString(), 1, "Start Outlook Process");
                    }
                }

                //Remote login
                globalFunctions.RemoteConnection(values2[6].ToString(), values2[7].ToString(), values2[8].ToString(), values2[9].ToString(), values2[0].ToString(), values2[5].ToString(), 1);

                //bool OutlookMinexist = false;
                ////Start Outlook
                //globalFunctions.SearchInCitrixWindowAndDoubleClick(values2[5].ToString(), filePath + "Images\\OutlookImage.png", 0.80f, 1, 1, "Outlook Shortcut");
                ////Minimize Outlook
                //OutlookMinexist = globalFunctions.SearchInCitrixWindowAndDoubleClick(values2[5].ToString(), filePath + "Images\\MinimizeOutlook.png", 0.80f, 0, 1, "Minimize Outlook");
                //if (!OutlookMinexist)
                //{
                //    globalFunctions.SearchInCitrixWindowAndDoubleClick(values2[5].ToString(), filePath + "Images\\OutlookMinimize_win7.png", 0.80f, 1, 1, "Alternative Minimize Outlook");
                //}
                // Minimize Skype
                bool exist19 = false;
                bool exist100 = globalFunctions.SearchInCitrixWindow(values2[5].ToString(), filePath + "Images\\SkypeForBusiness.png", 0.97f, 0, 1, "SkypeForBusiness initialize");
                if (!exist100)
                {
                    exist19 = globalFunctions.SearchInCitrixWindow(values2[5].ToString(), filePath + "Images\\AlternativeSkypeForBusiness.png", 0.97f, 2, 1, "AlternativeSkypeForBusiness initialize");
                }
                if (exist100 == true || exist19 == true)
                {
                    bool exist110 = globalFunctions.SearchInCitrixWindowAndDoubleClick(values2[5].ToString(), filePath + "Images\\Minimize.png", 0.92f, 0, 1, "Minimize initialize");
                    if (!exist110)
                    {
                        bool exist120 = globalFunctions.SearchInCitrixWindowAndDoubleClick(values2[5].ToString(), filePath + "Images\\AlternativeMinimize.png", 0.92f, 2, 1, "AlternativeMinimize initialize");
                        if (!exist120)
                        {
                            bool exist180 = globalFunctions.SearchInCitrixWindowAndDoubleClick(values2[5].ToString(), filePath + "Images\\Alternative1Minimize.png", 0.92f, 2, 1, "Aletrnative1Minimize initialize");
                            if (!exist180)
                            {
                                globalFunctions.SearchInCitrixWindowAndDoubleClick(values2[5].ToString(), filePath + "Images\\AlternativeMinimize2.png", 0.92f, 2, 1, "AletrnativeMinimize2 initialize");
                            }
                        }
                    }
                }
                globalFunctions.SearchInCitrixWindowAndDoubleClick(values2[5].ToString(), filePath + "Images\\Dismiss_All.png", 0.90f, 2, 1, "Dismiss_All");
                System.Threading.Thread.Sleep(1000);
                globalFunctions.SearchInCitrixWindow(values2[5].ToString(), filePath + "Images\\Yes.png", 0.90f, 2, 1, "Yes");

                bool shortcutexist = false;
                //Start Skype- by clicking on the Skype shortcut                      
                bool shortcutexist1 = globalFunctions.SearchInCitrixWindowAndDoubleClick(values2[5].ToString(), filePath + "Images\\SkypeForBusinessShortcut.png", 0.90f, 0, 1, "SkypeForBusiness Shortcut");
                if (!shortcutexist)
                {
                    shortcutexist = globalFunctions.SearchInCitrixWindowAndDoubleClick(values2[5].ToString(), filePath + "Images\\SkypeForBusinessShortcut1.png", 0.80f, 1, 1, "Alternative SkypeForBusiness Shortcut");
                }

                //Minimize remote window
                System.Windows.Forms.SendKeys.SendWait("^%{BREAK}");
                globalFunctions.ShowDesktop();
                //globalFunctions.SearchOnDesktop(filePath + "Images\\Remote_Minimize.png", 0.89f, 1, 1, "Remote Desktop Minimize");

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
        public void JoinSkypeMeetingFromOutlook()
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
                    string sr = globalFunctions.GetLine(DataPath + "JoinSkypeMeeting.csv", iteration);
                    string[] values = sr.Split(',');

                    //  Launch Outlook
                    WinWindow Outlook = globalFunctions.GetWindowByNameAndClassName(values[0].ToString() + " - Outlook", objOutlookclassname, iteration);

                    ////Send a Skype meeting request
                    Outlook.SetFocus();

                    //Click on Mail
                    globalFunctions.WinGroupButtonClickByName(Outlook, objNavigation, objMail, iteration);

                    //Click on NewItems Drop down button
                    WinWindow newitemwindow = globalFunctions.GetWindowByPrentAccessibleNameAndClassName(Outlook, objFileAccessibleName, objSeeContactCardMenuItem, iteration, "NewItems Window");

                    globalFunctions.ClickOnToolBarDropDownButton(newitemwindow, objNew, objNewItems, values[0].ToString() + " - Outlook", iteration);

                    // Click on  Skype meeting
                    globalFunctions.WinMenuItemClickByNameAndParent(newitemwindow, objSkypeMeeting, iteration);

                    //Check Reminder window exist or not
                    if (globalFunctions.GetWindowByNameAndClassNameIfWindowExist(objReminder, objMicrosoftClassName, iteration, "reminderwindow"))
                    {
                        //Click on Dismiss all
                        WinWindow reminder = globalFunctions.GetWindowByNameAndClassName(objReminder, objMicrosoftClassName, iteration, "reminderwindow");
                        WinWindow subreminderwindow = globalFunctions.GetWindowByParentControlId(reminder, objReminderControlId, iteration, "subreminderwindow");
                        globalFunctions.WinButtonClickByName(subreminderwindow, "Dismiss All", iteration);

                        //Click on Yes
                        WinWindow Yes = globalFunctions.GetWindowByNameAndClassName(objMicrosoftOutlook, objMicrosoftClassName, iteration, "Micrososftwindow");
                        WinWindow subMicrosoftwindow = globalFunctions.GetWindowByControlIdAndTitle(Yes, objYesControlId, objMicrosoftOutlook, iteration, "subMicrososftwindow");
                        globalFunctions.WinButtonExistClick(subMicrosoftwindow, objYesButton, objMicrosoftOutlook, iteration);

                    }
                    //Enter the To address
                    WinWindow meetingwindow = globalFunctions.GetWindowByNameAndClassName(objUntitledMeeting, objOutlookclassname, iteration);
                    WinWindow submeetingwindow = globalFunctions.GetWindowByPrentAccessibleNameAndClassName(meetingwindow, objFileAccessibleName, objSeeContactCardMenuItem, iteration, "To: address");
                    globalFunctions.WinEditByName(submeetingwindow, objTo, values[1].ToString(), iteration, "To address ");

                    Random rnd = new Random();
                    int Value = rnd.Next(1000, 1000000);
                    //Enter the subject
                    WinWindow subjectwindow = globalFunctions.GetWindowByParentControlId(meetingwindow, objMeetinControlId, iteration, "Subject window");
                    globalFunctions.WinEditByName(subjectwindow, objSubjectname, values[2].ToString() + Value, iteration, "Subject");

                    // string time = DateTime.Now.ToString("hh:mm");
                    // current date
                    string date = DateTime.Today.ToString("MM/dd/yyyy");

                // add 3 minutes to the current time
                 DateTime localTime = DateTime.Now.AddMinutes(3);
                    string startTime24Hour = localTime.ToString("HH:mm", CultureInfo.CurrentCulture);

                    //add 30 minutes to the current time
                    DateTime time = DateTime.Now.AddMinutes(30);
                    string endTime24Hour = time.ToString("HH:mm", System.Globalization.CultureInfo.CurrentCulture);

                    //Enter the today's date(START DATE)
                    WinWindow subjectmeetingwindow = globalFunctions.GetWindowByNameAndClassName(" - Meeting  ", objOutlookclassname, iteration);
                    WinWindow subjectmeetingsubwindow = globalFunctions.GetWindowByParentControlId(subjectmeetingwindow, objStartDateControlId, iteration, "Start date window");
                    globalFunctions.WinEditByName(subjectmeetingsubwindow, objStartDate, date, iteration, "Start Date");

                    //Enter the start time + 3 minutes(START TIME)
                    WinWindow subjectmeetingstarttimewindow = globalFunctions.GetWindowByParentControlId(subjectmeetingwindow, objStartTimeControlId, iteration, "Start time window");
                    globalFunctions.WinEditByName(subjectmeetingstarttimewindow, objStartTime, startTime24Hour, iteration, "Start time");

                    //Enter the today's date(END DATE)
                    //  WinWindow subjectmeetingwindow = globalFunctions.GetWindowByNameAndClassName("SFB - Meeting  ", "rctrl_renwnd32", iteration);
                    WinWindow enddatewindow = globalFunctions.GetWindowByParentControlId(subjectmeetingwindow, objEmailControlId, iteration, "End date window");
                    globalFunctions.WinEditByName(enddatewindow, objEndDate, date, iteration, "End date");

                    //WinWindow obj=globalFunctions.GetWindowByName("Microsoft Outlook",1);
                    if (globalFunctions.GetWindowByNameAndClassNameIfWindowExist("Microsoft Outlook", "#32770", 1))
                    {
                        WinWindow microsoftOutlookWindowPopUp = new WinWindow();
                        microsoftOutlookWindowPopUp.SearchProperties.Add(WinWindow.PropertyNames.Name, "Microsoft Outlook");
                        microsoftOutlookWindowPopUp.SetFocus();

                        WinWindow OKWindowPopUp = new WinWindow(microsoftOutlookWindowPopUp);
                        OKWindowPopUp.SearchProperties.Add(WinWindow.PropertyNames.ControlType, "Window");
                        OKWindowPopUp.SearchProperties.Add(WinWindow.PropertyNames.ControlId, "2");

                        WinButton OKButton = new WinButton(OKWindowPopUp);
                        OKButton.SearchProperties.Add(WinWindow.PropertyNames.Name, "OK");
                        Mouse.Click(OKButton);

                        //Adding 5 minutes to the local time
                        time = localTime.AddMinutes(5);
                        endTime24Hour = time.ToString("HH:mm", System.Globalization.CultureInfo.CurrentCulture);//"23:45";

                        // Enter End time
                        WinWindow subjectmeetingendtimewindow = globalFunctions.GetWindowByParentControlId(subjectmeetingwindow, objEndTimeControlId, iteration, "End time window");
                        globalFunctions.WinEditByName(subjectmeetingendtimewindow, objEndTime, endTime24Hour, iteration, "End time");

                        //Enter the today's date(END DATE)
                        enddatewindow = globalFunctions.GetWindowByParentControlId(subjectmeetingwindow, objEmailControlId, iteration, "End date window");
                        globalFunctions.WinEditByName(enddatewindow, objEndDate, date, iteration, "End date");

                    }
                    else
                    {
                        WinWindow subjectmeetingendtimewindow = globalFunctions.GetWindowByParentControlId(subjectmeetingwindow, objEndTimeControlId, iteration, "End time window");
                        globalFunctions.WinEditByName(subjectmeetingendtimewindow, objEndTime, endTime24Hour, iteration, "End time");
                    }




                    //Click on Send button
                    WinWindow sendwindow = globalFunctions.GetWindowByParentControlId(subjectmeetingwindow, objSendControlId, iteration, "Send window");
                    globalFunctions.WinButtonClickByName(sendwindow, objSendname, iteration);

                    ////Join meeting after getting the invitation
                    Outlook.SetFocus();

                    
                    //Check Reminder window exist or not
                    if (globalFunctions.GetWindowByNameAndClassNameIfWindowExist(objReminder, objMicrosoftClassName, iteration, "reminderwindow"))
                    {
                        //Click on Dismiss all
                        WinWindow reminder = globalFunctions.GetWindowByNameAndClassName(objReminder, objMicrosoftClassName, iteration, "reminderwindow");
                        WinWindow subreminderwindow = globalFunctions.GetWindowByParentControlId(reminder, objReminderControlId, iteration, "subreminderwindow");
                        globalFunctions.WinButtonClickByName(subreminderwindow, "Dismiss All", iteration);

                        //Click on Yes
                        WinWindow Yes = globalFunctions.GetWindowByNameAndClassName(objMicrosoftOutlook, objMicrosoftClassName, iteration, "Micrososftwindow");
                        WinWindow subMicrosoftwindow = globalFunctions.GetWindowByControlIdAndTitle(Yes, objYesControlId, objMicrosoftOutlook, iteration, "subMicrososftwindow");
                        globalFunctions.WinButtonExistClick(subMicrosoftwindow, objYesButton, objMicrosoftOutlook, iteration);

                    }

                    //Click on Calendar
                    globalFunctions.WinGroupButtonClickByName(Outlook, objNavigation, objCalendar, iteration);

                    // Click on Go To ToolBar
                    globalFunctions.ClickOnToolBarButton(newitemwindow, objGoTo, objGoToDate, iteration);

                    //Enter the date
                    WinWindow gotowindow = globalFunctions.GetWindowByNameAndClassName(objGoTodate, objMicrosoftClassName, iteration);
                    WinWindow gotodatewindow = globalFunctions.GetWindowByParentControlId(gotowindow, objStartTimeControlId, iteration, "Go To Date Sub window");
                    globalFunctions.WinEditByNameWithOutPassingEnterKey(gotodatewindow, objDate, date, iteration, "Date:");
                    // globalFunctions.WinEditByName(gotodatewindow, "Date:", date, iteration);

                    //Select Day Calendar
                    WinWindow gotoselectwindow = globalFunctions.GetWindowByParentControlId(gotowindow, objSelectControlId, iteration, "Go To Select window");
                    globalFunctions.GetComboBoxSelecetedItemSelectComboBox(gotoselectwindow, objShowin, objGoTodate, objDayCalendar, iteration, "Day Calender");

                    //Click on OK button
                    WinWindow OKwindow = globalFunctions.GetWindowByParentControlId(gotowindow, objOKSubWindowDialogControlId, iteration, " OK Window");
                    globalFunctions.WinButtonClickByName(OKwindow, objOKButton, iteration);

                    //select the meeting
                    string month = DateTime.Now.ToString("MMMM");
                    string todaydate = DateTime.Today.ToString("dd");
                    string presentyear = DateTime.Today.ToString("yyyy");
                    if (todaydate.StartsWith("0"))
                    {
                        //  todaydate.Trim();
                        todaydate = todaydate.Remove(0, 1);
                    }

                    var tomdate = DateTime.Now.AddDays(1);

                    //Click On Search Bar and enter Meeting ID

                    WinWindow calendarWindow = globalFunctions.GetWindowByNameAndClassNameWithoutTitle(objScheduleAMeeting, objOutlookclassname, iteration);
                    globalFunctions.WinClientAndEdit(calendarWindow, objSearchQuery, objOutlookclassname, values[2].ToString() + Value, iteration);

                    //Select the meeting from Calender
                    //  globalFunctions.ClickOnListItem(Outlook, todaydate, month, presentyear, startTime24Hour, endTime24Hour, Value, values[0].ToString(), values[2].ToString(), values[3].ToString(), iteration);


                    //Click on join skype meeting
                    WinWindow win = globalFunctions.GetWindowByNameAndClassNameWithoutTitle(values[2].ToString() + Value, objOutlookclassname, iteration);
                    WinWindow win1 = globalFunctions.GetWindowByPrentAccessibleNameAndClassName(win, objFileAccessibleName,objSeeContactCardMenuItem, iteration);
                    globalFunctions.ClickOnToolBarButton(win1, objSkypeMeeting, objJoinSkypeMeeting, iteration);

                    //Minimize
                    WinWindow SkypeForBusiness = globalFunctions.GetWindowByName(NGW_SharePoint.Utility.Constants.SFBName, iteration);
                    globalFunctions.WinButtonClickByParentClassNameAndButtonName(SkypeForBusiness, objSeeContactCardMenuItem, objMinimize, 1);

                    //Click on OK Button on Join Meeting Audio window
                    if (globalFunctions.GetWindowByNameAndClassNameIfWindowExist(objJoinMeetingAudio, objSkypeWindowClassName, iteration))
                    {
                        WinWindow JoinMeetingAudio = globalFunctions.GetWindowByNameAndClassName(objJoinMeetingAudio, objSkypeWindowClassName, iteration);
                        globalFunctions.WinDialogClickByName(JoinMeetingAudio, objJoinMeetingAudio, objOKButton, iteration);

                    }
                    System.Threading.Thread.Sleep(2000);

                    ////To Open and maximize remote window
                    //System.Windows.Forms.SendKeys.SendWait("%{TAB}");
                    

                    //citrix window
                    //Click On Remote window Icon
                    WinWindow citrixwindow = globalFunctions.GetWindowByNameAndClassName(objtoolbarname, objwindowclassname, iteration);
                    if (objSkypeversion == values2[0])
                    {
                        // click on the icon 
                        globalFunctions.ClickOnToolBarButton(citrixwindow, objtoolbarname, values2[5].ToString(), iteration);
                    }
                    else
                    {
                        SampleTestObjectRepository2016 sample = new SampleTestObjectRepository2016();
                        globalFunctions.ClickOnToolBarButton(citrixwindow, objtoolbarname, sample.objRemoteDesktopConnection, iteration);
                    }
                    WinWindow maximizeWindow = globalFunctions.GetWindowByNameAndClassName(values2[5].ToString(), objRemWindowMaxClassNAme, iteration);
                    globalFunctions.ClickOnWindowButtonIfExist(maximizeWindow, objMaximize, iteration);

                    //Click on Join Online
                    System.Threading.Thread.Sleep(4000);
                    // Minimize Skype
                    bool exist1 = false;
                    bool exist = globalFunctions.SearchInCitrixWindow(values2[5].ToString(), filePath + "Images\\SkypeForBusiness.png", 0.90f, 0, iteration, "SkypeForBusiness");
                    if (!exist)
                    {
                        exist1 = globalFunctions.SearchInCitrixWindow(values2[5].ToString(), filePath + "Images\\AlternativeSkypeForBusiness.png", 0.90f, 2, iteration, "AlternativeSkypeForBusiness");
                    }
                    if (exist == true || exist1 == true)
                    {
                        bool exist3 = globalFunctions.SearchInCitrixWindowAndDoubleClick(values2[5].ToString(), filePath + "Images\\Minimize.png", 0.90f, 0, iteration, "Minimize");
                        if (!exist3)
                        {
                            bool exist4 = globalFunctions.SearchInCitrixWindowAndDoubleClick(values2[5].ToString(), filePath + "Images\\AlternativeMinimize.png", 0.92f, 2, iteration, "AlternativeMinimize");
                            if (!exist4)
                            {
                                bool exist17 = globalFunctions.SearchInCitrixWindowAndDoubleClick(values2[5].ToString(), filePath + "Images\\Alternative1Minimize.png", 0.92f, 2, iteration, "Aletrnative1Minimize");
                                if (!exist17)
                                {
                                    globalFunctions.SearchInCitrixWindowAndDoubleClick(values2[5].ToString(), filePath + "Images\\AlternativeMinimize2.png", 0.92f, 2, iteration, "AletrnativeMinimize2");
                                }
                            }
                        }
                    }
                    System.Threading.Thread.Sleep(5000);
                    //Click on Join Online
                    globalFunctions.SearchInCitrixWindowAndDoubleClick(values2[5].ToString(), filePath + "Images\\JoinOnline.png", 0.90f, 1, iteration, "JoinOnline");
                    System.Threading.Thread.Sleep(15000);

                    //Click on OK Button
                    globalFunctions.SearchInCitrixWindow(values2[5].ToString(), filePath + "Images\\OK.png", 0.90f, 2, iteration, "OK");

                    //Minimize remote window
                    System.Windows.Forms.SendKeys.SendWait("^%{BREAK}");
                    globalFunctions.ShowDesktop();
                    //globalFunctions.SearchOnDesktopDoubleClick(filePath + "Images\\Remote_Minimize.png", 0.89f, 1, 1, "Remote Desktop Minimize");

                    //Minimize
                    globalFunctions.WinButtonClickByParentClassNameAndButtonName(SkypeForBusiness, objSeeContactCardMenuItem, objMinimize, 1);

                    // Check for the number of participates
                    WinWindow intiatorwindow = globalFunctions.GetWindowByNameAndClassName(values[2].ToString(), LyncConversationWindowClassName, iteration, "Conversation Window");
                    string name = globalFunctions.VerifyNoOfParticipates(intiatorwindow, objSeeContactCardMenuItem, values[2].ToString() + Value, iteration);

                    if (name == "2 Participants")
                    {
                        globalFunctions.Result("2 Participants are present in the meeting", "pass", iteration);
                    }
                    else
                    {
                        globalFunctions.Result("1 Participants are present in the meeting", "fail", iteration);
                    }
                    //End Call if exist
                    globalFunctions.WinButtonHoverByNameAndClickImageIfExist(intiatorwindow, objstopaudio, values[2].ToString(), iteration);

                    
                    //Fetching the system time and placing in data file
                    globalFunctions.InsertDateTimeToDataFile(DataPath + "JoinSkypeMeeting.csv", 4, iteration - 1);


                    ////close window if exist
                    if (globalFunctions.GetWindowByNameAndClassNameIfWindowExist(values[2].ToString(), LyncConversationWindowClassName, iteration, "Conversation Window"))
                    {
                        WinWindow closewindow1 = globalFunctions.GetWindowByNameAndClassName(values[2].ToString(), LyncConversationWindowClassName, iteration, "Conversation Window");
                        WinWindow closewindow = globalFunctions.GetWindowByParentClassName(closewindow1, values[2].ToString(), objSeeContactCardMenuItem, iteration, "Skype for Business");
                        globalFunctions.WinClickByControlTypeIfExist(closewindow, objDialog, objCloseButton, values[2].ToString(), iteration);
                        globalFunctions.GetWindowByAccessibleNameAndClassNameIfExist(objDoYouWantToClose, objSeeContactCardMenuItem, objDoYouWantToClose, objCloseAllTabs, iteration);
                    }

                    //Click on OK if exist
                    globalFunctions.GetWindowByNameAndClassNameIfExist(objbuttonname, objSkypeWindowClassName, objDialogName, objOKButton, iteration);

                    //Minimize
                    globalFunctions.WinButtonClickByParentClassNameAndButtonName(SkypeForBusiness, objSeeContactCardMenuItem, objMinimize, 1);


                    //vdi machine
                    // Close the 

                    //To Open and maximize remote window
                    //System.Windows.Forms.SendKeys.SendWait("%{TAB}");
                    // maximizeWindow = globalFunctions.GetWindowByNameAndClassName(values2[5].ToString(), objRemWindowMaxClassNAme, iteration);
                    //globalFunctions.ClickOnWindowCloseButton(maximizeWindow, objMaximize, iteration);
                    ////System.Windows.Forms.SendKeys.SendWait("^%{BREAK}");

                    #region Click On Remote window Icon
                    if (objSkypeversion == values2[0])
                    {
                        // click on the icon 
                        globalFunctions.ClickOnToolBarButton(citrixwindow, objtoolbarname, values2[5].ToString(), iteration);
                    }
                    else
                    {
                        SampleTestObjectRepository2016 sample = new SampleTestObjectRepository2016();
                        globalFunctions.ClickOnToolBarButton(citrixwindow, objtoolbarname, sample.objRemoteDesktopConnection, iteration);
                    }
                    maximizeWindow = globalFunctions.GetWindowByNameAndClassName(values2[5].ToString(), objRemWindowMaxClassNAme, iteration);
                    globalFunctions.ClickOnWindowButtonIfExist(maximizeWindow, objMaximize, iteration);
                    #endregion
                    // Minimize Skype
                    //bool exist9 = false;
                    //bool exist10 = globalFunctions.SearchInCitrixWindow(values2[5].ToString(), filePath + "Images\\SkypeForBusiness.png", 0.97f, 0, iteration, "SkypeForBusiness");
                    //if (!exist10)
                    //{
                    //    exist9 = globalFunctions.SearchInCitrixWindow(values2[5].ToString(), filePath + "Images\\AlternativeSkypeForBusiness.png", 0.97f, 2, iteration, "AlternativeSkypeForBusiness");
                    //}
                    //if (exist10 == true || exist9 == true)
                    //{
                    //    bool exist11 = globalFunctions.SearchInCitrixWindowAndDoubleClick(values2[5].ToString(), filePath + "Images\\Minimize.png", 0.92f, 0, iteration, "Minimize");
                    //    if (!exist11)
                    //    {
                    //        bool exist12 = globalFunctions.SearchInCitrixWindowAndDoubleClick(values2[5].ToString(), filePath + "Images\\AlternativeMinimize.png", 0.92f, 2, iteration, "AlternativeMinimize");
                    //        if (!exist12)
                    //        {
                    //            bool exist18 = globalFunctions.SearchInCitrixWindowAndDoubleClick(values2[5].ToString(), filePath + "Images\\Alternative1Minimize.png", 0.92f, 2, iteration, "Aletrnative1Minimize");
                    //            if (!exist18)
                    //            {
                    //                globalFunctions.SearchInCitrixWindowAndDoubleClick(values2[5].ToString(), filePath + "Images\\AlternativeMinimize2.png", 0.92f, 2, iteration, "AletrnativeMinimize2");
                    //            }
                    //        }
                    //    }
                    //}
                    //bool clicksuccess = globalFunctions.SearchInCitrixWindow(values2[5].ToString(), filePath + "Images\\Close.png", 0.93f, 0, iteration, "Close");
                    //if (!clicksuccess)
                    //{
                    //    globalFunctions.SearchInCitrixWindow(values2[5].ToString(), filePath + "Images\\AlternativeClose.png", 0.93f, 1, iteration, "AlternativeClose");
                    //}
                    //globalFunctions.SearchInCitrixWindow(values2[5].ToString(), filePath + "Images\\OK.png", 0.93f, 1, iteration, "OK");

                    //// Check for Close all tabs
                    //globalFunctions.SearchInCitrixWindow(values2[5].ToString(), filePath + "Images\\CloseAllTabs.png", 0.97f, 2, iteration, "Close All Tabs Button ");

                    //// Minimize Skype
                    //bool exist13 = false;
                    //bool exist14 = globalFunctions.SearchInCitrixWindow(values2[5].ToString(), filePath + "Images\\SkypeForBusiness.png", 0.97f, 0, iteration, "SkypeForBusiness");
                    //if (!exist14)
                    //{
                    //    exist13 = globalFunctions.SearchInCitrixWindow(values2[5].ToString(), filePath + "Images\\AlternativeSkypeForBusiness.png", 0.97f, 2, iteration, "AlternativeSkypeForBusiness");
                    //}
                    //if (exist14 == true || exist13 == true)
                    //{
                    //    bool exist15 = globalFunctions.SearchInCitrixWindowAndDoubleClick(values2[5].ToString(), filePath + "Images\\Minimize.png", 0.92f, 0, iteration, "Minimize");
                    //    if (!exist15)
                    //    {
                    //        bool exist16 = globalFunctions.SearchInCitrixWindowAndDoubleClick(values2[5].ToString(), filePath + "Images\\AlternativeMinimize.png", 0.92f, 2, iteration, "AlternativeMinimize");
                    //        if (!exist16)
                    //        {
                    //            bool exist19 = globalFunctions.SearchInCitrixWindowAndDoubleClick(values2[5].ToString(), filePath + "Images\\Alternative1Minimize.png", 0.92f, 2, iteration, "Aletrnative1Minimize");
                    //            if (!exist19)
                    //            {
                    //                globalFunctions.SearchInCitrixWindowAndDoubleClick(values2[5].ToString(), filePath + "Images\\AlternativeMinimize2.png", 0.92f, 2, iteration, "AletrnativeMinimize2");
                    //            }
                    //        }
                    //    }
                    //}
                    //powershell details set up, two list order should match 
                    List<string> wordToBeReplaced = new List<string>();
                    wordToBeReplaced.Clear();
                    wordToBeReplaced.Add("$server");
                    wordToBeReplaced.Add("$Username");
                    wordToBeReplaced.Add("$Password");


                    List<string> newWords = new List<string>();
                    newWords.Clear();
                    newWords.Add(values2[9].ToString());
                    newWords.Add(values2[7].ToString());
                    newWords.Add(values2[8].ToString());

                    globalFunctions.ReplaceTheFirstOccurances(DataPath + "CleanupLync.ps1", wordToBeReplaced, newWords, 1);

                    //Kill remote skype and start process only in task manager
                    globalFunctions.RunPowershell(DataPath + "CleanupLync.ps1", 1);
                    globalFunctions.SearchInCitrixWindowAndDoubleClick(values2[5].ToString(), filePath + "Images\\Dismiss_All.png", 0.90f, 1, iteration, "Dismiss_All");
                    System.Threading.Thread.Sleep(1000);
                    globalFunctions.SearchInCitrixWindow(values2[5].ToString(), filePath + "Images\\Yes.png", 0.90f, 1, iteration, "Yes");

                    //Minimize remote window
                    System.Windows.Forms.SendKeys.SendWait("^%{BREAK}");
                    globalFunctions.ShowDesktop();
                    //globalFunctions.SearchOnDesktopDoubleClick(filePath + "Images\\Remote_Minimize.png", 0.89f, 2, 1, "Remote Desktop Minimize");

                }
            }
            catch (Exception exception)
            {
                //{"The playback failed to find the control with the given search properties. Additional Details: \r\nTechnologyName:  'MSAA'\r\nControlType:  'ListItem'\r\nName:  '19. Juni2018,  from 06:55 to  07:22, Subject SFB853, Location Skype Meeting, Organizer automation_Skype user12_CI (CI/CBV), Time Busy, Meeting with others.'\r\n"}

                //19. Juni 2018,  from 06:55 to  07:22, Subject SFB853, Location Skype Meeting, Organizer automation_Skype user12_CI (CI/CBV), Time Busy, Meeting with others.
                _captureFailedScreenPath = Utility.FailureScreenShotCapture();

                globalFunctions.ReturningToLocalDesktop(exception.Message, iterationintialize);
                globalFunctions.StopProcess("lync");
                //Kill Outlook
                globalFunctions.StopProcess("OUTLOOK");

                //powershell details set up, two list order should match 
                List<string> wordToBeReplaced = new List<string>();
                wordToBeReplaced.Clear();
                wordToBeReplaced.Add("$server");
                wordToBeReplaced.Add("$Username");
                wordToBeReplaced.Add("$Password");


                List<string> newWords = new List<string>();
                newWords.Clear();
                newWords.Add(values2[9].ToString());
                newWords.Add(values2[7].ToString());
                newWords.Add(values2[8].ToString());

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



        #region Additional test attributes

        // You can use the following additional attributes as you write your tests:

        ////Use TestInitialize to run code before running each test 
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{        
        //    // To generate code for this test, select "Generate Code for Coded UI Test" from the shortcut menu and select one of the menu items.
        //}

        ////Use TestCleanup to run code after each test has run
        //[TestCleanup()]
        //public void MyTestCleanUp()
        //{        
        //    // To generate code for this test, select "Generate Code for Coded UI Test" from the shortcut menu and select one of the menu items.
        //}

        #endregion
        [TestCleanup()]
        public void MyTestCleanUp()
        {
            string testName = TestContext.TestName.ToString();
            string status = TestContext.CurrentTestOutcome.ToString();
            // Console.WriteLine(status);
            try
            {
                //Checking the current Test Status
                if (TestContext.CurrentTestOutcome == UnitTestOutcome.Failed)
                {
                    NGW_SharePoint.Utility.Constants.failCount = NGW_SharePoint.Utility.Constants.failCount + 1;
                  //  _captureFailedScreenPath = Utility.FailureScreenShotCapture();
                    // Utility.ListToDataTableShort(testName, status, _captureFailedScreenPath, _dataRow);
                    Utility.ListToDataTable(testName, status, _captureFailedScreenPath, _dataRow);
                }
                else if (TestContext.CurrentTestOutcome == UnitTestOutcome.Passed)
                {
                    NGW_SharePoint.Utility.Constants.passCount = NGW_SharePoint.Utility.Constants.passCount + 1;
                    // Utility.ListToDataTableShort(testName, status, Constants.NULLSTRING, _dataRow);
                    Utility.ListToDataTable(testName, status, NGW_SharePoint.Utility.Constants.nullString, _dataRow);
                }
                else if (TestContext.CurrentTestOutcome == UnitTestOutcome.Inconclusive)
                {
                    NGW_SharePoint.Utility.Constants.othersCount = NGW_SharePoint.Utility.Constants.othersCount + 1;
                 //   _captureFailedScreenPath = Utility.FailureScreenShotCapture();
                    // Utility.ListToDataTableShort(testName, status, _captureFailedScreenPath, _dataRow);
                    Utility.ListToDataTable(testName, status, _captureFailedScreenPath, _dataRow);
                }
                else if (TestContext.CurrentTestOutcome == UnitTestOutcome.Error || TestContext.CurrentTestOutcome == UnitTestOutcome.Aborted)
                {
                    NGW_SharePoint.Utility.Constants.errorCount = NGW_SharePoint.Utility.Constants.errorCount + 1;
                  //  _captureFailedScreenPath = Utility.FailureScreenShotCapture();
                    // Utility.ListToDataTableShort(testName, status, _captureFailedScreenPath, _dataRow);
                    Utility.ListToDataTable(testName, status, _captureFailedScreenPath, _dataRow);
                }
            }
            catch (System.Exception)
            {
                throw;
            }
            finally
            {
                //Long report
                NGW_SharePoint.Utility.Constants.totalRun = NGW_SharePoint.Utility.Constants.passCount + NGW_SharePoint.Utility.Constants.failCount + NGW_SharePoint.Utility.Constants.othersCount + NGW_SharePoint.Utility.Constants.errorCount;
                Reports.ExportDataTable2Excel(NGW_SharePoint.Utility.Constants.managementReportName, Color.White, Color.Maroon, 15, true, "01/06/2015", Color.White,
                Color.Black, 10, NGW_SharePoint.Utility.Constants.globalTable, Color.Orange, Color.White, "MyData",
                NGW_SharePoint.Utility.Constants.mREPORTPATH);

                Utility.ListAddition(testName, status, _dataRow, NGW_SharePoint.Utility.Constants.stepsCounter);
              //  Utility.DirectoryCopyToNewDirectory(Constants.globalResultsPath, Constants.globalRecentResultsPath, true);
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

        public UIMap UIMap
        {
            get
            {
                if ((this.map == null))
                {
                    this.map = new UIMap();
                }

                return this.map;
            }
        }

        private UIMap map;
    }
}





