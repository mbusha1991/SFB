using System;
using System.IO;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.TestTools.UITesting.HtmlControls;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using System.Drawing;
using AForge.Imaging;
using System.Drawing.Imaging;
using NGW_SharePoint.Utility;

using Microsoft.VisualStudio.TestTools.UITest.Extension;
using SFBTesting.LibraryFunctions;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UITesting.WpfControls;
using System.Data.SqlClient;
using System.Security.Principal;
using System.Linq;

namespace SFB.LibraryFunctions
{
    public partial class globalFunctions
    {
        [DllImport("user32.dll")]
        static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, int dwExtraInfo);
        #region Variable Declarations
        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);
        internal static string _logMessage = null;
        internal static string _pass = "Passed";
        internal static string _fail = "Failed";
        internal static string _abort = "Aborted";
        internal static string _methodStatus = null;
        internal static string _comparisonValue = null;
        public static bool result = false;
        public static List<Tuple<String, String, String ,String, String, String>> listOfTuples = new List<Tuple<String,String, String, String, String, String>>();
        public static SqlDataReader reader;

        #endregion
        #region  Function Definitions
 
        public static void ShowDesktop()
        {
            keybd_event(0x5B, 0, 0, 0);
            keybd_event(0x4D, 0, 0, 0);
            keybd_event(0x5B, 0, 0x2, 0);
        }
        public static WinWindow GetWindowByName(string _id, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WindowObj = new WinWindow();
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.Name, _id, WinWindow.PropertyNames.ControlType, "Window");
                _logMessage = String.Concat("Window: " + _id + " is found");
                WindowObj.WaitForControlReady();


                // Mouse.Click(uIHtmlEditObject);
                _methodStatus = _pass;
                return WindowObj;
            }

            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Window: " + _id + " is not found");
                throw;
            }
            finally
            {
                
                listOfTuples.Add(new Tuple<String, String, String ,String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static WinWindow GetWindowByNameWaitTillExist(string _id, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WindowObj = new WinWindow();
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.Name, _id, WinWindow.PropertyNames.ControlType, "Window");
                _logMessage = String.Concat("Window: " + _id + " is found");
                WindowObj.WaitForControlExist();
                WindowObj.WaitForControlReady();


                // Mouse.Click(uIHtmlEditObject);
                _methodStatus = _pass;
                return WindowObj;
            }

            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Window: " + _id + " is not found");
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static WinWindow GetWindowByClassName(string _classname, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WindowObj = new WinWindow();
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _classname, HtmlEdit.PropertyNames.ControlType, "Window");
              //  WindowObj.DrawHighlight();
                _logMessage = String.Concat("Window: "+_callerName +" is found");
                WindowObj.WaitForControlReady();


       
                _methodStatus = _pass;
                return WindowObj;
            }

            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Window: " + _callerName + " is not found");
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void WinTextFocusAndClick(WinText text, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
         
                text.SetFocus();
              //  text.DrawHighlight();
                Mouse.Click(text);

                _logMessage = string.Concat("Clicked on the text: " + text.Name);
                _methodStatus = _pass;
               
            }

            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to clicked on the text: " + text.Name);
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinTextFocusAndHover(WinText text, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {

                text.SetFocus();
               // text.DrawHighlight();
                Mouse.Hover(text);

                _logMessage = string.Concat("Hovered on the text: " + text.Name);
                _methodStatus = _pass;

            }

            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to hovered on the text: " + text.Name);
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void WinTextFocusAndRightClick(WinText text, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {

                text.SetFocus();
                Mouse.Click(text, System.Windows.Forms.MouseButtons.Right);

                _logMessage = string.Concat("Right Clicked on the text: " + text.Name);
                _methodStatus = _pass;

            }

            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to right click on the text: " + text.Name);
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void GetWindowByNameAndClassNameAndClickOnControl(string _id, string _classname, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WindowObj = new WinWindow();
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.Name, _id, PropertyExpressionOperator.Contains);
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _classname);
                WindowObj.WindowTitles.Add(_id);

                _logMessage = String.Concat("Window: " + _id + " is found");
                WindowObj.WaitForControlReady();

                WinClient client = new WinClient(WindowObj);
                client.WindowTitles.Add(_id);

                WinControl control = new WinControl(client);
                control.SearchProperties[WinControl.PropertyNames.ControlType] = "Image";
             //   control.SearchProperties["Instance"] = "2";
                control.WindowTitles.Add(_id);
                //    control.DrawHighlight();
                control.WaitForControlReady();
                //Mouse.Click(control, new Point(controlPt_X, controlPt_Y));

                UITestControlCollection uic = control.FindMatchingControls();

                foreach (UITestControl ui in uic)

                {

                    if (ui.BoundingRectangle.X > 0 && ui.BoundingRectangle.Y > 0 && ui.BoundingRectangle.Width > 300)

                    {
                        //ui.BoundingRectangle = {X = -32000 Y = -31884 Width = 374 Height = 140} ui.BoundingRectangle = {X = 162 Y = 486 Width = 318 Height = 140}
                        //ui.BoundingRectangle = {X = 162 Y = 626 Width = 374 Height = 139}
                        //ui.BoundingRectangle = {X = 474 Y = 782 Width = 46 Height = 46}
                        //ui.BoundingRectangle = {X = 178 Y = 782 Width = 46 Height = 46}
                        // ui.BoundingRectangle = { X = -31864 Y = -31588 Width = 46 Height = 46}
                        //ui.BoundingRectangle = {X = 358 Y = 782 Width = 46 Height = 46}
                        //ui.BoundingRectangle = {X = 192 Y = 461 Width = 16 Height = 16}
                        //ui.BoundingRectangle = {X = 475 Y = 411 Width = 46 Height = 46}
                        //ui.BoundingRectangle = {X = 322 Y = 729 Width = 30 Height = 30}
                        // ui.BoundingRectangle = { X = 366 Y = 736 Width = 16 Height = 16}
                        //ui.BoundingRectangle = {X = 396 Y = 728 Width = 30 Height = 30}
                        //ui.BoundingRectangle = {X = -31727 Y = -31642 Width = 30 Height = 30}
                        //ui.BoundingRectangle = {X = 474 Y = 728 Width = 30 Height = 30}
                        //ui.BoundingRectangle = {X = 177 Y = 421 Width = 10 Height = 10}
                    //    ui.DrawHighlight();
                        Mouse.Click(ui);

                      //  break;

                    }

                }
                _logMessage = String.Concat("Clicked on the " + _callerName);
                _methodStatus = _pass;
              
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Window: " + _id + " is not found");
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }


        public static void GetWindowByNameAndClassNameAndRightClickOnControl(string _id, string _classname, int controlPt_X, int controlPt_Y, string _windowAccessibleName, string _windowclassname, string _menuItemName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WindowObj = new WinWindow();
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.Name, _id);
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _classname);
                WindowObj.WindowTitles.Add(_id);

                _logMessage = String.Concat("Window: " + _id + " is found");
                WindowObj.WaitForControlReady();

                WinClient client = new WinClient(WindowObj);
                client.WindowTitles.Add(_id);

                WinControl control = new WinControl(client);
                control.SearchProperties[UITestControl.PropertyNames.ControlType] = "Image";
                control.SearchProperties["Instance"] = "2";
                control.WindowTitles.Add(_id);
                control.WaitForControlReady();
                Mouse.Click(control, System.Windows.Forms.MouseButtons.Right, System.Windows.Input.ModifierKeys.None, new System.Drawing.Point(controlPt_X, controlPt_Y));

                WinWindow winWindow = new WinWindow();
                winWindow.SearchProperties.Add(WinWindow.PropertyNames.AccessibleName, _windowAccessibleName);
                winWindow.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _windowclassname);
                winWindow.WaitForControlReady();

                WinMenuItem winMenuItem = new WinMenuItem(winWindow);
                winMenuItem.SearchProperties.Add(WinMenuItem.PropertyNames.Name, _menuItemName);
                winMenuItem.SearchProperties.Add(WinMenuItem.PropertyNames.ControlType, "MenuItem");
                winMenuItem.WaitForControlReady();

                WinText winText = new WinText(winMenuItem);
                winText.SearchProperties.Add(WinText.PropertyNames.Name, _menuItemName);
                winText.SearchConfigurations.Add(Microsoft.VisualStudio.TestTools.UITest.Extension.SearchConfiguration.ExpandWhileSearching);
               // winText.DrawHighlight();
                Mouse.Click(winText);
                _logMessage = String.Concat("Clicked on menu item:" + _menuItemName);
                _methodStatus = _pass;

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Window: " + _id + " is not found");
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        //public static void GetWindowByNameAndClassNameAndClickOnControl(string _id, string _classname, int controlPt_X, int controlPt_Y, string _windowAccessibleName, string _windowclassname, string _menuItemName, int _iteration, [CallerMemberName] string _callerName = null)
        //{
        //    try
        //    {
        //        WinWindow WindowObj = new WinWindow();
        //        WindowObj.SearchProperties.Add(WinWindow.PropertyNames.Name, _id);
        //        WindowObj.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _classname);
        //        WindowObj.WindowTitles.Add(_id);

        //        _logMessage = String.Concat("Window: " + _id + " is found");
        //        WindowObj.WaitForControlReady();

        //        WinClient client = new WinClient(WindowObj);
        //        client.WindowTitles.Add(_id);

        //        WinControl control = new WinControl(client);
        //        control.SearchProperties[UITestControl.PropertyNames.ControlType] = "Image";
        //        control.SearchProperties["Instance"] = "2";
        //        control.WindowTitles.Add(_id);
        //        control.WaitForControlReady();
        //        Mouse.Click(control, new Point(controlPt_X, controlPt_Y));


        //        _logMessage = String.Concat("Clicked on menu item:" + _menuItemName);
        //        _methodStatus = _pass;

        //    }
        //    catch (Exception)
        //    {
        //        _methodStatus = _fail;
        //        _logMessage = string.Concat("Window: " + _id + " is not found");
        //        throw;
        //    }
        //    finally
        //    {

        //        listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
        //    }
        //}

        public static WinWindow GetWindowByAccessibleNameAndClassName(string _windowAccessibleName, string _windowclassname, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WindowObj = new WinWindow();
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.AccessibleName, _windowAccessibleName);
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _windowclassname);

                _logMessage = String.Concat("Window: " + _callerName + _windowAccessibleName + " is found");
                WindowObj.WaitForControlExist();
                WindowObj.WaitForControlReady();
    
                _methodStatus = _pass;
                return WindowObj;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to find the window: " + _callerName + _windowAccessibleName);

                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static WinWindow GetWindowByPrentAccessibleNameAndClassName(UITestControl Parent ,string _windowAccessibleName, string _windowclassname, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WindowObj = new WinWindow(Parent);
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.AccessibleName, _windowAccessibleName);
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _windowclassname);

                _logMessage = String.Concat("Window: " + _callerName + " is found");
                WindowObj.WaitForControlExist();
                WindowObj.WaitForControlReady();

                _methodStatus = _pass;
                return WindowObj;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to find the window: " + _callerName);

                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static WinWindow GetWindowByParentControlIdAndInstance(UITestControl Parent, string _windowControlId, string _windowInstance, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WindowObj = new WinWindow(Parent);
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.ControlId, _windowControlId);
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.Instance, _windowInstance);

                _logMessage = String.Concat("Window: " + _callerName + " is found");
               // changed because to pass address in a send file
             //   WindowObj.WaitForControlExist();
             //   WindowObj.WaitForControlReady();

                _methodStatus = _pass;
                return WindowObj;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to find the window: " + _callerName);

                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static WinWindow GetWindowByParentControlId(UITestControl Parent, string _windowControlId, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WindowObj = new WinWindow(Parent);
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.ControlId, _windowControlId);

                _logMessage = String.Concat("Window: " + _callerName + " is found");

                _methodStatus = _pass;
                return WindowObj;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to find the window: " + _callerName);

                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void ClickOnClientsClient(UITestControl Parent,string _client1id, string _client2id, string _title, string _value, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {

                WinClient client = new WinClient(Parent);
                client.SearchProperties.Add(WinClient.PropertyNames.Name, _client1id, PropertyExpressionOperator.Contains);
                client.WindowTitles.Add(_title);
               // client.DrawHighlight();
                WinClient winclient = new WinClient(client);
                winclient.SearchProperties.Add(WinClient.PropertyNames.Name, _client2id, PropertyExpressionOperator.Contains);
                winclient.WindowTitles.Add(_title);
                //winclient.DrawHighlight();
                winclient.WaitForControlReady();
                Mouse.Click(winclient);
                Keyboard.SendKeys(_value);
              //  client.SetProperty("Value", _value);
                _logMessage = String.Concat("Clicked on Text Box");
                _methodStatus = _pass;

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Window: " + _client1id + " is not found");
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

       

        public static void ClickOnButtonByName(UITestControl Parent, string _Buttonname, int _iteration, [CallerMemberName] string _callerName = null)
        {
          try
          {       
                WinClient client = new WinClient(Parent);
                client.SearchProperties.Add(WinClient.PropertyNames.ControlType, "Client");
                client.WaitForControlReady();

                WinButton winButton = new WinButton(client);
                winButton.SearchProperties[UITestControl.PropertyNames.ControlType] = "Button";
                winButton.SearchProperties[UITestControl.PropertyNames.Name] = _Buttonname;
                winButton.WaitForControlReady();
                    if (winButton.Exists)
                {
                    winButton.SetFocus();
                    Mouse.Click(winButton);
                }
           
                _logMessage = String.Concat("Clicked on : " + _Buttonname);
                _methodStatus = _pass;

           }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on : " + _Buttonname);

                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void ClickOnListItem(UITestControl Parent, string todaydate,string month,string presentyear,string startTime24Hour,string endTime24Hour,int Value,string OutlookValue,string meetingName, string organizerName,  int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinList list = new WinList(Parent);
                // list.SearchProperties.Add(WinList.PropertyNames.Name, "Day View, " + month + " " + todaydate + ", " + presentyear + ", from 12:00am to " + month + " " + tomdate.Day + ", " + presentyear + " 12:00am");
                //  list.SearchProperties.Add(WinList.PropertyNames.Name, "Day View, " + todaydate + ". " + month + " " + presentyear + ", from 12:00am to " + month + " " + tomdate.Day + ", " + presentyear + " 12:00am", PropertyExpressionOperator.Contains);
                list.SearchProperties.Add(WinList.PropertyNames.Name, "Day View", PropertyExpressionOperator.Contains);

                list.WindowTitles.Add(OutlookValue + " - Outlook");
                //Day View, 21. Juni 2018, from 12:00am to Juni 22, 2018 12:00am
                //Day View, 21. Juni 2018, from 00:00 to 22. Juni 2018 00:00

                WinListItem listitem = new WinListItem(list);
                //   listitem.SearchProperties.Add(WinList.PropertyNames.Name, month + " " + todaydate + ", " + presentyear + ",  from " + startTime24Hour + " to  " + endTime24Hour + ", Subject " + values[2].ToString() + Value + ", Location Skype Meeting, Organizer " + values[3].ToString() + ", Time Busy, Meeting with others.", PropertyExpressionOperator.Contains);
                listitem.SearchProperties.Add(WinList.PropertyNames.Name, todaydate + ". " + month + " " + presentyear + ",  from " + startTime24Hour + " to  " + endTime24Hour + ", Subject " + meetingName + Value + ", Location Skype Meeting, Organizer " + organizerName + ", Time Busy, Meeting with others.", PropertyExpressionOperator.Contains);
                //21. Juni 2018,  from 09:17 to  09:44, Subject SFB450, Location Skype Meeting, Organizer automation_Skype user11_CI (CI/CBV), Time Busy, Meeting with others.
                //21. Juni 2018,  from 09:17 to  09:44, Subject SFB450, Location Skype Meeting, Organizer automation_Skype user12_CI (CI/CBV), Time Busy, Meeting with others.
                //21. Juni 2018,  from 09:09 to  09:36, Subject SFB364, Location Skype Meeting, Organizer automation_Skype user12_CI (CI/CBV), Time Busy, Meeting with others.
                listitem.WindowTitles.Add(OutlookValue + " - Outlook");
                UITestControlCollection uic = listitem.FindMatchingControls();
                foreach (UITestControl ui in uic)
                {
                    if (ui.BoundingRectangle.Width > 0)
                    {
                        Point location = ui.BoundingRectangle.Location;
                        location.Offset(ui.BoundingRectangle.Width / 2,
                        ui.BoundingRectangle.Height / 2);
                        //  Mouse.Click(location);
                        Mouse.Hover(location);
                        Mouse.Click(System.Windows.Forms.MouseButtons.Right, System.Windows.Input.ModifierKeys.None, location);

                        _logMessage = String.Concat("Clicked on : " + meetingName + Value);
                        _methodStatus = _pass;
                    }
                }
                if(uic.Count ==0)
                {
                    _methodStatus = _fail;
                    _logMessage = string.Concat("Failed to click on : " + meetingName + Value);

                }



            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on : listitem");

                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void ClickOnButtonByNameIfExist(UITestControl Parent, string _Buttonname, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinClient client = new WinClient(Parent);
                client.SearchProperties.Add(WinClient.PropertyNames.ControlType, "Client");
                client.WaitForControlReady();

                WinButton winButton = new WinButton(client);
                winButton.SearchProperties[UITestControl.PropertyNames.ControlType] = "Button";
                winButton.SearchProperties[UITestControl.PropertyNames.Name] = _Buttonname;
                winButton.WaitForControlReady();
                if (winButton.Exists)
                {
                    winButton.SetFocus();
                    Mouse.Click(winButton);
                    _logMessage = String.Concat("Clicked on : " + _Buttonname);
                    _methodStatus = _pass;
                }
                else
                {
                 //   _logMessage = String.Concat(_Buttonname + " button not found");
                    _methodStatus = _pass;
                }
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on : " + _Buttonname);

                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }


        public static void OpenFileAndCopy(string _id, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                FileInfo fi = new FileInfo(_id);
                if (fi.Exists)
                {
                    Process p = System.Diagnostics.Process.Start(_id);
                    p.WaitForInputIdle();
                    IntPtr h = p.MainWindowHandle;
                    SetForegroundWindow(h);

                    //   Mouse.Click();
                    //   SendKeys.SendWait("^a");
                    //   System.Threading.Thread.Sleep(100);
                    ////   SendKeys.SendWait("^c");
                    //   SendKeys.Send("^(C)");
                    Clipboard.SetText(File.ReadAllText(_id));
                    _logMessage = String.Concat("Copied the data");
                    _methodStatus = _pass;
                }
                else
                {
                    //file doesn't exist
                    _logMessage = String.Concat("file doesn't exist");
                    _methodStatus = _fail;
                }             

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Copy operation failed");
                throw;
            }
            finally
            {
                
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static string OpenFileAndCopyWord(string _id, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                Microsoft.Office.Interop.Word.ApplicationClass wordApp = new Microsoft.Office.Interop.Word.ApplicationClass();
                // string filename = @"D:\Good afternoon.docx";
                string filename = @_id;
                object file = filename;
                string allText = string.Empty;
                object nullobj = System.Reflection.Missing.Value;

                Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Open(ref file, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj);

                doc.ActiveWindow.Selection.WholeStory();

                doc.ActiveWindow.Selection.Copy();
                IDataObject data = Clipboard.GetDataObject();

                bool conv = true;
                allText = data.GetData(DataFormats.Text, conv).ToString();

                doc.Close(ref nullobj, ref nullobj, ref nullobj);
                wordApp.Quit(ref nullobj, ref nullobj, ref nullobj);
                return allText;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Copy operation failed");
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinControlHoverByParentButtonNameAndWindowName(string _id, string _ClassName, string _ButtonName, int iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {

                WinWindow WinImageWindow = new WinWindow();
                WinImageWindow.SearchProperties.Add(WinWindow.PropertyNames.Name, _id);
                WinImageWindow.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _ClassName, "ControlType", "Window");
                WinImageWindow.WindowTitles.Add(_id);

                WinButton button = new WinButton(WinImageWindow);
                button.SearchProperties.Add(WinButton.PropertyNames.Name, _ButtonName);
                button.SearchProperties.Add(WinButton.PropertyNames.ControlType, "Button");
                button.WindowTitles.Add(_id);
                button.WaitForControlReady();

                WinControl control = new WinControl(button);
                control.SearchProperties.Add(WinControl.PropertyNames.ControlType, "Image");
                control.WindowTitles.Add(_id);
                //   control.DrawHighlight();
                Mouse.HoverDuration = 100;
                Mouse.Hover(control);

                _logMessage = string.Concat("Clicked on " + _ButtonName + " button");


                _methodStatus = _pass;
            }
            catch (System.Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _ButtonName + " button");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), iteration.ToString()));
            }
        }

        public static void ReturningToLocalDesktop(string message, int iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                if (message.Contains("A generic error occurred in GDI+.") || message.Contains("Assert.Fail failed. Result string is not passed") || message.Contains("Image matching is not found "))
                {
                    //Minimize remote window
                    System.Windows.Forms.SendKeys.SendWait("^%{BREAK}");
                    globalFunctions.ShowDesktop();

                    _methodStatus = _pass;
                    _logMessage = string.Concat("Passed: to return to local desktop from remote machine");
                }

            }
            catch (System.Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to return to local desktop from remote machine");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), iteration.ToString()));
            }
        }
        public static void WinListItemHoverByNameAndWindowsTitleCompareFriendlyName(UITestControl Parent, string _objectName, string comparename, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinListItem _WinListItem = new WinListItem(Parent);
                _WinListItem.SearchProperties.Add(WinListItem.PropertyNames.Name, _objectName, PropertyExpressionOperator.Contains);

                //  _WinListItem.WindowTitles.Add(_windowtitle);
                _WinListItem.SetFocus();
                // _WinListItem.DrawHighlight();
                Mouse.HoverDuration = 100;
                Mouse.Hover(_WinListItem);

                string Name = _WinListItem.GetProperty("FriendlyName").ToString();
                string[] Namesplit = Regex.Split(Name.Trim(), "-");

                if (Namesplit[1].Trim().Contains(comparename))
                {
                    _logMessage = string.Concat(Name + _callerName);
                    _methodStatus = _pass;
                }
                else
                {
                    // Assert.Fail(_objectName + "is not found " + _callerName);
                    _logMessage = string.Concat(Name + "- Sync issue between Outlook and Skype");
                    _methodStatus = _pass;

                }
            }
            catch (System.Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat(_objectName + " status is not found as" + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static bool WinGroupButtonIfExistClickByName(UITestControl Parent, string _GroupName, string _ButtonName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                bool click = false;

                WinGroup group = new WinGroup(Parent);
                group.SearchProperties.Add(WinControl.PropertyNames.Name, "Group");
                //   group.WindowTitles.Add(_Title);

                WinButton winButton = new WinButton(group);
                winButton.SearchProperties.Add(WinControl.PropertyNames.Name, _ButtonName, "ControlType", "Button");
                if (winButton.Exists)
                {
                    winButton.WaitForControlReady();
                    Mouse.Click(winButton);
                    click = true;
                    System.Threading.Thread.Sleep(1000);
                    _logMessage = string.Concat("Clicked on " + _ButtonName);
                }
                else
                {
                    click = false;
                }

                _methodStatus = _pass;
                return click;
            }
            catch (System.Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _ButtonName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinTreeItemClickByName(UITestControl Parent, string _Name, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinTree wintree = new WinTree(Parent);
                WinTreeItem treeitem = new WinTreeItem(wintree);
                treeitem.SearchProperties.Add(WinTreeItem.PropertyNames.Name, _Name);
                treeitem.SearchProperties.Add(WinTreeItem.PropertyNames.ControlType, "TreeItem");
                Mouse.Click(treeitem);
            }
            catch (System.Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _Name);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static WinWindow GetWindowByParentAndOnlyClassName(UITestControl _parent, string _classname, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WindowObj = new WinWindow(_parent);
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _classname);


                _logMessage = String.Concat("Window: " + _callerName + " is found");
                WindowObj.WaitForControlReady();


                _methodStatus = _pass;
                return WindowObj;
            }
            catch (System.Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Window: " + _callerName + " is not found");
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static WinWindow GetWindowByNameAndClassNameAndTitle(string _id, string _classname, string title,int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WindowObj = new WinWindow();
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.Name, _id, PropertyExpressionOperator.Contains);
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _classname);
                WindowObj.WindowTitles.Add(title);
                WindowObj.WaitForControlReady();
                // WindowObj.WaitForControlExist();
              //  WindowObj.DrawHighlight();
                _logMessage = String.Concat("Window: " + _id + " is found");


                _methodStatus = _pass;
                return WindowObj;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Window: " + _id + " is not found");
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static WinWindow GetWindowByNameAndClassName(string _id, string _classname, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WindowObj = new WinWindow();
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.Name, _id, PropertyExpressionOperator.Contains);
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _classname);
                WindowObj.WindowTitles.Add(_id);
                //changed for schedule meeting from sfb // 
                 WindowObj.WaitForControlExist();
              //  WindowObj.DrawHighlight();
               _logMessage = String.Concat("Window: " + _id + " is found");
                WindowObj.WaitForControlReady();

                _methodStatus = _pass;
                return WindowObj;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Window: " + _id + " is not found");
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static bool GetWindowByNameAndClassNameIfWindowExist(string _id, string _classname, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                bool success = false;
                WinWindow WindowObj = new WinWindow();
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.Name, _id, PropertyExpressionOperator.Contains);
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _classname);
                WindowObj.WindowTitles.Add(_id);
                //changed for schedule meeting from sfb // 
                if (WindowObj.Exists)
                {
                    WindowObj.WaitForControlReady();
                    _logMessage = String.Concat("Window: " + _id + " is found");
                    success = true;
                }
                else
                {
                    _logMessage = String.Concat("Window: " + _id + " is not found");
                    success = false;
                }

                _methodStatus = _pass;
                return success;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Window: " + _id + " is not found");
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static bool GetWindowByNameAndClassNameIfExist(string _id, string _classname,string _DialogName, string _ButtonName,int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                bool success = false;
                WinWindow WindowObj = new WinWindow();
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.Name, _id, PropertyExpressionOperator.Contains);
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _classname);
                WindowObj.WindowTitles.Add(_id);

                System.Threading.Thread.Sleep(500);

                if (WindowObj.Exists)
                {
                    WinControl SkypeBusiness = new WinControl(WindowObj);
                    SkypeBusiness.SearchProperties.Add("ControlType", "Dialog");
                    SkypeBusiness.SearchProperties.Add(WinControl.PropertyNames.Name, _DialogName);
                    SkypeBusiness.WaitForControlReady();
                   // SkypeBusiness.DrawHighlight();

                    WinButton winClick = new WinButton(SkypeBusiness);
                    winClick.SearchProperties.Add(WinButton.PropertyNames.Name, _ButtonName);
                    winClick.WaitForControlReady();

                   // winClick.DrawHighlight();
                   while (winClick.Exists)
                    {
                    Mouse.Click(winClick);
                    _logMessage = string.Concat("Clicked on " + _ButtonName);
                    _methodStatus = _pass;
                   }
                    success = true;
                }
                else
                {
                    _logMessage = String.Concat("Window: " + _id + " is found");
                    _methodStatus = _pass;
                    success = false;
                }
                return success;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Window: " + _id + " is not found");
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static WinWindow GetWindowByNameAndClassNameWithoutTitle(string _id, string _classname, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WindowObj = new WinWindow();
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.Name, _id, PropertyExpressionOperator.Contains);
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _classname);
                WindowObj.WaitForControlReady();
  
                _logMessage = String.Concat("Window: " + _id + " is found");
              
                _methodStatus = _pass;
                return WindowObj;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Window: " + _id + " is not found");
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static WinWindow GetWindowByControlIdAndTitle(UITestControl Parent, string _ControlId, string _WindowTitle ,int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WindowObj = new WinWindow(Parent);
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.ControlId, _ControlId);
                WindowObj.WindowTitles.Add(_WindowTitle);
                WindowObj.WaitForControlReady();
             //   WindowObj.DrawHighlight();
                _logMessage = String.Concat("Window is found: " + _callerName);


                _methodStatus = _pass;
                return WindowObj;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Window is not found: " + _callerName);
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void RadioButtonClickByNameAndTitle(UITestControl Parent, string _Name, string _WindowTitle, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinRadioButton winRadioButton = new WinRadioButton(Parent);
                winRadioButton.SearchProperties[WinRadioButton.PropertyNames.Name] = _Name;
                winRadioButton.WindowTitles.Add(_WindowTitle);
                winRadioButton.WaitForControlReady();
                winRadioButton.SetFocus();
                Mouse.Click(winRadioButton);
                _logMessage = string.Concat("Clicked on Radio Button:  " + _Name);
                _methodStatus = _pass;

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on Radio Button: " + _Name);
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static WinWindow GetWindowByParentNameAndClassName(UITestControl _parent, string _id, string _classname, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WindowObj = new WinWindow(_parent);
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.Name, _id);
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _classname);
                WindowObj.WindowTitles.Add(_id);

                _logMessage = String.Concat("Window: " + _id + " is found");
                WindowObj.WaitForControlReady();


                _methodStatus = _pass;
                return WindowObj;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Window: " + _id + " is not found");
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static WinWindow GetWindowByParentClassName(UITestControl _parent, string _id, string _classname, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WindowObj = new WinWindow(_parent);
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _classname);
                WindowObj.WindowTitles.Add(_id);

                _logMessage = String.Concat("Window: " + _callerName + " is found");
                WindowObj.WaitForControlReady();


                _methodStatus = _pass;
                return WindowObj;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Window: " + _callerName + " is not found");
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static UITestControl GetWindowByParentClassNameAndInstance(UITestControl _parent, string _id, string title, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                UITestControl Selectedwindow = null;
                WinWindow WindowObj = new WinWindow(_parent);
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _id);
                //  WindowObj.SearchProperties.Add(WinWindow.PropertyNames.Instance, _id);
                WindowObj.WindowTitles.Add(title);
                //  WindowObj.DrawHighlight();
                UITestControlCollection uic = WindowObj.FindMatchingControls();
                _logMessage = String.Concat("Window: " + _callerName + " is found");

                foreach (UITestControl ui in uic)

                {

                    if (ui.BoundingRectangle.Width != 185)

                    {

                        Selectedwindow = ui;
                        //ui.BoundingRectangle = { X = 963 Y = 285 Width = 185 Height = 689} ui.BoundingRectangle = {X = 560 Y = 250 Width = 185 Height = 304} ui.BoundingRectangle = {X = 1087 Y = 184 Width = 185 Height = 793}
                        //  ui.BoundingRectangle = { X = -32000 Y = -31853 Width = 757 Height = 691} ui.BoundingRectangle = {X = 355 Y = 209 Width = 197 Height = 306}
                        //{Name [UnitializedBB839B89-49D2-4923-9F10-3C00A9902878], ControlType [Window], NativeControlType [window], ClassName [#32770], RuntimeId [1639768]}
                        //   ui.DrawHighlight();
                        //  object child= ui.GetChildren();
                        //  object instance= ui.GetProperty("Instance");
                        //  break;
                    }

                }

                _methodStatus = _pass;
                return Selectedwindow;

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Window: " + _callerName + " is not found");
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static WinWindow GetWindowByClassNameAndInstance( string _id, string _classname, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WindowObj = new WinWindow();
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _classname);
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.Instance, _id);
    
               // WindowObj.DrawHighlight();
                _logMessage = String.Concat("Window: " + _callerName + " is found");



                _methodStatus = _pass;
                return WindowObj;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Window: " + _callerName + " is not found");
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static WinWindow GetWindowByParentAndName(UITestControl _parent, string _id, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WindowObj = new WinWindow(_parent);
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.Name, _id);
               // WindowObj.DrawHighlight();
                 _logMessage = String.Concat("Window: " + _id + " is found");
                WindowObj.WaitForControlReady();


                _methodStatus = _pass;
                return WindowObj;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Window: " + _id + " is not found");
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinAlertByName(UITestControl _parent, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                //WinClient client = new WinClient(_parent);
                //client.SearchProperties.Add(WinCustom.PropertyNames.ControlType, "Client");
                //WinCustom custom = new WinCustom(client);
                //custom.SearchProperties.Add(WinCustom.PropertyNames.ControlType, "Custom");
                //custom.SearchProperties.Add(WinCustom.PropertyNames.ClassName, "NETUIHWND");
                //client.WaitForControlReady();
                WinControl CallAlert = new WinControl(_parent);
              //  CallAlert.SearchProperties.Add(WinControl.PropertyNames.ClassName, "NETUIHWND", "ControlType", "Alert");
                CallAlert.SearchProperties.Add(WinControl.PropertyNames.Name, "Press space to call, or press the up arrow for more options.", PropertyExpressionOperator.Contains);
                CallAlert.WaitForControlReady();
                Mouse.Click(CallAlert);
 


                //WinEdit uIWinEditObject = new WinEdit(_parent);
                //uIWinEditObject.SearchProperties.Add(WinEdit.PropertyNames.Name, _id, HtmlEdit.PropertyNames.ControlType, "Edit");
                //uIWinEditObject.WaitForControlReady();
                //uIWinEditObject.SetProperty("Value", _value);
                ////Mouse.Click(uIHtmlEditObject);
                //_logMessage = String.Concat("Value " + _value + " entered into input box");
                //Keyboard.SendKeys("{ENTER}");
                //_logMessage = String.Concat("Enter Key is pressed after typing the input");
                //uIWinEditObject.WaitForControlReady();

                _methodStatus = _pass;
            }

            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to enter " + 1 + " into input box");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinTabListClickTabPage(UITestControl _parent, string _tablist, string _tabpage, string _value, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinTabList tablist = new WinTabList(_parent);
                tablist.SearchProperties.Add(WinTabList.PropertyNames.Name, _tablist);
                tablist.WindowTitles.Add(_value);

                WinTabPage tabpage = new WinTabPage(tablist);
                tabpage.SearchProperties.Add(WinTabPage.PropertyNames.Name, _tabpage, PropertyExpressionOperator.Contains);
                tabpage.SearchProperties.Add(WinTabPage.PropertyNames.ControlType, "TabPage");
                tabpage.WindowTitles.Add(_value);
                tabpage.WaitForControlReady();
                Mouse.Click(tabpage);

                _logMessage = String.Concat("Clicked on the Tab page ");
                _methodStatus = _pass;
            }

            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to Click on the Tab page");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinEditByName(UITestControl _parent, string _id, string _value, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinEdit uIWinEditObject = new WinEdit(_parent);
                uIWinEditObject.SearchProperties.Add(WinEdit.PropertyNames.Name, _id, HtmlEdit.PropertyNames.ControlType, "Edit");
                uIWinEditObject.WaitForControlReady();
                uIWinEditObject.SetProperty("Value", _value);
                //Mouse.Click(uIHtmlEditObject);
                _logMessage = String.Concat("Value " + _value + " entered into input box : " + _callerName);
                Keyboard.SendKeys("{ENTER}");
                //     _logMessage = String.Concat("Enter Key is pressed after typing the input");
                uIWinEditObject.WaitForControlReady();

                _methodStatus = _pass;
            }

            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to enter " + _value + " into input box : " + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void WinEditClickByNameAndSendKeys(UITestControl _parent, string _id, string title, string _value, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinEdit uIWinEditObject = new WinEdit(_parent);
                uIWinEditObject.SearchProperties.Add(WinEdit.PropertyNames.Name, _id);
                uIWinEditObject.WindowTitles.Add("Skype for Business Contacts - " + title + " - Outlook");
                System.Threading.Thread.Sleep(6000);
                UITestControlCollection uic = uIWinEditObject.FindMatchingControls();
                _logMessage = String.Concat("Search input box find matching counts " + uic.Count);
                if (uIWinEditObject.Exists)
                {
                    _logMessage = String.Concat("Search input box is found");
                    uIWinEditObject.WaitForControlReady();
                    uIWinEditObject.Text = _value;
                    //Mouse.Click(uIWinEditObject);
                    //Keyboard.SendKeys(uIWinEditObject, _value);

                    _logMessage = String.Concat("Value " + _value + " entered into input box : " + _callerName);
                    Keyboard.SendKeys("{ENTER}");
                    _methodStatus = _pass;
                }
                else
                {
                    _methodStatus = _fail;
                    _logMessage = string.Concat("Failed to find the input box : " + _callerName);
                }
            }

            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to enter " + _value + " into input box : " + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        //public static void WinEditByName(UITestControl _parent, string _id, string _value, int _iteration, [CallerMemberName] string _callerName = null)
        //{
        //    try
        //    {
        //        WinEdit uIWinEditObject = new WinEdit(_parent);
        //        uIWinEditObject.SearchProperties.Add(WinEdit.PropertyNames.Name, _id, HtmlEdit.PropertyNames.ControlType, "Edit");
        //        uIWinEditObject.WaitForControlReady();
        //        uIWinEditObject.SetProperty("Value", _value);

        //        Mouse.Click(uIWinEditObject);

        //        _logMessage = String.Concat("Value "+_value+" entered into input box : " + _callerName);
        //        Keyboard.SendKeys("{ENTER}");
        //   //     _logMessage = String.Concat("Enter Key is pressed after typing the input");
        //        uIWinEditObject.WaitForControlReady();

        //        _methodStatus = _pass;
        //    }

        //    catch (Exception)
        //    {
        //        _methodStatus = _fail;
        //        _logMessage = string.Concat("Failed to enter "+ _value+" into input box : " + _callerName);
        //        throw;
        //    }
        //    finally
        //    {
        //        listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
        //    }
        //}
        public static string WinEditClickByInstanceAndGetText(UITestControl _parent, string _id, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinEdit uIWinEditObject = new WinEdit(_parent);
                uIWinEditObject.SearchProperties.Add(WinEdit.PropertyNames.Instance, _id, HtmlEdit.PropertyNames.ControlType, "Edit");
                uIWinEditObject.WaitForControlReady();
   

                Mouse.Click(uIWinEditObject);

                _logMessage = String.Concat("Clicked on: " + _callerName);
        
                //     _logMessage = String.Concat("Enter Key is pressed after typing the input");
                uIWinEditObject.WaitForControlReady();

               object text = uIWinEditObject.GetProperty("Text");
               _logMessage = String.Concat(" Text is :" + text);
                _methodStatus = _pass;
                return text.ToString();
            }

            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on : " + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static string WinClientEditGetName(UITestControl _parent, string _id,string _instance,int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinClient client = new WinClient(_parent);
                client.SearchProperties.Add(WinClient.PropertyNames.ControlType, "Client");

                WinEdit uIWinEditObject = new WinEdit(client);
                uIWinEditObject.SearchProperties.Add(WinEdit.PropertyNames.Name, _id, HtmlEdit.PropertyNames.ControlType, "Edit");
                uIWinEditObject.SearchProperties.Add(WinEdit.PropertyNames.Instance, _instance);

                uIWinEditObject.WaitForControlReady();

                object text = uIWinEditObject.GetProperty("Text");
                if(text == null)
                {
                    text = "";
                }
                _logMessage = String.Concat(" Text is :" + text);

                _methodStatus = _pass;
                return text.ToString();
            }

            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on : " + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void WinClientAndEdit(UITestControl _parent, string _id, string _className, string meetingSubject, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinClient client = new WinClient(_parent);
                client.SearchProperties.Add(WinClient.PropertyNames.ClassName, _className);

                WinEdit uIWinEditObject = new WinEdit(client);
                uIWinEditObject.SearchProperties.Add(WinEdit.PropertyNames.Name, _id,PropertyExpressionOperator.Contains);

                Mouse.Click(uIWinEditObject);

                Keyboard.SendKeys(meetingSubject);
                Keyboard.SendKeys("{ENTER}");
                Thread.Sleep(6000);
                Keyboard.SendKeys("{TAB 3}");
                Keyboard.SendKeys("{ENTER}");
                Thread.Sleep(6000);
                _logMessage = String.Concat(" Text entered in Search Box :" + meetingSubject);

                _methodStatus = _pass;
               
            }

            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat(" Failed to enter text in Search Box :" + meetingSubject);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinClientEditClickByNameAndInstance(UITestControl _parent, string _id, string _instance, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinClient client = new WinClient(_parent);
                client.SearchProperties.Add(WinClient.PropertyNames.ControlType, "Client");
               // client.DrawHighlight();


                WinEdit uIWinEditObject = new WinEdit(client);
                uIWinEditObject.SearchProperties.Add(WinEdit.PropertyNames.Name, _id, HtmlEdit.PropertyNames.ControlType, "Edit");
                uIWinEditObject.SearchProperties.Add(WinEdit.PropertyNames.Instance, _instance);
                //   uIWinEditObject.DrawHighlight();
                uIWinEditObject.WaitForControlReady();

                Mouse.Click(uIWinEditObject);

                _logMessage = String.Concat("Clicked on : " + _callerName);

                _methodStatus = _pass;

            }

            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on : " + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinEditClickByNameAndKeyboardInput(UITestControl _parent, string _id, string _value, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinEdit uIWinEditObject = new WinEdit(_parent);
                uIWinEditObject.SearchProperties.Add(WinEdit.PropertyNames.Name, _id, HtmlEdit.PropertyNames.ControlType, "Edit");
                uIWinEditObject.WaitForControlReady();
                Mouse.Click(uIWinEditObject);
                Keyboard.SendKeys(_value);

                _logMessage = String.Concat("Value " + _value + " entered into input box : " + _callerName);
                uIWinEditObject.WaitForControlReady();

                _methodStatus = _pass;
            }

            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to enter " + _value + " into input box : " + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
       }

        public static void WinEditByNameWithOutPassingEnterKey(UITestControl _parent, string _id, string _value, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinEdit uIWinEditObject = new WinEdit(_parent);
                uIWinEditObject.SearchProperties.Add(WinEdit.PropertyNames.Name, _id, HtmlEdit.PropertyNames.ControlType, "Edit");
                uIWinEditObject.WaitForControlReady();
                Mouse.Click(uIWinEditObject);
                Keyboard.SendKeys("a", System.Windows.Input.ModifierKeys.Control);
                Keyboard.SendKeys("{Delete}");
                Keyboard.SendKeys(uIWinEditObject, _value);
                _logMessage = String.Concat("Value " + _value + " entered into input box : " + _callerName);
               
                _methodStatus = _pass;
            }

            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to enter " + _value + " into input box : " + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinEditByNameIfExistWithOutPassingEnterKey(UITestControl _parent, string _id, string _value, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinEdit uIWinEditObject = new WinEdit(_parent);
                uIWinEditObject.SearchProperties.Add(WinEdit.PropertyNames.Name, _id, HtmlEdit.PropertyNames.ControlType, "Edit");
               // uIWinEditObject.DrawHighlight();
                if (uIWinEditObject.Exists)
                {
                    uIWinEditObject.WaitForControlReady();
                    Mouse.Click(uIWinEditObject);
                    Keyboard.SendKeys("a", System.Windows.Input.ModifierKeys.Control);
                    Keyboard.SendKeys("{Delete}");
                    //  Keyboard.SendKeys(uIWinEditObject, _value);
                    string specialChar = @"+^%{}#$&.-;'<>_,*";
                   // string normal = @"\|!/()=?»«@£§€";

                    //  @"\|!#$%&/()=?»«@£§€{}.-;'<>_,+*";
                    char[] textarray = _value.ToCharArray();
                    for (int i = 0; i < _value.Length; i++)
                    {
                        if (Char.IsLetterOrDigit(textarray[i]))
                        {
                            Keyboard.SendKeys(uIWinEditObject, textarray[i].ToString());
                        }
                        else
                        {
                            foreach (var item in specialChar)
                            {

                                if (textarray[i].ToString().Contains(item.ToString()))
                                {
                                    char character = item;
                                    int ascii = (int)character;
                                    globalFunctions.KeyBoardActionBasedOnAscii(uIWinEditObject, ascii);
                                    break;
                                }
                                else
                                {
                                    Keyboard.SendKeys(uIWinEditObject, textarray[i].ToString());
                                    break;
                                }
                            }
                        }
                    }
                        _logMessage = String.Concat("Value " + _value + " entered into input box : " + _callerName);

                    _methodStatus = _pass;
                }
                else
                {
                    _logMessage = String.Concat( _callerName + " is not found");

                    _methodStatus = _pass;
                }
            }

            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to enter " + _value + " into input box : " + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        //public static void WinEditByNameIfExistPassword(UITestControl _parent, string _id, string _value, int _iteration, [CallerMemberName] string _callerName = null)
        //{
        //    try
        //    {
        //        WinEdit uIWinEditObject = new WinEdit(_parent);
        //        uIWinEditObject.SearchProperties.Add(WinEdit.PropertyNames.Name, _id, HtmlEdit.PropertyNames.ControlType, "Edit");
        //        // uIWinEditObject.DrawHighlight();
        //        if (uIWinEditObject.Exists)
        //        {
        //            uIWinEditObject.WaitForControlReady();
        //            Mouse.Click(uIWinEditObject);
        //            Keyboard.SendKeys("a", System.Windows.Input.ModifierKeys.Control);
        //            Keyboard.SendKeys("{Delete}");
        //            int unicode = 43;
        //            char character = (char)unicode;
        //            string text = character.ToString();


        //            Keyboard.SendKeys(uIWinEditObject,"Test");
        //            System.Threading.Thread.Sleep(2000);
        //            Keyboard.SendKeys("{+}");
        //            System.Threading.Thread.Sleep(2000);
        //            Keyboard.SendKeys("12345");

        //            _logMessage = String.Concat("Value " + _value + " entered into input box : " + _callerName);

        //            _methodStatus = _pass;
        //        }
        //        else
        //        {
        //            _logMessage = String.Concat(_callerName + " is not found");

        //            _methodStatus = _pass;
        //        }
        //    }

        //    catch (Exception)
        //    {
        //        _methodStatus = _fail;
        //        _logMessage = string.Concat("Failed to enter " + _value + " into input box : " + _callerName);
        //        throw;
        //    }
        //    finally
        //    {
        //        listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
        //    }
        //}

        public static  int hasSpecialChar(string input, int _iteration, [CallerMemberName] string _callerName = null)
        {         
            try
            {
            
                var pwdSpecialCharacterCount = Regex.Matches(input, "[~!@#$%^&*()_+{}:\"<>?]").Count;
                return pwdSpecialCharacterCount;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to enter " + input + " into input box : " + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static bool WinEditByNameIfExistPassword(UITestControl _parent, string _id, string input, 　int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                 bool present = false;
           
                WinEdit uIWinEditObject = new WinEdit(_parent);
                uIWinEditObject.SearchProperties.Add(WinEdit.PropertyNames.Name, _id, HtmlEdit.PropertyNames.ControlType, "Edit");;
                if (uIWinEditObject.Exists)
                {
                    uIWinEditObject.WaitForControlReady();
                    Mouse.Click(uIWinEditObject);
                    Keyboard.SendKeys("a", System.Windows.Input.ModifierKeys.Control);
                    Keyboard.SendKeys("{Delete}");

                    string specialChar = @"@!+^%{}#$&.-;'<>_,*";
                 //   string normal =@"\|!/()=?»«@£§€";

                  //  @"\|!#$%&/()=?»«@£§€{}.-;'<>_,+*";
                    char[] textarray = input.ToCharArray();
                    for (int i = 0; i < input.Length; i++)
                    {
                        if (Char.IsLetterOrDigit(textarray[i]))
                        {
                            Keyboard.SendKeys(uIWinEditObject, textarray[i].ToString());
                        }
                        else
                        {
                            foreach (var item in specialChar)
                            {

                                if (textarray[i].ToString().Contains(item.ToString()))
                                {
                                    char character = item;
                                    int ascii = (int)character;                                  
                                    globalFunctions.KeyBoardActionBasedOnAscii(uIWinEditObject,ascii);
                                    break;
                                }
                            }
                        }
                    }
                }
                else
                {
                    _logMessage = String.Concat("Failed to find the text box" + _callerName);
                    _methodStatus = _pass;

                }
        
                return present;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to enter " + input + " into input box : " + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        
        public static void WinGetTextEditByName(UITestControl _parent, string _id,string value, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinEdit uIWinEditObject = new WinEdit(_parent);
                uIWinEditObject.SearchProperties.Add(WinEdit.PropertyNames.Name, _id, HtmlEdit.PropertyNames.ControlType, "Edit");
              //  uIWinEditObject.DrawHighlight();
                uIWinEditObject.WaitForControlReady();
                object text = null;
                 text = uIWinEditObject.GetProperty(WinEdit.PropertyNames.Text);
                 if (text.ToString() == "")
                { 
                     _logMessage = String.Concat("Validation :" + _callerName + " is empty");
                    _methodStatus = _pass;
                }
                 else if (text.ToString() == value)
                {
                    if (text.ToString() == " ")
                    {
                        _logMessage = String.Concat("Validation :" + _callerName + " is empty");
                        _methodStatus = _pass;
                    }
                    else
                    {
                        _logMessage = String.Concat("Validation : Value " + text + " is present in "  + _callerName );
                        _methodStatus = _pass;
                    }
                }
                else
                {
                    _logMessage = string.Concat("Validation : Falied to enter " + text + " in "+ _callerName);
                    Assert.Fail("Validation : Falied to enter " + text + " in " + _callerName);
                    _methodStatus = _fail;
                }              
            }

            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to enter "+value + " to the " + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void WinGetTextCompareEditByName(UITestControl _parent, string _id, string value,string value1, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinEdit uIWinEditObject = new WinEdit(_parent);
                uIWinEditObject.SearchProperties.Add(WinEdit.PropertyNames.Name, _id, HtmlEdit.PropertyNames.ControlType, "Edit");
                //  uIWinEditObject.DrawHighlight();
                uIWinEditObject.WaitForControlReady();
                object text = null;
                text = uIWinEditObject.GetProperty(WinEdit.PropertyNames.Text);
                if (text.ToString() == "")
                {
                    _logMessage = String.Concat("Validation :" + _callerName + " is empty");
                    _methodStatus = _pass;
                }
                else if (text.ToString() == value || text.ToString() == value1)
                {
                    if (text.ToString() == " ")
                    {
                        _logMessage = String.Concat("Validation :" + _callerName + " is empty");
                        _methodStatus = _pass;
                    }
                    else
                    {
                        _logMessage = String.Concat("Validation : Value " + text + " is present in " + _callerName);
                        _methodStatus = _pass;
                    }
                }
                else
                {
                    _logMessage = string.Concat("Validation : Falied to enter " + text + " in " + _callerName);
                    Assert.Fail("Validation : Falied to enter " + text + " in " + _callerName);
                    _methodStatus = _fail;
                }
            }

            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to enter " + value + " to the " + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static string WinGetFriendlNameEditByName(UITestControl _parent, string _id, string value, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinEdit uIWinEditObject = new WinEdit(_parent);
                uIWinEditObject.SearchProperties.Add(WinEdit.PropertyNames.Name, _id, HtmlEdit.PropertyNames.ControlType, "Edit");
             //   uIWinEditObject.DrawHighlight();
                uIWinEditObject.WaitForControlReady();
                object text = null;
                text = uIWinEditObject.GetProperty(WinEdit.PropertyNames.FriendlyName);
                if (text.ToString() == "")
                {
                    _logMessage = String.Concat("Validation :" + _callerName + " is empty");
                    Assert.Fail("Validation :" + _callerName + " is empty");
                    _methodStatus = _fail;
                }
                else if (text.ToString() == value)
                {
                    if (text.ToString() == " ")
                    {
                        _logMessage = String.Concat("Validation :" + _callerName + "  is empty");
                        Assert.Fail("Validation :" + _callerName + " is empty");
                        _methodStatus = _fail;
                    }
                    else
                    {
                        _logMessage = String.Concat("Validation : Value " + text + "is present in " + _callerName);
                        _methodStatus = _pass;
                    }
                }
                else
                {
                    _logMessage = string.Concat("Validation : Falied to find the " + text + "in " + _callerName);
                    Assert.Fail("Validation : Falied to find the " + text + "in " + _callerName);

                    _methodStatus = _fail;
                }
                return text.ToString();
            }

            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to enter" + value + "to the " + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static string WinGetFriendlNameEditByNameIfExist(UITestControl _parent, string _id, string value, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinEdit uIWinEditObject = new WinEdit(_parent);
                uIWinEditObject.SearchProperties.Add(WinEdit.PropertyNames.Name, _id, HtmlEdit.PropertyNames.ControlType, "Edit");
                object text = null;
                string newtext = string.Empty;
                if (uIWinEditObject.Exists)
                {                 
                    text = uIWinEditObject.GetProperty(WinEdit.PropertyNames.FriendlyName);
                    if (text.ToString() == "")
                    {
                        _logMessage = String.Concat("The " + _callerName + " is empty");
                        _methodStatus = _pass;
                        return text.ToString();
                    }
                    else if (text.ToString() == value)
                    {
                        if (text.ToString() == " ")
                        {
                            _logMessage = String.Concat("The " + _callerName + "  is empty");
                            _methodStatus = _pass;
                            return text.ToString();
                        }
                        else
                        {
                            _logMessage = String.Concat("The " + text + "is present in " + _callerName);
                            _methodStatus = _pass;
                            return text.ToString();
                        }
                    }
                  
                }
                return newtext.ToString();
            }

            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to enter" + value + "to the " + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
     
        public static void WinEditSendKeys(UITestControl _Parent, string _value,int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow win = new WinWindow(_Parent);
                Keyboard.SendKeys(_value);
                _logMessage = string.Concat("Value:" + _value + " entered into the " + _callerName);
                globalFunctions._methodStatus = _pass;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to enter value into " + _callerName+ ", refer exception for more information");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, globalFunctions._logMessage, globalFunctions._methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

 
        public static WinButton WinButtonClickByName(UITestControl Parent, string _objectName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinButton _winClick = new WinButton(Parent);
                _winClick.SearchProperties.Add(WinButton.PropertyNames.Name, _objectName, PropertyExpressionOperator.Contains);
                _winClick.WaitForControlReady();
                // _winClick.DrawHighlight();
                Mouse.Click(_winClick);
                _logMessage = string.Concat("Clicked on " + _objectName);
                //

                _methodStatus = _pass;
                return _winClick;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _objectName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinButtonHoverByName(UITestControl Parent, string _objectName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinButton _winClick = new WinButton(Parent);
                _winClick.SearchProperties.Add(WinButton.PropertyNames.Name, _objectName);
                _winClick.SearchProperties.Add(WinButton.PropertyNames.ControlType, "Button");
                _winClick.SearchProperties.Add(WinButton.PropertyNames.ClassName, "MSTaskListWClass");
              //  _winClick.DrawHighlight();
                 _winClick.SetFocus();
                //  Mouse.HoverDuration = 100;
                   Mouse.Hover(_winClick);
               // Mouse.Click(_winClick);


                _logMessage = string.Concat("Hovered on " + _objectName);
                //_winClick.WaitForControlReady();

                _methodStatus = _pass;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to hover on " + _objectName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void WinListItemDoubleClickByName(UITestControl Parent, string _objectName,int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinListItem _WinListItem = new WinListItem(Parent);
                _WinListItem.SearchProperties.Add(WinListItem.PropertyNames.Name, _objectName, PropertyExpressionOperator.Contains);
                Mouse.DoubleClick(_WinListItem);
                _logMessage = string.Concat("Clicked on " + _objectName);
                //_winClick.WaitForControlReady();

                _methodStatus = _pass;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _objectName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String,String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static WinListItem WinListItemClickByName(UITestControl Parent, string _objectName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinListItem _WinListItem = new WinListItem(Parent);
                _WinListItem.SearchProperties.Add(WinListItem.PropertyNames.Name, _objectName, PropertyExpressionOperator.Contains);
                _WinListItem.WaitForControlReady();
                if (_WinListItem.Exists)
                {
                    Mouse.Click(_WinListItem);
                    _logMessage = string.Concat("Clicked on " + _objectName);
                    _methodStatus = _pass;
                    //_winClick.WaitForControlReady();
                }
                else
                {
                    _logMessage = string.Concat( _objectName + " is not found");
                    _methodStatus = _fail;
                }           
                return _WinListItem;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _objectName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static WinListItem WinListItemByNameAndEditText(UITestControl Parent, string _objectName, string _title, string value, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinListItem _WinListItem = new WinListItem(Parent);
                _WinListItem.SearchProperties.Add(WinListItem.PropertyNames.Name, _objectName, PropertyExpressionOperator.Contains);
                _WinListItem.WindowTitles.Add(_title);
              //  _WinListItem.DrawHighlight();

                WinEdit winedit = new WinEdit(_WinListItem);
                winedit.WindowTitles.Add(_title);
         //       winedit.DrawHighlight();
                Mouse.Click(winedit);
                winedit.SetProperty("Value", value);
                _logMessage = string.Concat("Value set is " + value);
                //_winClick.WaitForControlReady();

                _methodStatus = _pass;
                return _WinListItem;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to set the value " + value);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinListItemRightClickByName( UITestControl Parent, string _objectName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinListItem _WinListItem = new WinListItem(Parent);
                _WinListItem.SearchProperties.Add(WinListItem.PropertyNames.Name, _objectName, PropertyExpressionOperator.Contains);
                // _WinListItem.WaitForControlReady();
                Thread.Sleep(2000);
                Mouse.Click(_WinListItem, System.Windows.Forms.MouseButtons.Right);
                _logMessage = string.Concat("Right clicked on " + _objectName);

             
                _methodStatus = _pass;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat(" Failed to click on " + _objectName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String,String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(),_iteration.ToString()));
            }
        }
        public static void WinListItemHoverByName(UITestControl Parent, string _objectName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinListItem _WinListItem = new WinListItem(Parent);
                _WinListItem.SearchProperties.Add(WinListItem.PropertyNames.Name, _objectName, PropertyExpressionOperator.Contains);
                _WinListItem.WaitForControlReady();
                Mouse.HoverDuration = 100;
                Mouse.Hover(_WinListItem);

                string Name = _WinListItem.GetProperty("DisplayText").ToString();
                if (Name.Contains(_objectName))
                {
                    _logMessage = string.Concat(_objectName + " found in " + _callerName);
                    _methodStatus = _pass;
                }
                else
                {
                    _logMessage = string.Concat(_objectName + " not found " + _callerName);
                    _methodStatus = _pass;

                }
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat( _objectName + " is not found " + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinListItemHoverContainsByName(UITestControl Parent, string _objectName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinListItem _WinListItem = new WinListItem(Parent);
                _WinListItem.SearchProperties.Add(WinListItem.PropertyNames.Name, _objectName, PropertyExpressionOperator.Contains);
                _WinListItem.WaitForControlReady();
                Mouse.HoverDuration = 100;
                Mouse.Hover(_WinListItem);

                string Name = _WinListItem.GetProperty("DisplayText").ToString();
                if (Name.Contains(_objectName))
                {
                    _logMessage = string.Concat(_objectName + " found in " + _callerName);
                    _methodStatus = _pass;
                }
                else
                {
                    _logMessage = string.Concat(_objectName + " not found " + _callerName);
                    _methodStatus = _pass;

                }
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat(_objectName + " is not found " + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinListItemHoverByFriendlyName(UITestControl Parent, string _objectName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinListItem _WinListItem = new WinListItem(Parent);
                _WinListItem.SearchProperties.Add(WinListItem.PropertyNames.Name, _objectName, PropertyExpressionOperator.Contains);
                _WinListItem.WaitForControlReady();
                Mouse.HoverDuration = 100;
                Mouse.Hover(_WinListItem);

                string Name = _WinListItem.GetProperty("FriendlyName").ToString();
                if (_objectName.Equals(Name))
                {
                    _logMessage = string.Concat(_objectName + " found in " + _callerName);
                    _methodStatus = _pass;
                }
                else
                {
                    _logMessage = string.Concat(_objectName + " not found " + _callerName);
                    _methodStatus = _pass;

                }
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat(_objectName + " is not found " + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void WinListItemHoverByNameCompareContains(UITestControl Parent, string _objectName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinListItem _WinListItem = new WinListItem(Parent);
                _WinListItem.SearchProperties.Add(WinListItem.PropertyNames.Name, _objectName, PropertyExpressionOperator.Contains);
                Mouse.HoverDuration = 100;
                Mouse.Hover(_WinListItem);

                string Name = _WinListItem.GetProperty("DisplayText").ToString();
                if (Name.Contains("Collapsed"))
               {
                     Mouse.Click(_WinListItem);
                    _logMessage = string.Concat(Name + " expanded");
                    _methodStatus = _pass;
                }
                else
                {
                    _logMessage = string.Concat(Name + " already expanded");
                    Mouse.Click(_WinListItem);
                    _logMessage = string.Concat(Name + " is collapsed");
                    System.Threading.Thread.Sleep(1000);
                    Mouse.Click(_WinListItem);
                    _logMessage = string.Concat(Name + " is expanded");
                    _methodStatus = _pass;

                } 
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat(_objectName + " is not found " + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinListItemHoverByNameAndGetChildText(UITestControl Parent, string _objectName, string _comparevalue,int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinListItem _WinListItem = new WinListItem(Parent);
                _WinListItem.SearchProperties.Add(WinListItem.PropertyNames.Name, _objectName, PropertyExpressionOperator.Contains);
                Mouse.HoverDuration = 100;
                Mouse.Hover(_WinListItem);
            //    _WinListItem.DrawHighlight();

               WinText wincontrol = new WinText(_WinListItem);
                wincontrol.SearchProperties.Add(WinText.PropertyNames.ControlType, "Text");
            //    wincontrol.DrawHighlight();
                object name = wincontrol.GetProperty("Name");
                string Name = name.ToString();
               string[] namearray = Name.Split('-');

                string trim = namearray[1].TrimStart();
                string[] splitstring = trim.Split(' ');
                if (splitstring[0].Equals(_comparevalue))
                {
                    _logMessage = string.Concat(Name + " :Status comparision is successfull" + _callerName);
                    _methodStatus = _pass;
                }
                else
                {
                    _logMessage = string.Concat(Name + " : Status comparision failed" + _callerName);
                    _methodStatus = _pass;

                }
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat(_objectName + " is not found " + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void WinListItemHoverByNameAndClickText(UITestControl Parent, string _objectName, string _textname, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinListItem _WinListItem = new WinListItem(Parent);
                _WinListItem.SearchProperties.Add(WinListItem.PropertyNames.Name, _objectName, PropertyExpressionOperator.Contains);
                //Mouse.HoverDuration = 100;
                //Mouse.Hover(_WinListItem);
         

                WinText wincontrol = new WinText(_WinListItem);
                wincontrol.SearchProperties.Add(WinText.PropertyNames.ControlType, "Text");
                wincontrol.SearchProperties.Add(WinText.PropertyNames.Name, _textname);
                //   wincontrol.DrawHighlight();
                wincontrol.WaitForControlReady();
                Mouse.Click(wincontrol);
               _logMessage = string.Concat("Clicked on " + _textname);
                _methodStatus = _pass;

                
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _textname);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static string WinListItemHoverByNameAndGetFriendlyName(UITestControl Parent, string _objectName, string _textname, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinListItem _WinListItem = new WinListItem(Parent);
                _WinListItem.SearchProperties.Add(WinListItem.PropertyNames.Name, _objectName, PropertyExpressionOperator.Contains);
                _WinListItem.SetFocus();
                Mouse.HoverDuration = 100;
                Mouse.Hover(_WinListItem);


                WinText wincontrol = new WinText(_WinListItem);
                wincontrol.SearchProperties.Add(WinText.PropertyNames.ControlType, "Text");
                wincontrol.SearchProperties.Add(WinText.PropertyNames.Name, _textname,PropertyExpressionOperator.Contains);
              //  wincontrol.DrawHighlight();

                object name = wincontrol.GetProperty("FriendlyName");
                _logMessage = string.Concat("Friendly Name is " + name);
                _methodStatus = _pass;
                return name.ToString();

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat(_textname + " Friendly name not found");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinListItemHoverByNameAndWindowsTitle(UITestControl Parent, string _objectName, string _windowtitle, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinListItem _WinListItem = new WinListItem(Parent);
                _WinListItem.SearchProperties.Add(WinListItem.PropertyNames.Name, _objectName, PropertyExpressionOperator.Contains);
                                                
                _WinListItem.WindowTitles.Add(_windowtitle);
                _WinListItem.SetFocus();
                Mouse.HoverDuration = 100;
                Mouse.Hover(_WinListItem);

                string Name = _WinListItem.GetProperty("DisplayText").ToString();
                if (_objectName.Equals(Name))
                {
                    _logMessage = string.Concat(_objectName + " found in " + _callerName);
                    _methodStatus = _pass;
                }
                else
                {
                    _logMessage = string.Concat(_objectName + " not found " + _callerName);
                    _methodStatus = _pass;

                }
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat(_objectName + " is not found " + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        

        public static void WinTextRightClickByText(UITestControl Parent, string _objectName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinText _WinListItem = new WinText(Parent);
                _WinListItem.SearchProperties.Add(WinText.PropertyNames.Name, _objectName, PropertyExpressionOperator.Contains);
                Mouse.Click(_WinListItem, System.Windows.Forms.MouseButtons.Right);
                _logMessage = string.Concat("Clicked on " + _objectName);
                //_winClick.WaitForControlReady();

                _methodStatus = _pass;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _objectName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static bool WinTextNotExist(UITestControl Parent, string _objectName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                bool textnotfound = false;
                WinText _WinListItem = new WinText(Parent);
                _WinListItem.SearchProperties.Add(WinText.PropertyNames.Name, _objectName, PropertyExpressionOperator.Contains);
                if (_WinListItem.Exists)
                {
                    Mouse.Hover(_WinListItem);
                    _logMessage = string.Concat("Hovered on " + _objectName);
                    //_winClick.WaitForControlReady();

                    _methodStatus = _pass;
                    textnotfound = false;
                }
                else
                {
                    _logMessage = string.Concat( _objectName+ " is not found");
                    _methodStatus = _pass;
                    textnotfound = true;
                }
                return textnotfound;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Process failed to click on " + _objectName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinScrollBarButtonRightClickByNameAndTitle(UITestControl Parent, string _objScrollbarName,string _objButtonName, string _objTitle, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinScrollBar winscroolbar = new WinScrollBar(Parent);
                winscroolbar.SearchProperties.Add(WinScrollBar.PropertyNames.Name, _objScrollbarName);
                winscroolbar.WindowTitles.Add(_objTitle);
             //   winscroolbar.DrawHighlight();

                WinButton winbutton = new WinButton(winscroolbar);
                winbutton.SearchProperties.Add(WinScrollBar.PropertyNames.Name, _objButtonName);
                winbutton.WindowTitles.Add(_objTitle);
                winbutton.SetFocus();
              //  winbutton.DrawHighlight();
                winbutton.WaitForControlReady();
                Mouse.Click(winbutton, System.Windows.Forms.MouseButtons.Right);

                _logMessage = string.Concat("Right clicked on " + _objButtonName);
           

                _methodStatus = _pass;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to right click  on " + _objButtonName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static bool WinScrollBarButtonExist(UITestControl Parent, string _objScrollbarName, string _objTitle, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinScrollBar winscroolbar = new WinScrollBar(Parent);
                winscroolbar.SearchProperties.Add(WinScrollBar.PropertyNames.Name, _objScrollbarName);
                winscroolbar.WindowTitles.Add(_objTitle);
                //   winscroolbar.DrawHighlight();
                bool success = false;
                if (winscroolbar.Exists)
                {
                    _logMessage = string.Concat(_objScrollbarName + _callerName + " is found");
                    success = true;
                }
                else
                {
                    _logMessage = string.Concat(_objScrollbarName + _callerName + " is not found");
                    success = false;
                }
                _methodStatus = _pass;
                return success;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to find " + _objScrollbarName + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void Result(string result, string value, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                if (value == "pass")
                {
                    _logMessage = string.Concat(result);
                    _methodStatus = _pass;
                }
                else
                {
                    //_logMessage = string.Concat(result);
                    //_methodStatus = _fail;
                    Assert.Fail("Result string is not passed ");
                }

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat(result);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static string WinScrollBarButtonGetTop(UITestControl Parent, string _objScrollbarName, string _objButtonName, string _objTitle, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinScrollBar winscroolbar = new WinScrollBar(Parent);
                winscroolbar.SearchProperties.Add(WinScrollBar.PropertyNames.Name, _objScrollbarName);
                winscroolbar.WindowTitles.Add(_objTitle);
                //   winscroolbar.DrawHighlight();

                WinControl winbutton = new WinControl(winscroolbar);
                winbutton.SearchProperties.Add(WinControl.PropertyNames.Name, _objButtonName);
                winbutton.SearchProperties.Add(WinControl.PropertyNames.ControlType, "Indicator");
                winbutton.WindowTitles.Add(_objTitle);
                winbutton.SetFocus();
              //  winbutton.DrawHighlight();

                object top = winbutton.GetProperty("Top");

            //    _logMessage = string.Concat("Top value is " + top);


                _methodStatus = _pass;
                return top.ToString();
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to get the top value ");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinTextClickByWindowsTitle(UITestControl Parent, string _WindowsTitle, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinText _WinListItem = new WinText(Parent);
                _WinListItem.WindowTitles.Add(_WindowsTitle);
              //  _WinListItem.DrawHighlight();
                _WinListItem.SetFocus();
                UITestControlCollection uic = _WinListItem.FindMatchingControls();

                foreach (UITestControl ui in uic)

                {

                    if (ui.BoundingRectangle.Width > 0)

                    {

                        Mouse.Click(ui);

                        break;

                    }

                }

            //    Mouse.Click(_WinListItem);
                _logMessage = string.Concat("Clicked on " + _callerName);
 

                _methodStatus = _pass;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static List<object> WinListItemHoverByNameAndGetName(UITestControl Parent, string _objectName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {

                List<object> names = new List<object>();
                WinListItem _WinListItem = new WinListItem(Parent);
                _WinListItem.SearchProperties.Add(WinListItem.PropertyNames.Name, _objectName, PropertyExpressionOperator.Contains);
                //_WinListItem.WaitForControlNotExist();
                UITestControlCollection uic = _WinListItem.FindMatchingControls();
                foreach (UITestControl ui in uic)
                {
                    Mouse.Hover(ui);
                    object name = ui.GetProperty("Name");
                    _logMessage = string.Concat("Hovered on " + name);
                    //_winClick.WaitForControlReady();
                    names.Add(name);


                }

                _methodStatus = _pass;
                return names;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to hover on " + _objectName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Parent"></param>
        /// <param name="_objectName"></param>
        /// <param name="_iteration"></param>
        /// <param name="_callerName"></param>
        /// <returns></returns>
        public static WinListItem WinListItemFindByName(UITestControl Parent, string _objectName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinListItem _WinListItem = new WinListItem(Parent);
                _WinListItem.SearchProperties.Add(WinListItem.PropertyNames.Name, _objectName, PropertyExpressionOperator.Contains);
                _WinListItem.WindowTitles.Add("Skype for Business ");
             //   _WinListItem.DrawHighlight();
                _WinListItem.WaitForControlReady();
             //   _WinListItem.SetFocus();
                //if (_WinListItem.Exists)
                //{

                //    _logMessage = string.Concat("Listitem found: " + _objectName);

                //    _methodStatus = _pass;

                //}
                //else
                //{
                //    _methodStatus = _fail;
                //}
                UITestControlCollection uic = _WinListItem.FindMatchingControls();

                foreach (UITestControl ui in uic)

                {

                    if (ui.BoundingRectangle.Width > 0)

                    {
                        Mouse.Click(new System.Drawing.Point(ui.BoundingRectangle.Location.X + 20, ui.BoundingRectangle.Location.Y + 10));

                        break;
                      
                    }

                }
                _logMessage = string.Concat("Clicked on the Listitem: " + _objectName);
                _methodStatus = _pass;
                return _WinListItem;
              
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on the List Item  " + _objectName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        //public static void WinListsListItemRightClick(UITestControl Parent, string _objectName, string title, string _Name, int _iteration, [CallerMemberName] string _callerName = null)
        //{
        //    try
        //    {
        //        WinList _WinList = new WinList(Parent);
        //        _WinList.SearchProperties.Add(WinListItem.PropertyNames.Name, _objectName, PropertyExpressionOperator.Contains);
        //        _WinList.WindowTitles.Add(title);
        //        _WinList.DrawHighlight();


        //        WinListItem _WinListItem = new WinListItem(Parent);
        //        _WinListItem.SearchProperties.Add(WinListItem.PropertyNames.Name, "July 3, 2017,  from 4:30pm to  5:00pm, Subject , Location Skype Meeting, Organize" +
        //                "r Patil Suma Vinod (RBEI/BST9), Time Busy, Meeting with others.", PropertyExpressionOperator.Contains);
        //        _WinListItem.WindowTitles.Add(title);
        //        _WinListItem.DrawHighlight();
        //        Mouse.Hover(_WinListItem);
          
        //        Mouse.Click(_WinListItem, MouseButtons.Right, System.Windows.Input.ModifierKeys.None,new Point(86, 8));

        //        _logMessage = string.Concat("Right clicked on listitem: " + _Name);
        //        //_winClick.WaitForControlReady();

        //        _methodStatus = _pass;
        //    }
        //    catch (Exception)
        //    {
        //        _methodStatus = _fail;
        //        _logMessage = string.Concat("Failed to right click on the listitem: " + _Name);
        //        throw;
        //    }
        //    finally
        //    {
        //        listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
        //    }
        //}
        public static void WinMenuItemHoverByName(UITestControl Parent, string _objectName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinMenuItem WinMenuitem = new WinMenuItem(Parent);
                WinMenuitem.SearchProperties.Add(WinControl.PropertyNames.Name, _objectName, PropertyExpressionOperator.Contains);
                WinMenuitem.WaitForControlReady();
             
                Mouse.HoverDuration = 100;
                Mouse.Hover(WinMenuitem);
                _logMessage = string.Concat(" Hovered on Menu Button: " + _objectName);
                _methodStatus = _pass;

                object text = null;
                text = WinMenuitem.GetProperty(WinMenuItem.PropertyNames.FriendlyName);
                if (text.ToString().Contains( _objectName))
                {

                    _logMessage = string.Concat("Validation:Passed. "+ _callerName  + _objectName);
                    _methodStatus = _pass;
                }
                else
                {
                    _logMessage = string.Concat("Validation:Failed. Failed to find the status: " + _objectName);
                    _methodStatus = _fail;
                }
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to hover on Menu Item: " + _objectName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinMenuButtonClickByName(UITestControl Parent, string _objectName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinControl WinMenuButtonClickByName = new WinControl(Parent);
                WinMenuButtonClickByName.SearchProperties.Add(WinControl.PropertyNames.Name, _objectName, "ControlType", "MenuButton");
                // WinMenuButtonClickByName.WaitForControlReady();
                string[] MenuItemName = _objectName.Split(',');
                _logMessage = string.Concat(" Clicked on Menu Button " + MenuItemName[0]);
                _methodStatus = _pass;
                Mouse.Click(WinMenuButtonClickByName);
             
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _objectName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinMenuButtonByNameValidation(UITestControl Parent, string _objectName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinControl WinMenuButtonClickByName = new WinControl(Parent);
                WinMenuButtonClickByName.SearchProperties.Add(WinControl.PropertyNames.Name, _objectName, "ControlType", "MenuButton");
                // WinMenuButtonClickByName.WaitForControlReady();
                object text = null;
                text = WinMenuButtonClickByName.GetProperty(WinEdit.PropertyNames.Name);
                string[] MenuItemName = _objectName.Split(',');
                if (text.ToString() == _objectName)
                {
                   
                    _logMessage = string.Concat("Validation:Passed. Status is :"+ MenuItemName[0]);
                    _methodStatus = _pass;
                }
                else
                {
                    _logMessage = string.Concat("Validation:Failed. Failed to find the status: "+  MenuItemName[0]);
                    _methodStatus = _fail;
                }
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to find the  " + _objectName + "MenuButton");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void WinMenuItemClickByClassName(UITestControl Parent, string _ClassName, string _MenuItem,int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WinMenuItemWindow = new WinWindow();
                WinMenuItemWindow.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _ClassName, "ControlType", "Window");
                if (_MenuItem.Contains("_"))
                {
                    _MenuItem = _MenuItem.Replace("_", " ");
                }
                
                WinMenuItem MenuItem = new WinMenuItem(WinMenuItemWindow);
                MenuItem.SearchProperties.Add(WinMenuItem.PropertyNames.Name, _MenuItem,PropertyExpressionOperator.Contains);
                MenuItem.WaitForControlReady();
                //newly added because of Send an IM Click
                //    while (MenuItem.Exists)
                //     {
                //   Mouse.Click(MenuItem);
                //      }
                // added because of RenameorDeleteGroup
                //if(MenuItem.Enabled)
                //{     
               Mouse.Click(MenuItem);
                    _logMessage = string.Concat("Clicked on " + _MenuItem);
                    _methodStatus = _pass;
    
                //}
                //else
                //{
                //    _logMessage = string.Concat(_MenuItem + " is disabled");
                //}


            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _MenuItem);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String ,String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(),_iteration.ToString()));
            }
        }

        public static void WinMenuItemClickByClassNameAndSubMenuItemClick(UITestControl Parent, string _ClassName, string _MenuItem, string _submenuitem, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WinMenuItemWindow = new WinWindow();
                WinMenuItemWindow.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _ClassName, "ControlType", "Window");
                if (_MenuItem.Contains("_"))
                {
                    _MenuItem = _MenuItem.Replace("_", " ");
                }

                WinMenuItem MenuItem = new WinMenuItem(WinMenuItemWindow);
                MenuItem.SearchProperties.Add(WinMenuItem.PropertyNames.Name, _MenuItem);
                MenuItem.WaitForControlReady();
          
                Mouse.Click(MenuItem);
                _logMessage = string.Concat("Clicked on " + _MenuItem);
                

                WinGroup winGrp = new WinGroup(MenuItem);
              //  winGrp.DrawHighlight();
                WinMenuItem uINewGroupMenuItem = new WinMenuItem(winGrp);
                uINewGroupMenuItem.SearchProperties.Add(WinMenuItem.PropertyNames.Name, _submenuitem,PropertyExpressionOperator.Contains);
             //   uINewGroupMenuItem.DrawHighlight();
                Mouse.Click(uINewGroupMenuItem);
                _logMessage = string.Concat("Clicked on " + _submenuitem);
                _methodStatus = _pass;

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _MenuItem);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void WinMenuItemClickByClassNameTillExist(UITestControl Parent, string _ClassName, string _MenuItem, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WinMenuItemWindow = new WinWindow();
                WinMenuItemWindow.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _ClassName, "ControlType", "Window");
                if (_MenuItem.Contains("_"))
                {
                    _MenuItem = _MenuItem.Replace("_", " ");
                }

                WinMenuItem MenuItem = new WinMenuItem(WinMenuItemWindow);
                MenuItem.SearchProperties.Add(WinMenuItem.PropertyNames.Name, _MenuItem);
                MenuItem.WaitForControlReady();
                //newly added because of Send an IM Click
                    while (MenuItem.Exists)
                   {
                Mouse.Click(MenuItem);
                     }
                _logMessage = string.Concat("Clicked on " + _MenuItem);
                _methodStatus = _pass;

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _MenuItem);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinMenuItemByClassNameValidationChecked(UITestControl Parent, string _ClassName, string _MenuItem, string _checked, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WinMenuItemWindow = new WinWindow();
                WinMenuItemWindow.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _ClassName, "ControlType", "Window");
                if (_MenuItem.Contains("_"))
                {
                    _MenuItem = _MenuItem.Replace("_", " ");
                }

                WinMenuItem MenuItem = new WinMenuItem(WinMenuItemWindow);
                MenuItem.SearchProperties.Add(WinMenuItem.PropertyNames.Name, _MenuItem);
                // MenuItem.WaitForControlReady();
                object check;
                check = MenuItem.GetProperty(WinMenuItem.PropertyNames.Checked);
                string checkedvalue = check.ToString();
                if(checkedvalue == _checked)
                 {
                    _logMessage = string.Concat("Checked: " + _MenuItem);
                    // to uncheck
                    Mouse.Click(MenuItem);
                    _methodStatus = _pass;
                }
                else
                {
                    _logMessage = string.Concat("UnChecked: " + _MenuItem);
                    _methodStatus = _fail;

                }
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _MenuItem);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void WinMenuItemGetName(UITestControl Parent, string _ClassName, string _MenuItem, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {


                WinWindow WinMenuItemWindow = new WinWindow();
                WinMenuItemWindow.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _ClassName, "ControlType", "Window");
                if (_MenuItem.Contains("_"))
                {
                    _MenuItem = _MenuItem.Replace("_", " ");
                }

                WinMenuItem MenuItem = new WinMenuItem(WinMenuItemWindow);
                MenuItem.SearchProperties.Add(WinMenuItem.PropertyNames.Name, _MenuItem);
                MenuItem.WaitForControlReady();
                 Mouse.Hover(MenuItem);
                string Name = MenuItem.GetProperty("FriendlyName").ToString();
                Assert.AreEqual(_MenuItem, Name);
                _logMessage = string.Concat("Found the " + _callerName + " Name " + Name);
                _methodStatus = _pass;


                //        //WinWindow WinMenuItemWindow = new WinWindow(Parent);
                //        //WinMenuItemWindow.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _ClassName, "ControlType", "Window");

                //        WinClient client = new WinClient(Parent);
                //        client.SearchProperties.Add(WinControl.PropertyNames.ControlType, "Client");

                //        WinControl MenuItem = new WinControl(client);
                //        MenuItem.SearchProperties.Add(WinControl.PropertyNames.ControlType, "MenuButton");
                //        MenuItem.SearchProperties.Add(WinControl.PropertyNames.ClassName, _ClassName);
                //        MenuItem.SearchProperties.Add(WinControl.PropertyNames.Name, "Available, Busy, Do Not Distrub, Be Right Back, Off Work,Away,Offline", PropertyExpressionOperator.Contains);

                //        string Name = MenuItem.GetProperty("Name").ToString();
                //        _logMessage = string.Concat("Present Status is " + Name);
                //        _methodStatus = _pass;

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Menu Item " + _callerName + " is not found");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }


        public static string VerifyOwnPresnce(UITestControl Parent, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                string[] status = { "Available, Change My Status", "Busy, Change My Status", "Do not disturb, Change My Status", "Be right back, Change My Status", "Off work, Change My Status", "Away, Change My Status", "Offline, Change My Status", "In a call, Change My Status","In a conference call, Change My Status", "In a meeting, Change My Status", "Presenting, Change My Status" };
                bool success = false;
                int count = 0;
                string Name = string.Empty;
                string[] Namesplit = null;
                do
                {
                    WinControl WinMenuButtonClickByName = new WinControl(Parent);
                    WinMenuButtonClickByName.SearchProperties.Add(WinControl.PropertyNames.ControlType, "MenuButton");
                    WinMenuButtonClickByName.SearchProperties.Add(WinControl.PropertyNames.Name, status[count], PropertyExpressionOperator.Contains);

                    if (WinMenuButtonClickByName.Exists)
                    {
                      //  WinMenuButtonClickByName.DrawHighlight();
                        Mouse.Hover(WinMenuButtonClickByName);
                         Name = WinMenuButtonClickByName.GetProperty("Name").ToString();
                       Namesplit = Regex.Split(Name, ", Change My Status");
                        success = true;
                    }
                  
                    count++;
                } while (count < 11 & success == false);
                _logMessage = string.Concat("Present Status is " + Namesplit[0]);
            
                _methodStatus = _pass;
                return Name.ToString();

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to find the Present Status ");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void CompareStrings(string signin_Status,string aftSign_in, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
               
                if(signin_Status.Contains(aftSign_in))
                    _methodStatus = _pass;
                _logMessage = string.Concat("Validation : The " + _callerName + " is same");

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Validation : The " + _callerName + " is different");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void CompareNotEqualStrings(string firstvalue, string secondvalue, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {

                if (!firstvalue.Equals(secondvalue))
                    _methodStatus = _pass;
                _logMessage = string.Concat("Validation : The " + _callerName + " is different");

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Validation : The " + _callerName + " is same");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static object WinMenuItemByClassNameCheckCheckedStatus(UITestControl Parent, string _ClassName, string _MenuItem, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WinMenuItemWindow = new WinWindow();
                WinMenuItemWindow.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _ClassName, "ControlType", "Window");
                if (_MenuItem.Contains("_"))
                {
                    _MenuItem = _MenuItem.Replace("_", " ");
                }

                WinMenuItem MenuItem = new WinMenuItem(WinMenuItemWindow);
                MenuItem.SearchProperties.Add(WinMenuItem.PropertyNames.Name, _MenuItem);
                // MenuItem.WaitForControlReady();
                object check;
                check = MenuItem.GetProperty(WinMenuItem.PropertyNames.Checked);
                _methodStatus = _pass;
                _logMessage = string.Concat("check box status of "+ _MenuItem+ " is " + check);
                return check;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to check the check box status " + _MenuItem);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinDialogClickByName(UITestControl Parent, string _DialogName, string _ButtonName, int _iteration,[CallerMemberName] string _callerName = null)
        {
            try
            {
                //WinWindow Skype_for_Business = new WinWindow();
                //Skype_for_Business.SearchProperties.Add(WinControl.PropertyNames.ClassName, _WindowClassName);
                //Skype_for_Business.SearchProperties.Add(WinControl.PropertyNames.ControlType, "Window");
                    
                WinControl SkypeBusiness = new WinControl(Parent);
                SkypeBusiness.SearchProperties.Add("ControlType", "Dialog");
                SkypeBusiness.SearchProperties.Add(WinControl.PropertyNames.Name, _DialogName);
                SkypeBusiness.WaitForControlReady();
               // SkypeBusiness.DrawHighlight();

                WinButton winClick = new WinButton(SkypeBusiness);
                winClick.SearchProperties.Add(WinButton.PropertyNames.Name, _ButtonName);
                winClick.WaitForControlReady();

              //  winClick.DrawHighlight();
              //  while (winClick.Exists)
             //   {
                    Mouse.Click(winClick);
                    _logMessage = string.Concat("Clicked on " + _ButtonName);
                    _methodStatus = _pass;
              //  }

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _ButtonName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(),_iteration.ToString()));
            }
        }
        public static void WinDialogClickByNameTillExist(UITestControl Parent, string _DialogName, string _ButtonName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                //WinWindow Skype_for_Business = new WinWindow();
                //Skype_for_Business.SearchProperties.Add(WinControl.PropertyNames.ClassName, _WindowClassName);
                //Skype_for_Business.SearchProperties.Add(WinControl.PropertyNames.ControlType, "Window");

                WinControl SkypeBusiness = new WinControl(Parent);
                SkypeBusiness.SearchProperties.Add("ControlType", "Dialog");
                SkypeBusiness.SearchProperties.Add(WinControl.PropertyNames.Name, _DialogName);
                SkypeBusiness.WaitForControlReady();
            //    SkypeBusiness.DrawHighlight();

                WinButton winClick = new WinButton(SkypeBusiness);
                winClick.SearchProperties.Add(WinButton.PropertyNames.Name, _ButtonName);
                winClick.WaitForControlReady();

            //    winClick.DrawHighlight();
                while (winClick.Exists)
                {
                    Mouse.Click(winClick);
                _logMessage = string.Concat("Clicked on " + _ButtonName);
                _methodStatus = _pass;
                 }

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _ButtonName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinDialogClickByControlTypeAndButtonName(UITestControl Parent, string _ButtonName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinControl SkypeBusiness = new WinControl(Parent);
                SkypeBusiness.SearchProperties.Add("ControlType", "Dialog");
                SkypeBusiness.WaitForControlReady();
           //     SkypeBusiness.DrawHighlight();

                WinButton winClick = new WinButton(SkypeBusiness);
                winClick.SearchProperties.Add(WinButton.PropertyNames.Name, _ButtonName);
                winClick.SearchProperties.Add("ControlType", "Button");
                winClick.WaitForControlReady();

              //  winClick.DrawHighlight();
              
                Mouse.Click(winClick);
                _logMessage = string.Concat("Clicked on " + _ButtonName);
                _methodStatus = _pass;
               

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _ButtonName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void WinDialogClickByNameWindowTitle(UITestControl Parent, string _WindowTitle, string _DialogName, string _ButtonName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
        
                WinControl SkypeBusiness = new WinControl(Parent);
                SkypeBusiness.SearchProperties.Add("ControlType", "Dialog");
                SkypeBusiness.SearchProperties.Add(WinControl.PropertyNames.Name, _DialogName);
                SkypeBusiness.WindowTitles.Add(_WindowTitle);
                SkypeBusiness.WaitForControlReady();
         //     SkypeBusiness.DrawHighlight();

                WinButton winClick = new WinButton(SkypeBusiness);
                winClick.SearchProperties.Add(WinButton.PropertyNames.Name, _ButtonName);
                winClick.WaitForControlReady();
                winClick.WindowTitles.Add(_WindowTitle);
        //        winClick.DrawHighlight();
                while (winClick.Exists)
                {
                    Mouse.Click(winClick);
                    _logMessage = string.Concat("Clicked on " + _ButtonName);
                    _methodStatus = _pass;
                }

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _ButtonName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinClickByControlType(UITestControl Parent, string _ControlType, string _ButtonName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
               
                WinControl SkypeBusiness = new WinControl(Parent);
                SkypeBusiness.SearchProperties.Add("ControlType", _ControlType);


                WinButton winClick = new WinButton(SkypeBusiness);
                winClick.SearchProperties.Add(WinButton.PropertyNames.Name, _ButtonName);
                winClick.WaitForControlReady();
                // changes made because of SendFile
            //    while (winClick.Exists)
             //   {
                    Mouse.Click(winClick);
                    _logMessage = string.Concat("Clicked on " + _ButtonName);
                    _methodStatus = _pass;
             //   }

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _ButtonName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinClickByControlsControlTypeNameAndButtonName(UITestControl Parent, string _Name,string _ControlType, string _ButtonName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {

                WinControl winControl = new WinControl(Parent);
                winControl.SearchProperties.Add(WinControl.PropertyNames.Name, _Name);
                winControl.SearchProperties.Add("ControlType", _ControlType);


                WinButton winButton = new WinButton(winControl);
                winButton.SearchProperties.Add(WinButton.PropertyNames.Name, _ButtonName);
                winButton.WaitForControlReady();
                winButton.SetFocus();
    
          
                Mouse.Click(winButton);
                _logMessage = string.Concat("Clicked on " + _ButtonName);
                _methodStatus = _pass;
                

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _ButtonName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinDialogEditBoxClickByName(UITestControl Parent, string _DialogName, string _ButtonName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {  
                WinControl SkypeBusiness = new WinControl(Parent);
                SkypeBusiness.SearchProperties.Add("ControlType", "Dialog");
                SkypeBusiness.SearchProperties.Add(WinControl.PropertyNames.Name, _DialogName);

                WinEdit winClick = new WinEdit(SkypeBusiness);
                winClick.SearchProperties.Add(WinEdit.PropertyNames.Name, _ButtonName);
                winClick.SearchProperties.Add("ControlType", "Edit");
              //  winClick.DrawHighlight();
                winClick.WaitForControlReady();
                 Mouse.Click(winClick);

                _logMessage = string.Concat("Clicked on " + _ButtonName);
                _methodStatus = _pass;
                

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _ButtonName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinDialogDoubleClickByName(string _WindowClassName, string _DialogName, string _ButtonName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow Skype_for_Business = new WinWindow();
                Skype_for_Business.SearchProperties.Add(WinControl.PropertyNames.ClassName, _WindowClassName);
                Skype_for_Business.SearchProperties.Add(WinControl.PropertyNames.ControlType, "Window");

                WinControl SkypeBusiness = new WinControl(Skype_for_Business);
                SkypeBusiness.SearchProperties.Add("ControlType", "Dialog");
                SkypeBusiness.SearchProperties.Add(WinControl.PropertyNames.Name, _DialogName);

                WinButton winClick = new WinButton(SkypeBusiness);
                winClick.SearchProperties.Add(WinButton.PropertyNames.Name, _ButtonName);

                while (winClick.Exists)
                {
                    Mouse.Click(winClick);
                    _logMessage = string.Concat("Clicked on " + _ButtonName);
                    _methodStatus = _pass;
                }

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _ButtonName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static string GetAlphanumericValueWithBracketsAndDecimalAndHypen(string word,int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                char[] chars = new char[word.Length];
                int myindex = 0;
                for (int i = 0; i < word.Length; i++)
                {
                    char c = word[i];

                    if ((int)c >= 65 && (int)c <= 90)
                    {
                        chars[myindex] = c;
                        myindex++;
                    }
                    else if ((int)c >= 48 && (int)c <= 57)
                    {
                        chars[myindex] = c;
                        myindex++;
                    }
                    else if ((int)c >= 97 && (int)c <= 122)
                    {
                        chars[myindex] = c;
                        myindex++;
                    }
                    else if ((int)c == 40 || (int)c == 41 || (int)c == 45 || (int)c == 46 || (int)c == 32)
                    {
                        chars[myindex] = c;
                        myindex++;
                    }
                    else if ((int)c == 44)
                    {
                        chars[myindex] = c;
                        myindex++;
                    }
                }

                word = new string(chars);
                return word;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to identify the text " + word);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinDialogTextAssertByName(UITestControl Parent, string _DialogName, string _TextName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinControl WinDialog = new WinControl(Parent);
                WinDialog.SearchProperties.Add("ControlType", "Dialog");
                WinDialog.SearchProperties.Add(WinControl.PropertyNames.Name, _DialogName);

                WinText winText = new WinText(WinDialog);
                winText.SearchProperties.Add("ControlType", "Text");
                //    winText.SearchProperties.Add(WinButton.PropertyNames.Name, _TextName, PropertyExpressionOperator.Contains);
                string Name = winText.GetProperty("Name").ToString();
                if (Name.Contains(_TextName))
                {
                    _logMessage = string.Concat("The " + _callerName + " text :" + _TextName);
                    _methodStatus = _pass;
                }
                else
                {

                    string name = GetAlphanumericValueWithBracketsAndDecimalAndHypen(Name, _iteration);
                    string textname = GetAlphanumericValueWithBracketsAndDecimalAndHypen(_TextName, _iteration);
                    string Newname = name.Trim('\0');
                    string NewTextname = textname.Trim('\0');
                    Assert.AreEqual(Newname.TrimEnd(), NewTextname.TrimEnd());
                }
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to identify the text " + _TextName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }


        public static void WinDropDownButtonByName(UITestControl Parent, string _Name, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinControl WinDropDownButtonByName = new WinControl(Parent);
                WinDropDownButtonByName.SearchProperties.Add(WinControl.PropertyNames.Name, _Name, "ControlType", "DropDownButton");
                WinDropDownButtonByName.WaitForControlReady();
                Mouse.Click(WinDropDownButtonByName);
                _logMessage = string.Concat("Clicked on " + _Name);
                //_winClick.WaitForControlReady();

                _methodStatus = _pass;
 
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Process failed to click on " + _Name);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static string WinDropDownButtonGetChildName(UITestControl Parent, string _Name, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinControl WinDropDownButtonByName = new WinControl(Parent);
                WinDropDownButtonByName.SearchProperties.Add(WinControl.PropertyNames.Name, _Name, "ControlType", "DropDownButton");

                UITestControlCollection uic = WinDropDownButtonByName.GetChildren();
                Object name = null;
                foreach (UITestControl ui in uic)
                {
                    name = ui.GetProperty("Name");
                    _logMessage = string.Concat("Present Status is " + name);
                    _methodStatus = _pass;
                    break;
                }
                return name.ToString();

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to find the present status");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static string WinDropDownButtonGetChildNameValidation(UITestControl Parent, string _Name,string status, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinControl WinDropDownButtonByName = new WinControl(Parent);
                WinDropDownButtonByName.SearchProperties.Add(WinControl.PropertyNames.Name, _Name, "ControlType", "DropDownButton");

                UITestControlCollection uic = WinDropDownButtonByName.GetChildren();
                Object name = null;
                foreach (UITestControl ui in uic)
                {
                    name = ui.GetProperty("Name");
                    if (name.ToString() == status)
                    {
                        _logMessage = string.Concat("Present Status is" + name +" and compared succesfully");
                        _methodStatus = _pass;
                        break;
                    }
                    else
                    {
                        _logMessage = string.Concat("Present Status is" + name + " and compared failed");
                        _methodStatus = _fail;
                        Assert.Fail("Present Status is" + name + " and compared failed");
                        break;
                    }
                }
                return name.ToString();

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to compare the present status");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinDropDownButtonClickOnChildName(UITestControl Parent, string _Name, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinControl WinDropDownButtonByName = new WinControl(Parent);
                WinDropDownButtonByName.SearchProperties.Add(WinControl.PropertyNames.Name, _Name, "ControlType", "DropDownButton");

                UITestControlCollection uic = WinDropDownButtonByName.GetChildren();
                Object name = null;
                foreach (UITestControl ui in uic)
                {
                    name = ui.GetProperty("Name");
                    Mouse.Click(ui);
                    _logMessage = string.Concat("Clicked on " + name);
                    _methodStatus = _pass;
                    break;
                }
                

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on status");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }


        public static void WinMenuItemClickByName(string _className,string _Name, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow window = new WinWindow();
                window.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _className, "ControlType", "Window");

                WinClient client = new WinClient(window);
                client.SearchProperties.Add(WinControl.PropertyNames.ControlType, "Client");

                WinGroup group = new WinGroup(client);
                group.SearchProperties.Add(WinControl.PropertyNames.ControlType, "Group");

                WinControl WinMenuItemClickByName = new WinControl(group);
                WinMenuItemClickByName.SearchProperties.Add(WinControl.PropertyNames.Name, _Name, "ControlType", "MenuItem");
                WinMenuItemClickByName.WaitForControlExist();
                WinMenuItemClickByName.WaitForControlReady();
                Mouse.Click(WinMenuItemClickByName);

                System.Threading.Thread.Sleep(1000);
                _logMessage = string.Concat("Clicked on " + _Name);
                //_winClick.WaitForControlReady();

                _methodStatus = _pass;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Process failed to click on " + _Name);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

       
        public static void WinMenuItemClickByNameAndParent(UITestControl Parent, string _Name, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
 
                WinControl WinMenuItemClickByName = new WinControl(Parent);
                WinMenuItemClickByName.SearchProperties.Add(WinControl.PropertyNames.Name, _Name, "ControlType", "MenuItem");
                WinMenuItemClickByName.WaitForControlExist();
                WinMenuItemClickByName.WaitForControlReady();
                Mouse.Click(WinMenuItemClickByName);
           
                System.Threading.Thread.Sleep(1000);
                _logMessage = string.Concat("Clicked on " + _Name);
                //_winClick.WaitForControlReady();

                _methodStatus = _pass;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Process failed to click on " + _Name);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static bool WinGroupButtonClickByName(UITestControl Parent, string _GroupName, string _ButtonName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            bool success = false;
            try
            {
             

                WinGroup group = new WinGroup(Parent);
                group.SearchProperties.Add(WinControl.PropertyNames.Name, "Group");
             //   group.WindowTitles.Add(_Title);

                WinButton winButton = new WinButton(group);
                winButton.SearchProperties.Add(WinControl.PropertyNames.Name, _ButtonName, "ControlType", "Button");

                if (winButton.Exists)
                { //    winButton.WindowTitles.Add(_Title);
                    winButton.WaitForControlReady();
                    Mouse.Click(winButton);

                    System.Threading.Thread.Sleep(1000);
                    _logMessage = string.Concat("Clicked on " + _ButtonName);
                    //_winClick.WaitForControlReady();

                    _methodStatus = _pass;
                    success = true;
                }
                else
                {
                    _logMessage = string.Concat("Button not found" + _ButtonName);
                    //_winClick.WaitForControlReady();

                    _methodStatus = _pass;
                    success = false;
                }
                return success;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _ButtonName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static bool WinGroupButtonRightClickByName(UITestControl Parent, string _GroupName, string _ButtonName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            bool success = false;
            try
            {


                WinGroup group = new WinGroup(Parent);
                group.SearchProperties.Add(WinControl.PropertyNames.Name, "Group");
                //   group.WindowTitles.Add(_Title);

                WinButton winButton = new WinButton(group);
                winButton.SearchProperties.Add(WinControl.PropertyNames.Name, _ButtonName, "ControlType", "Button");

                if (winButton.Exists)
                { //    winButton.WindowTitles.Add(_Title);
                    winButton.WaitForControlReady();
                    Mouse.Click(winButton,System.Windows.Forms.MouseButtons.Right);

                    System.Threading.Thread.Sleep(1000);
                    _logMessage = string.Concat("Right Clicked on " + _ButtonName);
                    //_winClick.WaitForControlReady();

                    _methodStatus = _pass;
                    success = true;
                }
                else
                {
                    _logMessage = string.Concat("Button not found" + _ButtonName);
                    //_winClick.WaitForControlReady();

                    _methodStatus = _fail;
                    success = false;
                }
                return success;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to right click on " + _ButtonName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinNestedTreeItemButtonClickByName(UITestControl Parent, string _TreeItemName, string _NestedTreeItemName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {


                WinTreeItem winTreeItem = new WinTreeItem(Parent);
                winTreeItem.SearchProperties.Add(WinTreeItem.PropertyNames.Name, _TreeItemName);
                winTreeItem.SearchProperties.Add(WinTreeItem.PropertyNames.ControlType, "TreeItem");

                WinTreeItem winTreeItemSkype = new WinTreeItem(winTreeItem);
                winTreeItemSkype.SearchProperties.Add(WinTreeItem.PropertyNames.Name, _NestedTreeItemName);
                winTreeItemSkype.WaitForControlReady();
                Mouse.Click(winTreeItemSkype);
                _logMessage = string.Concat("Clicked on " + _NestedTreeItemName);

                _methodStatus = _pass;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Process failed to click on " + _NestedTreeItemName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }


        public static void WinTextClickByName(UITestControl Parent, string _Name, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinText WinTextByName = new WinText(Parent);
                WinTextByName.SearchProperties.Add(WinText.PropertyNames.Name, _Name, "ControlType", "Text");
                Mouse.Click(WinTextByName);
                //if (_MenuItem.Contains("_"))
                //{

                //    _MenuItem = _MenuItem.Replace("_", " ");
                //}
                //WinMenuItem MenuItem = new WinMenuItem(WinDropDownButtonByName);
                //MenuItem.SearchProperties.Add(WinMenuItem.PropertyNames.Name, _MenuItem);
                //Mouse.Click(MenuItem);
                _logMessage = string.Concat("Clicked on " + _Name);
                //_winClick.WaitForControlReady();

                _methodStatus = _pass;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Process failed to click on " + _Name);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void WinTextAssertByName(UITestControl Parent, string _Name, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinText WinTextByName = new WinText(Parent);
                WinTextByName.SearchProperties.Add(WinText.PropertyNames.Name, _Name, PropertyExpressionOperator.Contains);
                string Name = WinTextByName.GetProperty("Name").ToString();
                Assert.AreEqual(Name, _Name);
              
                _logMessage = string.Concat("Version is same: " + _Name);
                //_winClick.WaitForControlReady();

                _methodStatus = _pass;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Process failed to click on " + _Name);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String ,String> (System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
    
        public static void KeyBoardText(string _keyword, int iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                Keyboard.SendKeys(_keyword + "{Enter}");
                _logMessage = string.Concat("The value : " + _keyword + " is entered" + _callerName);
                _methodStatus = _pass;
            }
            catch(Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to enter the value");
                 throw;
            }
            finally
           {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), iteration.ToString()));
            }
            }
        
        public static void ClickOnWindowCloseButton(UITestControl Parent, string _name, int _iteration, [CallerMemberName] string _callerName = null)
        {                                                  
            try
            {
             
                WinTitleBar titleBar = new WinTitleBar(Parent);
                titleBar.SearchProperties.Add("ControlType", "TitleBar");
                WinButton button = new WinButton(titleBar);
                button.SearchProperties.Add("Name", _name);
                Mouse.Click(button);
                _logMessage = string.Concat("Clicked on close button in the " + _callerName + " window");
                _methodStatus = _pass;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click close button in window");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String,String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, globalFunctions._logMessage, globalFunctions._methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static bool ClickOnWindowButtonIfExist(UITestControl Parent, string _name, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                bool exist = false;
                Parent.SetFocus();
                WinTitleBar titleBar = new WinTitleBar(Parent);
                titleBar.SearchProperties.Add("ControlType", "TitleBar");
                //if (titleBar.Exists)
                //{
                    WinButton button = new WinButton(titleBar);
                    button.SearchProperties.Add("Name", _name);
                    if (button.Exists)
                    {
                        //button.DrawHighlight();
                        Mouse.Click(button);
                        exist = true;
                        _logMessage = string.Concat("Clicked on Maximize button in the " + _callerName + " window");
                        _methodStatus = _pass;
                    }
                //}
                return exist;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click maximize button in window");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, globalFunctions._logMessage, globalFunctions._methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void ClickOnComboBox(UITestControl Parent, string _name, string title, string filename,int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {

                WinComboBox winComboBox = new WinComboBox(Parent);
                winComboBox.SearchProperties.Add(WinComboBox.PropertyNames.Name, _name);
                winComboBox.WindowTitles.Add(title);
                winComboBox.WaitForControlReady();
                Mouse.Click(winComboBox);
                _logMessage = string.Concat("Clicked on the combo box: " + _name);
                Keyboard.SendKeys(filename);
                _logMessage = string.Concat("Value entered in the combo box is " + filename);
                _methodStatus = _pass;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on the combo box: " + _name);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, globalFunctions._logMessage, globalFunctions._methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void GetComboBoxSelecetedItem(UITestControl Parent, string _classname, string title, string value, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {

                WinComboBox winComboBox = new WinComboBox(Parent);
                winComboBox.SearchProperties.Add(WinComboBox.PropertyNames.ClassName, _classname);
                winComboBox.WindowTitles.Add(title);
               // winComboBox.DrawHighlight();
                object selecteditem = winComboBox.GetProperty("SelectedItem");
                if (selecteditem.ToString() == value)
                {
                    _logMessage = string.Concat("Validation : Value present in the combo box is " + value);
                    _methodStatus = _pass;
                }
                else
                {
                    _logMessage = string.Concat("Validation : Value present in the combo box is " + selecteditem.ToString());
                    _methodStatus = _fail;
                }
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Validation : The combo box is empty");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, globalFunctions._logMessage, globalFunctions._methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static bool HoverOnToolBarButtonOrMenuButton(UITestControl Parent, string _toolbarname, string _buttonname, string _menubuttonname, int _iteration, [CallerMemberName] string _callerName = null)
        {
            bool exist = false;
            try
            {

                WinToolBar winToolBar = new WinToolBar(Parent);
                winToolBar.SearchProperties.Add("Name", _toolbarname);
                winToolBar.WindowTitles.Add(_toolbarname);
                if (winToolBar.Exists)
                {
                    WinButton winButton = new WinButton(winToolBar);
                    winButton.SearchProperties.Add("Name", _buttonname, PropertyExpressionOperator.Contains);
                    winButton.WindowTitles.Add(_toolbarname);

                    if (winButton.Exists)
                    {
                        exist = true;
                       // Mouse.Hover(winButton);
                    }
                    else
                    {

                        WinMenu winMenuButton = new WinMenu(winToolBar);
                        winMenuButton.SearchProperties.Add("Name", _menubuttonname, PropertyExpressionOperator.Contains);
                        winMenuButton.SearchProperties.Add("ControlType", "MenuButton");
                        winMenuButton.WindowTitles.Add(_toolbarname);
                        UITestControlCollection ui = winMenuButton.FindMatchingControls();
                        foreach (UITestControl u in ui)
                        {
                          //  Mouse.Hover(winMenuButton);
                            break;
                        }
                    }
                }
                _logMessage = string.Concat("Hovered on the button: " + _buttonname + " in the tool bar");
                _methodStatus = _pass;
                return exist;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to hover on the button: " + _buttonname);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, globalFunctions._logMessage, globalFunctions._methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }


        public static void ClickOnToolBarButton(UITestControl Parent, string _toolbarname, string _buttonname, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {

                WinToolBar winToolBar = new WinToolBar(Parent);
                winToolBar.SearchProperties.Add("Name", _toolbarname);

                WinButton winButton = new WinButton(winToolBar);
                winButton.SearchProperties.Add("Name", _buttonname, PropertyExpressionOperator.Contains);
               // winButton.DrawHighlight();
                winButton.WaitForControlReady();
                if (winButton.Exists)
                {
                    Mouse.DoubleClick(winButton);
                }
                _logMessage = string.Concat("Clicked on the button: " + _buttonname + " in the tool bar");
                _methodStatus = _pass;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on the button: " + _buttonname);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, globalFunctions._logMessage, globalFunctions._methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void ClickOnToolBarAndPassValues(UITestControl Parent, string _toolbarname, string _title, string value, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {

                WinToolBar winToolBar = new WinToolBar(Parent);
                winToolBar.SearchProperties.Add("Name", _toolbarname,PropertyExpressionOperator.Contains);
                winToolBar.WindowTitles.Add(_title);
                winToolBar.WaitForControlReady();
                winToolBar.SetFocus();
                Mouse.Click(winToolBar);
                
                _logMessage = string.Concat("Clicked on : " + _toolbarname + " in the tool bar");
                Keyboard.SendKeys(value);
                Keyboard.SendKeys("{ENTER}");
                _logMessage = string.Concat("Value entered is " +value);

                _methodStatus = _pass;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on : " + _toolbarname);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, globalFunctions._logMessage, globalFunctions._methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        //public static void ClickOnWebPageOk(string _name, [CallerMemberName] string _callerName = null)
        //{
        //    try
        //    {
        //        WinWindow window = new WinWindow();
        //        window.SearchProperties.Add("Name", "Windows Internet Explorer", "ClassName", "#32770");
        //        window.DrawHighlight();
        //        WinWindow window2 = new WinWindow(window);
        //        window.SearchProperties.Add("ControlId", "1", "ControlType", "Window");
        //        window.DrawHighlight();
        //        WinButton button = new WinButton(window2);
        //        button.SearchProperties.Add("Name", _name);
        //        //button.DrawHighlight();
        //        button.WaitForControlReady();
        //        Mouse.Click(button);
        //        _logMessage = string.Concat("Process clicked on "+_name +" button in pop up");
        //        _methodStatus = _pass;
        //    }
        //    catch (Exception)
        //    {
        //        _methodStatus = _fail;
        //        _logMessage = string.Concat("Process failed to click " + _name + "  button from pop up");
        //        throw;
        //    }
        //    finally
        //    {
        //        listOfTuples.Add(new Tuple<String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, globalFunctions._logMessage, globalFunctions._methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime()));
        //    }
        //}
        //public static string GetBrowserURL(UITestControl _parent, [CallerMemberName] string _callerName = null)
        //{
        //    try
        //    {
        //        WinEdit uIHtmlHyperLinkObject = new WinEdit(_parent);
        //        uIHtmlHyperLinkObject.SearchProperties.Add("ControlType", "Edit");
        //        uIHtmlHyperLinkObject.SearchProperties.Add("Name", "Address and search using Google");
        //        // uIHtmlHyperLinkObject.DrawHighlight();
        //        uIHtmlHyperLinkObject.WaitForControlReady();
        //        string URL = uIHtmlHyperLinkObject.GetProperty("Text").ToString();
        //        _logMessage = string.Concat("Url:"+URL );
        //        _methodStatus = _pass;
        //        return URL;
        //    }
        //    catch (Exception)
        //    {
        //        _methodStatus = _fail;
        //        _logMessage = string.Concat("Process failed to get the URL of current browser window/Tab");
        //        throw;
        //    }
        //    finally
        //    {
        //        listOfTuples.Add(new Tuple<String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, globalFunctions._logMessage, globalFunctions._methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime()));
        //    }
        //}
        //public static void AssertURL( string _expectedValue, string _actualValue, [CallerMemberName] string _callerName = null)
        //{
        //    try
        //    { 
        //        StringAssert.Contains(_actualValue,_expectedValue);
        //        _logMessage = string.Concat("Validation::::Expected: " + _expectedValue + "::::Actual: " + _actualValue);
        //        _methodStatus = _pass;
        //    }
        //    catch (Exception)
        //    {
        //        _methodStatus = _fail;
        //        _logMessage = string.Concat("Validation failed");
        //        throw;
        //    }
        //    finally
        //    {
        //        listOfTuples.Add(new Tuple<String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime()));
        //    }
        //}
        //public static void AssertionToCheckForAvailabilityOFLink_Id(UITestControl _parent, string _id,string _value, [CallerMemberName] string _callerName = null)
        //{
        //    try
        //    {
        //          HtmlHyperlink _uIHyperlinkControl = new HtmlHyperlink(_parent);
        //         _uIHyperlinkControl.SearchProperties.Add(HtmlHyperlink.PropertyNames.Id, _id);
        //         _comparisonValue = _uIHyperlinkControl.GetProperty("InnerText").ToString();
        //          if(_uIHyperlinkControl.Exists)
        //               {
        //                    _logMessage = string.Concat("Validation::::Expected: " + _value + "::::Actual: " + _comparisonValue);
        //                    _methodStatus = _pass;
        //               }          
        //          else
        //               {
        //            Assert.IsFalse(true);
        //                }    
        //    }
        //    catch (Exception)
        //    {
        //         _methodStatus = _fail;
        //         _logMessage = string.Concat("Validation::::Expected: " + _value + "::::Actual: Error(Unable to find the control and validate)" );
        //         throw;
        //    }
        //    finally
        //    {
        //        listOfTuples.Add(new Tuple<String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime()));
        //    }
        //}  
        //public static void AssertionToCheckForAvailabilityOFLink_InnerText(UITestControl _parent, string _innerText, string _value, [CallerMemberName] string _callerName = null)
        //{
        //    try
        //    {
        //        HtmlHyperlink _uIHyperlinkControl = new HtmlHyperlink(_parent);
        //        _uIHyperlinkControl.SearchProperties.Add(HtmlHyperlink.PropertyNames.InnerText, _innerText);
        //        _comparisonValue = _uIHyperlinkControl.GetProperty("InnerText").ToString();
        //        if (_uIHyperlinkControl.Exists)
        //        {
        //            _logMessage = string.Concat("Validation::::Expected: " + _value + "::::Actual: " + _comparisonValue);
        //            _methodStatus = _pass;
        //        }
        //        else
        //        {
        //            Assert.IsFalse(true);
        //        }
        //    }
        //    catch (Exception)
        //    {
        //        _methodStatus = _fail;
        //        _logMessage = string.Concat("Validation::::Expected: " + _value + "::::Actual: Error(Unable to find the control and validate)");
        //        throw;
        //    }
        //    finally
        //    {
        //        listOfTuples.Add(new Tuple<String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime()));
        //    }
        //}
        //public static string GetSpanUIInnerTextById(UITestControl _parent, string _id, [CallerMemberName] string _callerName = null)
        //{
        //    try
        //    {
        //        HtmlSpan _uISpanControl = new HtmlSpan(_parent);
        //        _uISpanControl.SearchProperties.Add("Id", _id);
        //        //_uISpanControl.DrawHighlight();
        //        _uISpanControl.WaitForControlReady();
        //        string innerText=_uISpanControl.GetProperty("InnerText").ToString();
        //        _logMessage = string.Concat("getSpanInnerText:"+ innerText);
        //        _methodStatus = _pass;
        //        return innerText;
        //    }
        //    catch (Exception)
        //    {
        //        _methodStatus = _fail;
        //        _logMessage = string.Concat("Error#(Unable to get the innertext of specified control)");
        //        throw;
        //    }
        //    finally
        //    {
        //        listOfTuples.Add(new Tuple<String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime()));
        //    }
        //}
        //public static void AssertValues(string value1, string value2, [CallerMemberName] string _callerName = null)
        //{
        //    try
        //    {
        //      if(value2.Contains(value1))
        //        {
        //            _methodStatus = _pass;
        //            _logMessage = string.Concat("Validation::::Expected: " + value1 + "::::Actual:" + value2);
        //        }
        //    }
        //    catch (Exception)
        //    {
        //        _methodStatus = _fail;
        //        _logMessage = string.Concat("Validation::::Expected: " + value1 + "::::Actual: Error(Unable to find the control and validate)");
        //        throw;
        //    }
        //    finally
        //    {
        //        listOfTuples.Add(new Tuple<String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime()));
        //    }
        //}
        //public static string GetHtmlTableInnerText(UITestControl _parent, string _id, [CallerMemberName] string _callerName = null)
        //{
        //    try
        //    {
        //        HtmlTable table = new HtmlTable(_parent);
        //        table.SearchProperties.Add("Id",_id);
        //        table.DrawHighlight();
        //        string innerText = table.GetProperty("InnerText").ToString();
        //        _logMessage = string.Concat("getTableInnerText:" + innerText);
        //        _methodStatus = _pass;
        //        return innerText;
        //    }
        //    catch (Exception)
        //    {
        //        _methodStatus = _fail;
        //        _logMessage = string.Concat("Error#(Unable to get the innerText of table)");
        //        throw;
        //    }
        //    finally
        //    {
        //        listOfTuples.Add(new Tuple<String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime()));
        //    }
        //}
        //public static string GetEditBoxDefaultText(UITestControl _parent,string _id,[CallerMemberName] string _callerName = null)
        //{
        //    string url = null;
        //    try
        //    {
        //        HtmlEdit uIHtmlHyperLinkObject = new HtmlEdit(_parent);
        //        uIHtmlHyperLinkObject.SearchProperties.Add("Id", _id,PropertyExpressionOperator.Contains);
        //        if (uIHtmlHyperLinkObject.Exists)
        //        {
        //            url = uIHtmlHyperLinkObject.GetProperty("DefaultText").ToString();
        //            _logMessage = string.Concat("getEditBoxDefaultText: " + url);
        //            _methodStatus = _pass;
        //        }
        //        return url;
        //    }
        //    catch (Exception)
        //    {
        //        _methodStatus = _fail;
        //        _logMessage = string.Concat("Error#(Unable to get the DefaultText of uIHtmlHyperLinkObject box)");
        //        throw;
        //    }
        //    finally
        //    {
        //        listOfTuples.Add(new Tuple<String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime()));
        //    }
        //}
        //public static void CheckHyperlinkClassContainsDisabledById(UITestControl _parent,string _id,string value, [CallerMemberName] string _callerName = null)
        //{
        //    try
        //    {
        //        HtmlHyperlink uIHtmlHyperLinkObject = new HtmlHyperlink(_parent);
        //        uIHtmlHyperLinkObject.SearchProperties.Add("Id", _id);
        //        if (uIHtmlHyperLinkObject.GetProperty("Class").ToString().Contains(value))
        //        {                  
        //            _logMessage = string.Concat(uIHtmlHyperLinkObject.GetProperty("InnerText").ToString() + " control is disabled");
        //            _methodStatus = _pass;
        //        }
        //        else
        //        {
        //            Assert.IsTrue(false,"Control is not disabled");
        //        }
        //    }
        //    catch (Exception)
        //    {
        //        _methodStatus = _fail;
        //        _logMessage = string.Concat("Error#(Edit items control is not disabled)");
        //        throw;
        //    }
        //    finally
        //    {
        //        listOfTuples.Add(new Tuple<String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime()));
        //    }
        //}
        //public static void CheckSpanClassContainsDisabledById(UITestControl _parent, string _id, string value, [CallerMemberName] string _callerName = null)
        //{
        //    try
        //    {
        //        HtmlSpan uIHtmlSpanObject = new HtmlSpan(_parent);
        //        uIHtmlSpanObject.SearchProperties.Add("Id", _id);
        //        if (uIHtmlSpanObject.GetProperty("Class").ToString().Contains(value))
        //        {
        //            _logMessage = string.Concat(uIHtmlSpanObject.GetProperty("InnerText").ToString() + " control is disabled");
        //            _methodStatus = _pass;
        //        }
        //        else
        //        {
        //            Assert.IsTrue(false, "Control is not disabled");
        //        }
        //    }
        //    catch (Exception)
        //    {
        //        _methodStatus = _fail;
        //        _logMessage = string.Concat("Error#(Edit items control is not disabled)");
        //        throw;
        //    }
        //    finally
        //    {
        //        listOfTuples.Add(new Tuple<String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime()));
        //    }
        //}
        //public static string GetHrefFromHyperlinkByInnerText(UITestControl _parent, string _innerText, [CallerMemberName] string _callerName = null)
        //{
        //    string href = null;
        //    try
        //    {
        //        HtmlHyperlink uIHtmlHyperLinkObject = new HtmlHyperlink(_parent);
        //        uIHtmlHyperLinkObject.SearchProperties.Add("InnerText", _innerText, PropertyExpressionOperator.Contains);
        //        if (uIHtmlHyperLinkObject.Exists)
        //        {
        //            href = uIHtmlHyperLinkObject.GetProperty("Href").ToString();
        //            _logMessage = string.Concat("gethyperLinkHrefText: " + href);
        //            _methodStatus = _pass;
        //        }
        //        return href;
        //    }
        //    catch (Exception)
        //    {
        //        _methodStatus = _fail;
        //        _logMessage = string.Concat("Error#(Unable to get the hyperlink Href )");
        //        throw;
        //    }
        //    finally
        //    {
        //        listOfTuples.Add(new Tuple<String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime()));
        //    }
        //}
        //public static void CheckHyperLinkExistsByInnerText(UITestControl _parent,string _innerText, [CallerMemberName] string _callerName = null)
        //{
        //    try
        //    {
        //        HtmlHyperlink uIHtmlHyperLinkObject = new HtmlHyperlink(_parent);
        //        uIHtmlHyperLinkObject.SearchProperties.Add("InnerText", _innerText);
        //        if (uIHtmlHyperLinkObject.Exists)
        //        {
        //            _logMessage = string.Concat(_innerText + " control is available");
        //            _methodStatus = _pass;
        //        }
        //        else
        //        {
        //            Assert.IsTrue(false, _innerText + "Control is not available");
        //        }
        //    }
        //    catch (Exception)
        //    {
        //        _methodStatus = _fail;
        //        _logMessage = string.Concat(_innerText + " Control not available");
        //        throw;
        //    }
        //    finally
        //    {
        //        listOfTuples.Add(new Tuple<String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime()));
        //    }
        //}
        //public static void CheckHyperLinkExistsById(UITestControl _parent, string _id, [CallerMemberName] string _callerName = null)
        //{
        //    string innerTextValue = string.Empty;
        //    try
        //    {
        //        HtmlHyperlink uIHtmlHyperLinkObject = new HtmlHyperlink(_parent);
        //        uIHtmlHyperLinkObject.SearchProperties.Add("id", _id);
        //        if (uIHtmlHyperLinkObject.Exists)
        //        {
        //            innerTextValue = uIHtmlHyperLinkObject.GetProperty("InnerText").ToString();
        //            _logMessage = string.Concat(innerTextValue + " control is available");
        //            _methodStatus = _pass;
        //        }
        //        else
        //        {
        //            Assert.IsTrue(false, innerTextValue + "Control is not available");
        //        }
        //    }
        //    catch (Exception)
        //    {
        //        _methodStatus = _fail;
        //        _logMessage = string.Concat( "Hyperlink Control not available");
        //        throw;
        //    }
        //    finally
        //    {
        //        listOfTuples.Add(new Tuple<String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime()));
        //    }
        //}
        //public static void CheckHyperLinkExistsByDisplayeText(UITestControl _parent, string _displayTest, [CallerMemberName] string _callerName = null)
        //{
        //    string innerTextValue = string.Empty;
        //    try
        //    {
        //        HtmlHyperlink uIHtmlHyperLinkObject = new HtmlHyperlink(_parent);
        //        uIHtmlHyperLinkObject.SearchProperties.Add("DisplayText", _displayTest);
        //        if (uIHtmlHyperLinkObject.Exists)
        //        {
        //            innerTextValue = uIHtmlHyperLinkObject.GetProperty("InnerText").ToString();
        //            _logMessage = string.Concat(innerTextValue + " control is available");
        //            _methodStatus = _pass;
        //        }
        //        else
        //        {
        //            Assert.IsTrue(false, innerTextValue + "Control is not available");
        //        }
        //    }
        //    catch (Exception)
        //    {
        //        _methodStatus = _fail;
        //        _logMessage = string.Concat("Hyperlink Control not available");
        //        throw;
        //    }
        //    finally
        //    {
        //        listOfTuples.Add(new Tuple<String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime()));
        //    }
        //}
        //public static void CheckLabelExistsById(UITestControl _parent, string _id,string _labelName, [CallerMemberName] string _callerName = null)
        //{
        //    try
        //    {
        //        HtmlLabel label = new HtmlLabel(_parent);
        //        label.SearchProperties.Add("Id", _id);
        //        if (label.Exists)
        //        {
        //            _logMessage = string.Concat(_labelName + " label is available");
        //            _methodStatus = _pass;
        //        }
        //        else
        //        {
        //            Assert.IsTrue(false, _labelName + "label is not available");
        //        }
        //    }
        //    catch (Exception)
        //    {
        //        _methodStatus = _fail;
        //        _logMessage = string.Concat(_labelName + " label is not available");
        //        throw;
        //    }
        //    finally
        //    {
        //        listOfTuples.Add(new Tuple<String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime()));
        //    }
        //}
        //public static string GetHtmlDivInnerTextByClassAndInnerText(UITestControl _parent, string _class,string _innerText, [CallerMemberName] string _callerName = null)
        //{
        //    try
        //    {
        //        HtmlDiv div = new HtmlDiv(_parent);
        //        div.SearchProperties.Add("Class", _class,"InnerText",_innerText);
        //        div.DrawHighlight();
        //        string innerText = div.GetProperty("InnerText").ToString();
        //        _logMessage = string.Concat("getDivInnerText value is:" + innerText);
        //        _methodStatus = _pass;
        //        return innerText;
        //    }
        //    catch (Exception)
        //    {
        //        _methodStatus = _fail;
        //        _logMessage = string.Concat("Error#(Unable to get the innerText of pane control)");
        //        throw;
        //    }
        //    finally
        //    {
        //        listOfTuples.Add(new Tuple<String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime()));
        //    }
        //}


        public static void WinTabPageClickByName(UITestControl Parent,  string _TabPage, string _title, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
              

                WinTabPage TabPageItem = new WinTabPage(Parent);
                TabPageItem.SearchProperties.Add(WinMenuItem.PropertyNames.Name, _TabPage);
                TabPageItem.WindowTitles.Add(_title);

                Mouse.Click(TabPageItem);
                _logMessage = string.Concat("Clicked on " + _TabPage);
                //_winClick.WaitForControlReady();

                _methodStatus = _pass;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _TabPage);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinTabPageClickByName(UITestControl Parent, string _TabPage, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
           
                WinTabPage TabPageItem = new WinTabPage(Parent);
                TabPageItem.SearchProperties.Add(WinTabPage.PropertyNames.Name, _TabPage);

                if (TabPageItem.Exists)
                {
                    Mouse.Click(TabPageItem);;
                    _logMessage = string.Concat("Clicked on " + _TabPage);

                    _methodStatus = _pass;
                }
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _TabPage);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinImageClickByClassName(UITestControl Parent, string _ClassName, string _Image, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WinImageWindow = new WinWindow();
                WinImageWindow.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _ClassName, "ControlType", "Window");
                if (_Image.Contains("_"))
                {
                    _Image = _Image.Replace("_", " ");
                }

                UITestControl Image = new UITestControl(WinImageWindow);
                Image.SearchProperties.Add("ControlType", "Image");

                Mouse.Click(Image);

                _logMessage = string.Concat("Clicked on " + _Image);


                _methodStatus = _pass;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _Image);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(),_iteration.ToString()));
            }
        }

        public static void WinImageClickByParentDropDown(UITestControl Parent, string _windowname, string _dropdownname, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinControl WinDropDownButtonByName = new WinControl(Parent);
                WinDropDownButtonByName.SearchProperties.Add(WinControl.PropertyNames.Name, _dropdownname, "ControlType", "DropDownButton");
             //   WinDropDownButtonByName.WaitForControlReady();
                WinDropDownButtonByName.WindowTitles.Add(_windowname);
                WinDropDownButtonByName.WaitForControlReady();
              //  WinDropDownButtonByName.DrawHighlight();

                UITestControl Image = new UITestControl(WinDropDownButtonByName);
                Image.SearchProperties.Add("ControlType", "Image");
                Image.WindowTitles.Add(_windowname);
                //Mouse.Hover(Image);
                Image.WaitForControlReady();
               // Mouse.Click(WinDropDownButtonByName);
                UITestControlCollection uic = Image.FindMatchingControls();

                foreach (UITestControl ui in uic)

                {

                    if (ui.BoundingRectangle.Width > 0)

                    {

                        Mouse.Click(ui);
                        break;


                    }

                }
                //Mouse.Click(Image, new Point(Present_X, Present_Y));
                _logMessage = string.Concat("Clicked on " + _callerName);
                _methodStatus = _pass;

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinTitleBarLocation(UITestControl Parent, string _ClassName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WinImageWindow = new WinWindow(Parent);
                WinImageWindow.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _ClassName, "ControlType", "Window");
                WinImageWindow.WaitForControlExist();

                WinTitleBar Image = new WinTitleBar(WinImageWindow);
                Image.SearchProperties.Add("ControlType", "TitleBar");
                Image.WaitForControlReady();
                Image.WaitForControlExist();
                Mouse.Click(Image);

                _logMessage = string.Concat("Clicked on TitleBar: " +Image.AccessibleDescription);
                //Point titlelocation = Microsoft.VisualStudio.TestTools.UITesting.Mouse.Location;
                //Point Statuslocation = new Point(titlelocation.X, titlelocation.Y + 30);
                // Mouse.Click(Statuslocation);




                _methodStatus = _pass;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to Click on TitleBar");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void WinClientAndClickImage(UITestControl Parent, string _title, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinClient winclient = new WinClient(Parent);
                winclient.WindowTitles.Add(_title);

                WinControl wincontrol = new WinControl(winclient);
                wincontrol.SearchProperties.Add(WinControl.PropertyNames.ControlType , "Image");
                wincontrol.WindowTitles.Add(_title);
                UITestControlCollection uic = wincontrol.FindMatchingControls();

                foreach (UITestControl ui in uic)

                {

                    if (ui.BoundingRectangle.Width > 399)

                    {

                        Mouse.DoubleClick(ui);
                        break;
                  

                    }

                }

                _logMessage = string.Concat("Clicked on " +_callerName);
                _methodStatus = _pass;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to Click on " + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinButtonClickByParentClassNameAndButtonName(UITestControl Parent, string _ClassName, string _ButtonName, int iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WinImageWindow = new WinWindow(Parent);
                WinImageWindow.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _ClassName, "ControlType", "Window");
                WinImageWindow.SetFocus();

                WinButton button = new WinButton(WinImageWindow);
                button.SearchProperties.Add(WinButton.PropertyNames.Name, _ButtonName);
                button.WaitForControlExist();
                button.WaitForControlReady();
                // if it doesnt minimize by single click
                if (button.Exists)
                {
                    Mouse.Click(button);
                }
                _logMessage = string.Concat("Clicked on " + _ButtonName + " button");


                _methodStatus = _pass;
            }
            catch (Exception e)
            {
                try
                {
                    if (e.Message.Equals("Element not available"))
                    {
                        WinWindow WinImageWindow = new WinWindow(Parent);
                        WinImageWindow.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _ClassName, "ControlType", "Window");
                        WinImageWindow.SetFocus();

                        WinButton button = new WinButton(WinImageWindow);
                        button.SearchProperties.Add(WinButton.PropertyNames.Name, _ButtonName);
                        button.WaitForControlExist();
                        button.WaitForControlReady();
                        // if it doesnt minimize by single click
                        if (button.Exists)
                        {
                            Mouse.Click(button);
                        }
                        _logMessage = string.Concat("Clicked on " + _ButtonName + " button");
                        _methodStatus = _pass;
                    }
                }
                catch (Exception)
                {
                    _methodStatus = _fail;
                    _logMessage = string.Concat("Failed to click on " + _ButtonName + " button");
                    throw;
                }
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), iteration.ToString()));
            }
        }


        public static void WinButtonClickByParentClassNameAndButtonNameIfExist( string _WindowName, string _ButtonName, int iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WinImageWindow = new WinWindow();
                WinImageWindow.SearchProperties.Add(WinWindow.PropertyNames.Name, _WindowName, "ControlType", "Window");

                if (WinImageWindow.Exists)
                {
                    WinButton button = new WinButton(WinImageWindow);
                    button.SearchProperties.Add(WinButton.PropertyNames.Name, _ButtonName);
                    

                        Mouse.Click(button);
                        _logMessage = string.Concat("Clicked on " + _ButtonName + " button");

                        _methodStatus = _pass;
                    
                    
                }
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _ButtonName + " button");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), iteration.ToString()));
            }
        }
        public static void WinControlClickByParentButtonNameAndWindowName(string _id, string _ClassName, string _ButtonName, int iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {

                WinWindow WinImageWindow = new WinWindow();
                WinImageWindow.SearchProperties.Add(WinWindow.PropertyNames.Name, _id,PropertyExpressionOperator.Contains);
                WinImageWindow.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _ClassName, "ControlType", "Window");
                WinImageWindow.WindowTitles.Add(_id);


                WinButton button = new WinButton(WinImageWindow);
                button.SearchProperties.Add(WinButton.PropertyNames.Name, _ButtonName);
                button.SearchProperties.Add(WinButton.PropertyNames.ControlType, "Button");
                button.WindowTitles.Add(_id);
                //  button.WaitForControlReady();

                WinControl control = new WinControl(button);
                control.SearchProperties.Add(WinControl.PropertyNames.ControlType, "Image");
                control.WindowTitles.Add(_id);

                Mouse.Click(control);

                _logMessage = string.Concat("Clicked on " + _ButtonName + " button");


                _methodStatus = _pass;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _ButtonName + " button");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), iteration.ToString()));
            }
        }
        public static string SplitCurrentDirectory(string ProjectName,string ProjectDataFilePath)
        {
            string directory = System.IO.Directory.GetCurrentDirectory();
            string[] split = Regex.Split(directory,ProjectName);
            string DataPath = Path.Combine(split[0] + ProjectDataFilePath);
            return DataPath;
        }
        public static long CountLinesInFile(string filename)
        {
            long count = 0;
            using (System.IO.StreamReader r = new System.IO.StreamReader(filename))
            {
                string line;
                while ((line = r.ReadLine()) != null)
                {
                    count++;
                }
            }
            return count;
        }


        public static string GetLine(string fileName, int line)
        {
            using (var sr = new System.IO.StreamReader(fileName))
            {
                for (int i = 1; i < line; i++)
                    sr.ReadLine();
                return sr.ReadLine();
            }
        }
        //public enum KeyWord
        //{
        //    Backspace, Pause, Capslock, Delete, End, Return, Esc, Help, Pos1, Insert, Numlock, PageDown, PageUp, Print, Scrollock, Tab, CursorUp,
        //    CursorDown, CursorLeft, CursorRight, Addition, Subtraction, Multiplication, Division,Plus,Caret, Percent, parentheses

        //}
        //public static void KeyBoardAction(string _message, KeyWord _key)
        //{
        //    try
        //    {
        //        switch (_key)
        //        {
        //            case KeyWord.Backspace:
        //                Keyboard.SendKeys(" {BACKSPACE} ");
        //                break;
        //            case KeyWord.Pause:
        //                Keyboard.SendKeys(" {BREAK} ");
        //                break;
        //            case KeyWord.Capslock:
        //                Keyboard.SendKeys(" {CAPSLOCK} ");
        //                break;
        //            case KeyWord.CursorLeft:
        //                Keyboard.SendKeys(" {LEFT} ");
        //                break;
        //            case KeyWord.CursorDown:
        //                Keyboard.SendKeys(" {DOWN} ");
        //                break;
        //            case KeyWord.CursorRight:
        //                Keyboard.SendKeys(" {RIGHT} ");
        //                break;
        //            case KeyWord.CursorUp:
        //                Keyboard.SendKeys(" {UP} ");
        //                break;
        //            case KeyWord.Delete:
        //                Keyboard.SendKeys(" {DEL} ");
        //                break;
        //            case KeyWord.Division:
        //                Keyboard.SendKeys(" {DIVIDE} ");
        //                break;
        //            case KeyWord.End:
        //                Keyboard.SendKeys(" {END} ");
        //                break;
        //            case KeyWord.Esc:
        //                Keyboard.SendKeys(" {ESC} ");
        //                break;
        //            case KeyWord.Help:
        //                Keyboard.SendKeys(" {HELP} ");
        //                break;
        //            case KeyWord.Insert:
        //                Keyboard.SendKeys(" {INSERT} ");
        //                break;
        //            case KeyWord.Multiplication:
        //                Keyboard.SendKeys(" {MULTIPLY} ");
        //                break;
        //            case KeyWord.Numlock:
        //                Keyboard.SendKeys(" {NUMLOCK} ");
        //                break;
        //            case KeyWord.PageDown:
        //                Keyboard.SendKeys(" {PGDN} ");
        //                break;
        //            case KeyWord.Addition:
        //                Keyboard.SendKeys(" {ADD} ");
        //                break;
        //            case KeyWord.Pos1:
        //                Keyboard.SendKeys(" {HOME} ");
        //                break;
        //            case KeyWord.Print:
        //                Keyboard.SendKeys(" {PRTSC} ");
        //                break;
        //            case KeyWord.Return:
        //                Keyboard.SendKeys(" {~} ");
        //                break;
        //            case KeyWord.Scrollock:
        //                Keyboard.SendKeys(" {SCROLLLOCK} ");
        //                break;
        //            case KeyWord.Subtraction:
        //                Keyboard.SendKeys(" {SUBTRACT} ");
        //                break;
        //            case KeyWord.Tab:
        //                Keyboard.SendKeys("{Tab} ");
        //                break;
                   
        //            default:
        //                break;
        //        }



        //    }
        //    catch (Exception)
        //    {
        //        throw;
        //    }

        //}

   
        public static void KeyBoardActionBasedOnAscii(WinEdit uIWinEditObject,int _key)
        {
            try
            {
                switch (_key)
                {
                    case 64:
                        Keyboard.SendKeys(uIWinEditObject, "{@}");
                        break;
                    case 33:
                        Keyboard.SendKeys(uIWinEditObject, "{!}");
                        break;
                    case 43:
                        Keyboard.SendKeys(uIWinEditObject,"{+}");
                        break;
                    case 94:
                        Keyboard.SendKeys(uIWinEditObject,"{^}");
                        break;
                    case 37:
                        Keyboard.SendKeys(uIWinEditObject,"{%}");
                        break;
                    case 123:
                        Keyboard.SendKeys(uIWinEditObject, "{{}");
                        break;
                    case 125:
                        Keyboard.SendKeys(uIWinEditObject, "{}}");
                        break;
                    case 35:
                        Keyboard.SendKeys(uIWinEditObject, "{#}");
                        break;
                    case 36:
                        Keyboard.SendKeys(uIWinEditObject, "{$}");
                        break;
                    case 38:
                        Keyboard.SendKeys(uIWinEditObject, "{&}");
                        break;
                    case 46:
                        Keyboard.SendKeys(uIWinEditObject, "{.}");
                        break;
                    case 45:
                        Keyboard.SendKeys(uIWinEditObject, "{-}");
                        break;
                    case 59:
                        Keyboard.SendKeys(uIWinEditObject, "{;}");
                        break;
                    case 39:
                        Keyboard.SendKeys(uIWinEditObject, "{'}");
                        break;
                    case 60:
                        Keyboard.SendKeys(uIWinEditObject, "{<}");
                        break;
                    case 62:
                        Keyboard.SendKeys(uIWinEditObject, "{>}");
                        break;
                    case 95:
                        Keyboard.SendKeys(uIWinEditObject, "{_}");
                        break;
                    case 44:
                        Keyboard.SendKeys(uIWinEditObject, "{,}");
                        break;
                    case 42:
                        Keyboard.SendKeys(uIWinEditObject, "{*}");
                        break;
                    default:
                        break;
            }



            }
            catch (Exception)
            {
                throw;
            }

        }
        public static void WinMenuItemByClassNameValidationCheckedWithoutUncheck(UITestControl Parent, string _ClassName, string _MenuItem, string _checked, string newMenuItem, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WinMenuItemWindow = new WinWindow();
                WinMenuItemWindow.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _ClassName, "ControlType", "Window");
                if (_MenuItem.Contains("_"))
                {
                    _MenuItem = _MenuItem.Replace("_", " ");
                }

                WinMenuItem MenuItem = new WinMenuItem(WinMenuItemWindow);
                MenuItem.SearchProperties.Add(WinMenuItem.PropertyNames.Name, _MenuItem);
                // MenuItem.WaitForControlReady();
                object check;
                check = MenuItem.GetProperty(WinMenuItem.PropertyNames.Checked);
                string checkedvalue = check.ToString();

                Assert.AreEqual(checkedvalue, _checked);
                _logMessage = string.Concat("Checked: " + _MenuItem);


                WinMenuItem newMenuItem1 = new WinMenuItem(WinMenuItemWindow);
                newMenuItem1.SearchProperties.Add(WinMenuItem.PropertyNames.Name, newMenuItem);
                Mouse.Click(newMenuItem1);
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("UnChecked: " + _MenuItem);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void GetWindowClassNameDialogClickButton(UITestControl Parent, string _classname, string _ButtonName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WindowObj = new WinWindow(Parent);
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _classname);
                // WindowObj.WindowTitles.Add(_id);

                // System.Threading.Thread.Sleep(500);


                WinControl SkypeBusiness = new WinControl(WindowObj);
                SkypeBusiness.SearchProperties.Add("ControlType", "Dialog");
                SkypeBusiness.WaitForControlReady();
                //  SkypeBusiness.DrawHighlight();

                WinButton winClick = new WinButton(SkypeBusiness);
                winClick.SearchProperties.Add(WinButton.PropertyNames.Name, _ButtonName);


                //  winClick.DrawHighlight();

                if (winClick.Exists)
                {
                    winClick.WaitForControlReady();
                    winClick.SetFocus();
                    Mouse.Click(winClick);
                    _logMessage = string.Concat("Clicked on " + _ButtonName);
                    _methodStatus = _pass;
                }


            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Window: " + _callerName + " is not found");
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static List<string> GetWindowByNameAndClassNameAndAllCaptureIM(string _id, string _classname, int _iteration, [CallerMemberName] string _callerName = null)
        {
            List<string> paths = new List<string>();
            paths.Clear();
            string capturedpath = "";
            try
            {
                WinWindow WindowObj = new WinWindow();
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.Name, _id, PropertyExpressionOperator.Contains);
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _classname);
                WindowObj.WindowTitles.Add(_id);

                // _logMessage = String.Concat("Window: " + _id + " is found");
                WindowObj.WaitForControlReady();

                WinClient client = new WinClient(WindowObj);
                client.WindowTitles.Add(_id);

                WinControl control = new WinControl(client);
                control.SearchProperties[WinControl.PropertyNames.ControlType] = "Image";
                //   control.SearchProperties["Instance"] = "2";
                control.WindowTitles.Add(_id);
                //    control.DrawHighlight();
                control.WaitForControlReady();
                //Mouse.Click(control, new Point(controlPt_X, controlPt_Y));
                int count = 0;
                UITestControlCollection uic = control.FindMatchingControls();

                foreach (UITestControl ui in uic)

                {
                    //ui.BoundingRectangle.X > 0 && ui.BoundingRectangle.Y > 0 &&
                    if (ui.BoundingRectangle.Width > 300 && count < 3)

                    {


                        string Verificationsnapshotpath = String.Concat(NGW_SharePoint.Utility.Constants.globalResultsPath + "\\" + "VerificationImage");

                        if (!Directory.Exists(Verificationsnapshotpath))
                        {

                            Directory.CreateDirectory(Verificationsnapshotpath);

                        }
                        Mouse.Click(ui);
                        System.Drawing.Image MyImage = ui.CaptureImage();
                        capturedpath = Path.Combine(Verificationsnapshotpath, _callerName + count + "_" + _iteration);
                        MyImage.Save(capturedpath);
                        paths.Add(capturedpath);
                        //  break;
                        count++;
                    }

                }
                _logMessage = String.Concat("Clicked on the " + _callerName);
                _methodStatus = _pass;
                return paths;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Window: " + _id + " is not found");
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static string OCRPictrueInDisk(string PicturePath, int _iteration, [CallerMemberName] string _callerName = null)
        {

            try
            {
                System.Drawing.Image image = System.Drawing.Image.FromFile(PicturePath);
                using (var ocrEngine = new SFBTesting.LibraryFunctions.OCR())
                    if (!image.Equals(null))
                    {
                        var text = ocrEngine.Recognize(image);
                        if (text != null)
                        {
                            //Console.WriteLine("OcrException:\n");
                            _logMessage = string.Concat("The text recognization of the image is completed");
                            return text;
                        }
                        else
                        {
                            _logMessage = string.Concat("Text is not found in an image");
                            Assert.Fail("Text is not found in an image");
                        }
                    }
                    else
                    {
                        _logMessage = string.Concat("Image is not found");
                        Assert.Fail("Image is not found");
                    }
            }
            catch (Exception ex)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat(ex.Message);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }

            return "Error";
        }

        public static void WinDialogEditByInstance(UITestControl parent, string _DialogName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinControl CreatePollDialog = new WinControl(parent);
                CreatePollDialog.SearchProperties.Add(WinControl.PropertyNames.Name, _DialogName);
                CreatePollDialog.SearchProperties.Add(WinControl.PropertyNames.ControlType, "Dialog");
                CreatePollDialog.WaitForControlReady();
                string[] pollinput = { "SFB", "What is the status of the project?", "Complete", "In-Progress" };
                for (int i = 1; i < 5; i++)
                {
                    WinEdit edit = new WinEdit(CreatePollDialog);
                    edit.SearchProperties.Add(WinEdit.PropertyNames.Instance, i.ToString());
                    edit.WaitForControlReady();
                    if (i == 1)
                    {

                        Mouse.Click(edit);
                        Keyboard.SendKeys("a", System.Windows.Input.ModifierKeys.Control);
                        Keyboard.SendKeys("{DELETE}");
                    }
                    Keyboard.SendKeys(edit, pollinput[i - 1]);
                }
                _logMessage = string.Concat("Values are entered into the edit box respectively ");
                globalFunctions.WinButtonClickByName(CreatePollDialog, "Create", _iteration);
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to find the Dialog: " + _DialogName);

                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static string WinButtonClickByText(UITestControl Parent, string _ButtonName, string _text, string location, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {

                WinButton winButton = new WinButton(Parent);
                winButton.SearchProperties.Add(WinButton.PropertyNames.Name, _ButtonName, "ControlType", "Button");
                //    winButton.WindowTitles.Add(_Title);
                winButton.WaitForControlReady();
                // winButton.DrawHighlight();

                UITestControlCollection child = winButton.GetChildren();
                object a = (object)_text;
                string capturedpath = "";
                foreach (UITestControl ui in child)
                {
                    // ui.DrawHighlight();
                    object name = ui.GetProperty("Name");
                    string Verificationsnapshotpath = String.Concat(NGW_SharePoint.Utility.Constants.globalResultsPath + "\\" + "VerificationImage");

                    if (!Directory.Exists(Verificationsnapshotpath))
                    {

                        Directory.CreateDirectory(Verificationsnapshotpath);

                    }
                    if (name.Equals(a))
                    {
                        Mouse.Click(ui);
                        Keyboard.SendKeys(location);
                        Keyboard.SendKeys("{ENTER}");
                        System.Threading.Thread.Sleep(4500);
                        ui.SetFocus();
                        Mouse.HoverDuration = 4000;
                        Mouse.Hover(ui);

                        //  Mouse.Click(ui);  
                    }
                    else
                    {
                        if (name.Equals(location))
                        {
                            Mouse.Click(ui);
                            Keyboard.SendKeys("a", System.Windows.Input.ModifierKeys.Control);
                            Keyboard.SendKeys("{Delete}");
                            Keyboard.SendKeys(location + "-" + _iteration);
                            Keyboard.SendKeys("{ENTER}");
                            System.Threading.Thread.Sleep(4500);

                            ui.SetFocus();
                            Mouse.HoverDuration = 4000;
                            Mouse.Hover(ui);
                        }
                        else
                        {
                            Mouse.Click(ui);
                            Keyboard.SendKeys("a", System.Windows.Input.ModifierKeys.Control);
                            Keyboard.SendKeys("{Delete}");
                            Keyboard.SendKeys(location);
                            Keyboard.SendKeys("{ENTER}");
                            System.Threading.Thread.Sleep(4500);

                            ui.SetFocus();
                            Mouse.HoverDuration = 4000;
                            Mouse.Hover(ui);
                        }
                    }
                    //WinPane pane = new WinPane(Parent);
                    //pane.SearchProperties.Add(WinPane.PropertyNames.Name, "Contacts");
                    //pane.SearchProperties.Add(WinPane.PropertyNames.ControlType, "Pane");
                    //  pane.DrawHighlight();

                    WinWindow SFBWindow = new WinWindow();
                    SFBWindow.SearchProperties.Add(WinWindow.PropertyNames.ClassName,"CommunicatorMainWindowClass");
                    SFBWindow.SearchProperties.Add(WinWindow.PropertyNames.Name, "Skype for Business ");

                    //System.Drawing.Image MyImage = pane.CaptureImage();
                    System.Drawing.Image MyImage = SFBWindow.CaptureImage();
                    capturedpath = Path.Combine(Verificationsnapshotpath, "LocaltimeCheck.bmp");
                    MyImage.Save(capturedpath);

                }
                _logMessage = String.Concat("Location is set ");
                _methodStatus = _pass;
                return capturedpath;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _ButtonName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static BrowserWindow BrowseURL(string url, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                BrowserWindow browser = BrowserWindow.Launch(url);
                _logMessage = string.Concat(url + " is launched");
                _methodStatus = _pass;

                return browser;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to launch the url: " + url);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static string Sharepoint(UITestControl parent, string Name, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                //BrowserWindow browser = BrowserWindow.Launch(url);
                //_logMessage = string.Concat(url + " is launched");
                //_methodStatus = _pass;

                Microsoft.VisualStudio.TestTools.UITesting.HtmlControls.HtmlDocument document = new Microsoft.VisualStudio.TestTools.UITesting.HtmlControls.HtmlDocument(parent);
                document.SearchProperties.Add(Microsoft.VisualStudio.TestTools.UITesting.HtmlControls.HtmlDocument.PropertyNames.Id, "null");
                document.SearchProperties.Add(Microsoft.VisualStudio.TestTools.UITesting.HtmlControls.HtmlDocument.PropertyNames.RedirectingPage, "False");
                document.SearchProperties.Add(Microsoft.VisualStudio.TestTools.UITesting.HtmlControls.HtmlDocument.PropertyNames.FrameDocument, "False");

                //  document.DrawHighlight();
                HtmlRow row = new HtmlRow(document);
                row.SearchProperties.Add(HtmlRow.PropertyNames.Id, "30,1,1");
                row.SearchProperties.Add(HtmlRow.PropertyNames.Name, null);
                row.FilterProperties[HtmlRow.PropertyNames.InnerText] = "Select or deselect this item\r\n\r\nShared w";
                //  row.DrawHighlight();

                HtmlHyperlink hyperlink = new HtmlHyperlink(row);
                hyperlink.SearchProperties.Add(HtmlHyperlink.PropertyNames.Id, null);
                hyperlink.SearchProperties.Add(HtmlHyperlink.PropertyNames.Name, null);
                hyperlink.SearchProperties.Add(HtmlHyperlink.PropertyNames.Target, null);
                hyperlink.SearchProperties.Add(HtmlHyperlink.PropertyNames.InnerText, Name);
                //  hyperlink.DrawHighlight();
                hyperlink.WaitForControlReady();
                hyperlink.SetFocus();

                string Verificationsnapshotpath = String.Concat(NGW_SharePoint.Utility.Constants.globalResultsPath + "\\" + "VerificationImage");

                if (!Directory.Exists(Verificationsnapshotpath))
                {

                    Directory.CreateDirectory(Verificationsnapshotpath);

                }

                string capturedpath = "";
                Mouse.HoverDuration = 2000;
                Mouse.Hover(hyperlink);
                _logMessage = string.Concat("Hovered on " + Name);
                _methodStatus = _pass;

                HtmlControl control = new HtmlControl(document);
                control.SearchProperties.Add(HtmlControl.PropertyNames.Id, "contentBox");
                control.SearchProperties.Add(HtmlControl.PropertyNames.ControlType, "Pane");
                //  control.DrawHighlight();

                System.Drawing.Image MyImage = control.CaptureImage();
                capturedpath = Path.Combine(Verificationsnapshotpath, _callerName + "_" + _iteration);
                MyImage.Save(capturedpath);



                // _logMessage = string.Concat("Image is saved in the path: " + capturedpath);
                //  _methodStatus = _pass;
                return capturedpath;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to hover on " + Name);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void GetWindowClassNameDialogClickButtonWithTitle(UITestControl Parent, string _classname, string title, string _ButtonName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WindowObj = new WinWindow(Parent);
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _classname);

                WindowObj.WindowTitles.Add(title);

                // System.Threading.Thread.Sleep(500);


                WinControl SkypeBusiness = new WinControl(WindowObj);
                SkypeBusiness.SearchProperties.Add("ControlType", "Dialog");
                //  SkypeBusiness.WaitForControlReady();
                //  SkypeBusiness.DrawHighlight();

                SkypeBusiness.WindowTitles.Add(title);
                WinButton winClick = new WinButton(SkypeBusiness);
                winClick.SearchProperties.Add(WinButton.PropertyNames.Name, _ButtonName);

                winClick.WindowTitles.Add(title);


               // winClick.DrawHighlight();

                if (winClick.Exists)
                {
                    winClick.WaitForControlReady();
                    winClick.SetFocus();
                    Mouse.Click(winClick);
                    _logMessage = string.Concat("Clicked on " + _ButtonName);
                    _methodStatus = _pass;
                }


            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Window: " + _callerName + " is not found");
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void WinClientAndClickAlert(UITestControl Parent, string _name, string _controltype, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinClient winclient = new WinClient(Parent);


                WinControl wincontrol = new WinControl(winclient);
                wincontrol.SearchProperties.Add(WinControl.PropertyNames.Name, _name);
                wincontrol.SearchProperties.Add(WinControl.PropertyNames.ControlType, _controltype);
                if (wincontrol.Exists)
                {
                    Mouse.Hover(wincontrol);
                    _logMessage = string.Concat("Hovered on " + _callerName);
                    _methodStatus = _pass;
                }
                else
                {
                    _methodStatus = _fail;
                    Assert.Fail("Failed to hover on " + _callerName);
                }

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to hover on " + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void WinButtonHoverByNameAndClickImageIfExist(UITestControl Parent, string _ButtonName, string title, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinButton _winClick = new WinButton(Parent);
                _winClick.SearchProperties.Add(WinButton.PropertyNames.Name, _ButtonName);
                _winClick.WindowTitles.Add(title);

                // _winClick.WaitForControlReady();
                if (_winClick.Exists)
                {
                    _winClick.WaitForControlReady();
                    WinControl Image = new WinControl(_winClick);
                    Image.SearchProperties.Add(WinControl.PropertyNames.ControlType, "Image");
                    Image.WindowTitles.Add(title);
                    //  Image.DrawHighlight();
                    // Mouse.Click(Image);
                    UITestControlCollection uic = Image.FindMatchingControls();
                    foreach (UITestControl ui in uic)
                    {
                        _methodStatus = _pass;
                        _logMessage = string.Concat("Clicked on " + _ButtonName + " button");
                        Mouse.Click(ui);
                        break;
                    }
                }
                else
                {
                    _methodStatus = _pass;
                    _logMessage = string.Concat(_ButtonName + " button is not found");
                }

                _methodStatus = _pass;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _ButtonName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static WinWindow GetWindowByClassNameIfExist(string _classname, string title, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WindowObj = new WinWindow();
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _classname, WinWindow.PropertyNames.ControlType, "Window");
                WindowObj.WindowTitles.Add(title);
                // WindowObj.DrawHighlight();
                if (WindowObj.Exists)
                {
                    WindowObj.WaitForControlReady();

                    _logMessage = String.Concat("Window: " + _callerName + " is found");

                }
                else
                {
                    _logMessage = String.Concat("Window: " + _callerName + " is not found");

                }
                _methodStatus = _pass;
                return WindowObj;
            }

            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Window: " + _callerName + " is not found");
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void WinClickByControlTypeIfExist(UITestControl Parent, string _ControlType, string _ButtonName, string title, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {

                WinControl SkypeBusiness = new WinControl(Parent);
                SkypeBusiness.SearchProperties.Add("ControlType", _ControlType);
                SkypeBusiness.WindowTitles.Add(title);
                //commented because og callonhold test case close window
                // SkypeBusiness.SetFocus();
                // SkypeBusiness.DrawHighlight();

                WinButton winClick = new WinButton(SkypeBusiness);
                winClick.SearchProperties.Add(WinButton.PropertyNames.Name, _ButtonName);
                winClick.WindowTitles.Add(title);
                // winClick.DrawHighlight();
                // changes made because of SendFile
                if (winClick.Exists)
                {
                    winClick.WaitForControlReady();
                    Mouse.Click(winClick);
                    _logMessage = string.Concat("Clicked on " + _ButtonName);
                    _methodStatus = _pass;
                }
                else
                {
                    _logMessage = string.Concat(_ButtonName + " button is not found");
                    _methodStatus = _pass;
                }

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _ButtonName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void WinButtonHoverByNameAndClickImage(UITestControl Parent, string _ButtonName, string title, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinButton _winClick = new WinButton(Parent);
                _winClick.SearchProperties.Add(WinButton.PropertyNames.Name, _ButtonName);
                _winClick.WindowTitles.Add(title);

                _winClick.WaitForControlReady();
                //  _winClick.DrawHighlight();
                WinControl Image = new WinControl(_winClick);
                Image.SearchProperties.Add(WinControl.PropertyNames.ControlType, "Image");
                Image.WindowTitles.Add(title);
                //  Image.DrawHighlight();
                // Mouse.Click(Image);
                UITestControlCollection uic = Image.FindMatchingControls();
                foreach (UITestControl ui in uic)
                {
                    _logMessage = string.Concat("Clicked on " + _ButtonName);
                    Mouse.Click(ui);
                    break;
                }

                _methodStatus = _pass;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _ButtonName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinButtonExistClick(UITestControl Parent, string _ButtonName1, string title, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {

                WinButton _winClick1 = new WinButton(Parent);
                _winClick1.SearchProperties.Add(WinButton.PropertyNames.Name, _ButtonName1);
                _winClick1.WindowTitles.Add(title);

                if (_winClick1.Exists)
                {
                    Mouse.Click(_winClick1);
                    _logMessage = string.Concat("Found the button " + _ButtonName1);
                    _methodStatus = _pass;
                }
                else
                {
                    Assert.Fail("Failed to find the " + _ButtonName1);
                    _methodStatus = _fail;
                }

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to find the " + _ButtonName1);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void WinAlertAndClickButton(UITestControl Parent, string _alertname, string _controltype, string _buttonname, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {

                WinControl wincontrol = new WinControl(Parent);
                wincontrol.SearchProperties.Add(WinControl.PropertyNames.Name, _alertname);
                wincontrol.SearchProperties.Add(WinControl.PropertyNames.ControlType, _controltype);
                if (wincontrol.Exists)
                {
                    WinButton button = new WinButton(wincontrol);
                    button.SearchProperties.Add(WinButton.PropertyNames.Name, _buttonname);
                    button.SearchProperties.Add(WinButton.PropertyNames.ControlType, "Button");
                    Mouse.Click(button);
                    _logMessage = string.Concat("Clicked on " + _callerName);
                    _methodStatus = _pass;
                }
                else
                {
                    _methodStatus = _fail;
                    Assert.Fail("Failed to click on " + _callerName);
                }

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void WinControlClickByParentButtonNameAndWindowNameContains(string _id, string _ClassName, string _ButtonName, int iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {

                WinWindow WinImageWindow = new WinWindow();
                WinImageWindow.SearchProperties.Add(WinWindow.PropertyNames.Name, _id, PropertyExpressionOperator.Contains);
                WinImageWindow.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _ClassName, "ControlType", "Window");
                WinImageWindow.WindowTitles.Add(_id);
                WinImageWindow.WaitForControlExist();

                WinButton button = new WinButton(WinImageWindow);
                button.SearchProperties.Add(WinButton.PropertyNames.Name, _ButtonName, PropertyExpressionOperator.Contains);
                button.SearchProperties.Add(WinButton.PropertyNames.ControlType, "Button");
                button.WindowTitles.Add(_id);
                button.WaitForControlReady();

                WinControl control = new WinControl(button);
                control.SearchProperties.Add(WinControl.PropertyNames.ControlType, "Image");
                control.WindowTitles.Add(_id);
                //   control.DrawHighlight();
                control.WaitForControlExist();
                control.WaitForControlReady();
                Mouse.Click(control);

                _logMessage = string.Concat("Clicked on " + _ButtonName + " button");


                _methodStatus = _pass;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _ButtonName + " button");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), iteration.ToString()));
            }
        }
        public static void GetClientWinTextName(UITestControl Parent, string _TextName, string title, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinClient Winclient = new WinClient(Parent);
                Winclient.SearchProperties.Add("ControlType", "Client");
               // Winclient.WindowTitles.Add(title);
                WinText winText = new WinText(Winclient);
                winText.SearchProperties.Add(WinText.PropertyNames.Name, _TextName, PropertyExpressionOperator.Contains);
              //  winText.WindowTitles.Add(title);
                string Name = winText.GetProperty("Name").ToString();

                //Assert.AreEqual(Name, _TextName);
                if (Name == _TextName)
                {
                    _logMessage = string.Concat("The " + _callerName + " text :" + _TextName);
                    _methodStatus = _pass;
                }
                else
                {

                    string name = GetAlphanumericValueWithBracketsAndDecimalAndHypen(Name, _iteration);
                    string textname = GetAlphanumericValueWithBracketsAndDecimalAndHypen(_TextName, _iteration);
                    string Newname = name.Trim('\0');
                    string NewTextname = textname.Trim('\0');
                    Assert.AreEqual(Newname.TrimEnd(), NewTextname.TrimEnd());
                }
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to identify the text " + _TextName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, globalFunctions._logMessage, globalFunctions._methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinClientAndClickAlertControl(UITestControl Parent, string _name, string _controltype, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinClient winclient = new WinClient(Parent);


                WinControl wincontrol = new WinControl(winclient);
                wincontrol.SearchProperties.Add(WinControl.PropertyNames.Name, _name);
                wincontrol.SearchProperties.Add(WinControl.PropertyNames.ControlType, _controltype);
                if (wincontrol.Exists)
                {
                    wincontrol.SetFocus();
                    Mouse.Click(wincontrol);
                    _logMessage = string.Concat("Clicked on " + _name);
                    _methodStatus = _pass;
                }
                else
                {
                    _methodStatus = _fail;
                    Assert.Fail("Failed to click on " + _name);
                }

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void WinClientAndHoverAlertControlAndGetName(UITestControl Parent, string _name, string _controltype, int _iteration, [CallerMemberName] string _callerName = null)
        {
            object name;
            try
            {
                WinClient winclient = new WinClient(Parent);


                WinControl wincontrol = new WinControl(winclient);
                wincontrol.SearchProperties.Add(WinControl.PropertyNames.Name, _name);
                wincontrol.SearchProperties.Add(WinControl.PropertyNames.ControlType, _controltype);
                if (wincontrol.Exists)
                {
                    wincontrol.SetFocus();
                    Mouse.Hover(wincontrol);
                    name = wincontrol.GetProperty("Name");
                    _logMessage = string.Concat("Validation : Hovered on " + name);
                    _methodStatus = _pass;
                }
                else
                {
                    _methodStatus = _fail;
                    Assert.Fail("Failed to hover on " + _callerName);
                }

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to hover on " + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static string VerifyNoOfParticipates(UITestControl Parent, string classname, string title, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                string[] Participant = { "1 Participant", "2 Participants", "3 Participants" };
                bool success = false;
                int count = 0;
                string Name = string.Empty;

                do
                {
                    WinClient client = new WinClient(Parent);
                    WinText text = new WinText(client);
                    text.SearchProperties.Add(WinText.PropertyNames.Name, Participant[count]);
                    // text.SearchProperties.Add(WinText.PropertyNames.ClassName, classname);
                    text.WindowTitles.Add(title);

                    if (text.Exists)
                    {
                        //  WinMenuButtonClickByName.DrawHighlight();
                        text.SetFocus();
                        Mouse.Hover(text);
                        Name = text.GetProperty("Name").ToString();
                        success = true;
                    }

                    count++;
                } while (count < 3 & success == false);
                _logMessage = string.Concat("Participant is :" + Name);

                _methodStatus = _pass;
                return Name.ToString();

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to find the Present Status ");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static object GetComboBoxSelecetedItemSelectComboBox(UITestControl Parent, string _name, string title, string value, string buttonnname, string combo, int _iteration, [CallerMemberName] string _callerName = null)
        {

            try
            {

                WinComboBox winComboBox = new WinComboBox(Parent);
                // winComboBox.SearchProperties.Add(WinComboBox.PropertyNames.ClassName, _classname);
                winComboBox.SearchProperties.Add(WinComboBox.PropertyNames.Name, _name);
                winComboBox.WindowTitles.Add(title);
                UITestControlCollection uic = winComboBox.FindMatchingControls();

                object selecteditem = winComboBox.GetProperty("SelectedItem");
                if (selecteditem.ToString() != value)
                {
                    //Click on Open
                    WinButton button = new WinButton(winComboBox);
                    button.SearchProperties.Add(WinButton.PropertyNames.Name, buttonnname);
                    button.WindowTitles.Add(title);
                    Mouse.Click(button);

                    //Select Day Calendar
                    WinListItem winlistitem = new WinListItem(winComboBox);
                    winlistitem.SearchProperties.Add(WinListItem.PropertyNames.Name, combo);
                    winlistitem.WindowTitles.Add(title);
                    Mouse.Click(winlistitem);
                    _logMessage = String.Concat("Combo box is selected: " + _callerName);

                }

                return selecteditem;

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Validation : Failed to find the value in the combo box is ");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, globalFunctions._logMessage, globalFunctions._methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void ClickOnToolBarDropDownButton(UITestControl Parent, string _toolbarname, string _buttonname, string _title, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {

                WinToolBar winToolBar = new WinToolBar(Parent);
                winToolBar.SearchProperties.Add("Name", _toolbarname);
                winToolBar.WindowTitles.Add(_title);

                WinControl winDropDownButton = new WinControl(winToolBar);
                winDropDownButton.SearchProperties.Add("Name", _buttonname, PropertyExpressionOperator.Contains);
                winDropDownButton.SearchProperties.Add("ControlType", "DropDownButton");
                winDropDownButton.WindowTitles.Add(_title);
                Mouse.Click(winDropDownButton);

                _logMessage = string.Concat("Clicked on the drop down button: " + _buttonname + " in the tool bar");
                _methodStatus = _pass;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on the  drop down button: " + _buttonname);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, globalFunctions._logMessage, globalFunctions._methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinClientButtonEditExist(UITestControl _parent, string _clientclassname, string buttonname, string editname, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinClient client = new WinClient(_parent);
                client.SearchProperties.Add(WinClient.PropertyNames.ClassName, _clientclassname);

                WinButton button = new WinButton(client);
                button.SearchProperties.Add(WinButton.PropertyNames.Name, buttonname);
                if (button.Exists)
                {
                    WinEdit edit = new WinEdit(button);
                    edit.SearchProperties.Add(WinEdit.PropertyNames.Name, editname);
                    edit.WaitForControlReady();
                    //edit.DrawHighlight();
                    Mouse.Click(edit);
                }
                else
                {
                    Assert.Fail(buttonname + " is not found");
                }

                //object currentvalue = edit.GetProperty("");
            }

            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat(buttonname + " is not found");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinClientPaneEditByName(UITestControl _parent, string _clientclassname, string panename, string editname, string value, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinClient client = new WinClient(_parent);
                client.SearchProperties.Add(WinClient.PropertyNames.ClassName, _clientclassname);

                WinPane pane = new WinPane(client);
                pane.SearchProperties.Add(WinPane.PropertyNames.Name, panename);

                WinEdit edit = new WinEdit(pane);
                edit.SearchProperties.Add(WinEdit.PropertyNames.Name, editname);
                edit.WaitForControlReady();
                //edit.DrawHighlight();
                Mouse.Click(edit);
                Keyboard.SendKeys(edit,value);
                Keyboard.SendKeys("{ENTER}");

                //object currentvalue = edit.GetProperty("");
            }

            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to enter " + value + " into input box : " + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static WinMenuItem ClickOnMenuBarMenuItem(UITestControl Parent, string _MenuBarname, string _MenuItemname, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinMenuBar menubar = new WinMenuBar(Parent);
                menubar.SearchProperties.Add("ControlType", "MenuBar");
                menubar.SearchProperties.Add(WinMenuBar.PropertyNames.Name, _MenuBarname);

                WinMenuItem menuitem = new WinMenuItem(menubar);
                menuitem.SearchProperties.Add(WinMenuItem.PropertyNames.Name, _MenuItemname);
                menuitem.SearchProperties.Add(WinMenuItem.PropertyNames.ControlType, "MenuItem");

                Mouse.Click(menuitem);
                _logMessage = string.Concat("Clicked on  " + menuitem);
                _methodStatus = _pass;
                return menuitem;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _MenuItemname);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, globalFunctions._logMessage, globalFunctions._methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static List<string> GetWindowByNameAndClassNameAndCaptureIMbreakaftertwoimages(string _id, string _classname, int _iteration, [CallerMemberName] string _callerName = null)
        {
            List<string> paths = new List<string>();
            string capturedpath = "";
            try
            {
                WinWindow WindowObj = new WinWindow();
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.Name, _id, PropertyExpressionOperator.Contains);
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _classname);
                WindowObj.WindowTitles.Add(_id);

                // _logMessage = String.Concat("Window: " + _id + " is found");
                WindowObj.WaitForControlReady();

                WinClient client = new WinClient(WindowObj);
                client.WindowTitles.Add(_id);

                WinControl control = new WinControl(client);
                control.SearchProperties[WinControl.PropertyNames.ControlType] = "Image";
                //   control.SearchProperties["Instance"] = "2";
                control.WindowTitles.Add(_id);
                //    control.DrawHighlight();
                control.WaitForControlReady();
                //Mouse.Click(control, new Point(controlPt_X, controlPt_Y));
                int count = 0;
                UITestControlCollection uic = control.FindMatchingControls();

                foreach (UITestControl ui in uic)

                {
                    //ui.BoundingRectangle.X > 0 && ui.BoundingRectangle.Y > 0 &&
                    if (ui.BoundingRectangle.Width > 300)

                    {


                        string Verificationsnapshotpath = String.Concat(NGW_SharePoint.Utility.Constants.globalResultsPath + "\\" + "VerificationImage");

                        if (!Directory.Exists(Verificationsnapshotpath))
                        {

                            Directory.CreateDirectory(Verificationsnapshotpath);

                        }
                        Mouse.Click(ui);
                        System.Drawing.Image MyImage = ui.CaptureImage();
                        capturedpath = Path.Combine(Verificationsnapshotpath, _callerName + count + "_" + _iteration);
                        MyImage.Save(capturedpath);
                        paths.Add(capturedpath);

                        count++;
                        if (count == 2)
                            break;
                    }

                }
                _logMessage = String.Concat("Clicked on the " + _callerName);
                _methodStatus = _pass;
                return paths;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Window: " + _id + " is not found");
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static bool WinControlClickByParentButtonNameAndWindowNameContainsIfExist(string _id, string _ClassName, string _ButtonName, int iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                bool success = false;
                WinWindow WinImageWindow = new WinWindow();
                WinImageWindow.SearchProperties.Add(WinWindow.PropertyNames.Name, _id, PropertyExpressionOperator.Contains);
                WinImageWindow.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _ClassName, "ControlType", "Window");
                WinImageWindow.WindowTitles.Add(_id);
                WinImageWindow.WaitForControlExist();

                WinButton button = new WinButton(WinImageWindow);
                button.SearchProperties.Add(WinButton.PropertyNames.Name, _ButtonName);
                button.SearchProperties.Add(WinButton.PropertyNames.ControlType, "Button");
                button.WindowTitles.Add(_id);

                if (button.Exists)
                {
                    success = true;
                    _logMessage = string.Concat(_ButtonName + " button is found");

                }
                else
                {
                    success = false;
                    _logMessage = string.Concat(_ButtonName + " button is not found");
                }

                _methodStatus = _pass;
                return success;

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _ButtonName + " button");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), iteration.ToString()));
            }
        }

        public static void WinListItemHoverByNameAndDoubleClickText(UITestControl Parent, string _objectName, string _textname, string _meetingname, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinListItem _WinListItem = new WinListItem(Parent);
                _WinListItem.SearchProperties.Add(WinListItem.PropertyNames.Name, _objectName, PropertyExpressionOperator.Contains);
                Mouse.HoverDuration = 100;
                Mouse.Hover(_WinListItem);


                WinText wincontrol = new WinText(_WinListItem);
                wincontrol.SearchProperties.Add(WinText.PropertyNames.ControlType, "Text");
                wincontrol.SearchProperties.Add(WinText.PropertyNames.Name, _textname);
                //   wincontrol.DrawHighlight();
                wincontrol.WaitForControlReady();
                Mouse.DoubleClick(wincontrol);
                Keyboard.SendKeys(("^a"));
                Keyboard.SendKeys(" { DELETE} ");
                System.Threading.Thread.Sleep(1000);
                Keyboard.SendKeys(_meetingname);
                _logMessage = string.Concat("Clicked on " + _textname);
                _methodStatus = _pass;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _textname);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void ClickOnFirstButton(UITestControl Parent, int _iteration, [CallerMemberName] string _callerName = null)
        {
            object name;
            try
            {
                WinClient client = new WinClient(Parent);
                client.SearchProperties.Add(WinClient.PropertyNames.ControlType, "Client");
                client.WaitForControlReady();



                WinButton winButton = new WinButton(client);
                winButton.SearchProperties[UITestControl.PropertyNames.ControlType] = "Button";

                UITestControlCollection winButton1 = winButton.FindMatchingControls();
                if (winButton1.Count > 1)
                {
                    foreach (UITestControl ui in winButton1)
                    {
                        ui.WaitForControlReady();
                        Mouse.Click(ui);

                        name = ui.GetProperty("Name");

                        _logMessage = String.Concat("Clicked on : " + name);
                        _methodStatus = _pass;
                        break;
                    }
                }
                else
                {
                    Assert.Fail("No programs were found");
                    _methodStatus = _fail;
                }

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on button");

                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void ButtonByNameExist(UITestControl Parent, string _Buttonname, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinClient client = new WinClient(Parent);
                client.SearchProperties.Add(WinClient.PropertyNames.ControlType, "Client");
                client.WaitForControlReady();

                WinButton winButton = new WinButton(client);
                winButton.SearchProperties[UITestControl.PropertyNames.ControlType] = "Button";
                winButton.SearchProperties[UITestControl.PropertyNames.Name] = _Buttonname;

                if (winButton.Exists)
                {
                    winButton.SetFocus();
                    Mouse.Click(winButton);
                }
                else
                {
                    _logMessage = String.Concat(_Buttonname + " doesn't exist");
                }


                _methodStatus = _pass;

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on : " + _Buttonname);

                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void StartOutlookProcess(string Usernameid, int iteration, [CallerMemberName] string _callerName = null)
        {

            try
            {

                System.Diagnostics.Process myProcess = null;
                //Start 
                System.Diagnostics.ProcessStartInfo Info = new System.Diagnostics.ProcessStartInfo()
                {
                    FileName = "OUTLOOK.exe",
                    WindowStyle = System.Diagnostics.ProcessWindowStyle.Maximized,

                };
                myProcess = System.Diagnostics.Process.Start(Info);

                //  Launch Outlook
                bool Outlook = globalFunctions.GetWindowByNameAndClassNameIfWindowExist(Usernameid + " - Outlook", "rctrl_renwnd32", iteration);

                if (Outlook == false)
                {
                    System.Threading.Thread.Sleep(10000);

                    bool processstatus = globalFunctions.DisplayProcessStatus(myProcess, 1);
                    if (processstatus == true)
                    {
                        Outlook = globalFunctions.GetWindowByNameAndClassNameIfWindowExist(Usernameid + " - Outlook", "rctrl_renwnd32", iteration);
                    }
                    else
                    {
                        myProcess = System.Diagnostics.Process.Start(Info);
                        System.Threading.Thread.Sleep(10000);
                        Outlook = globalFunctions.GetWindowByNameAndClassNameIfWindowExist(Usernameid + " - Outlook", "rctrl_renwnd32", iteration);
                    }
                }
            }
            catch (System.Exception ex)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat(ex.Message);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), iteration.ToString()));
            }
        }
        public static void StartLyncProcess(string Finalusername,string Skypeversion,string email, string username,string password, int iteration, [CallerMemberName] string _callerName = null)
        {

            try
            {

                System.Diagnostics.Process myProcess = null;
                //Start 
                System.Diagnostics.ProcessStartInfo Info = new System.Diagnostics.ProcessStartInfo()
                {
                    FileName = "lync.exe",
                    WindowStyle = System.Diagnostics.ProcessWindowStyle.Maximized,

                };
                myProcess = System.Diagnostics.Process.Start(Info);

                //string details = NGW_SharePoint.Utility.Email.GetCurrentUserInfo();
                //int nameidx = details.IndexOf(":");
                //string[] username = details.Split('\n');
                //string Finalusername = username[0].Substring(nameidx + 1);

                WinWindow SkypeForBusiness = globalFunctions.GetWindowByName("Skype for Business ", 1);


                bool exist = globalFunctions.GetWindowByNameAndClassNameWithoutTitleExist("Skype for Business ", "CommunicatorMainWindowClass", 1);
                if (exist == false)
                {
                    System.Threading.Thread.Sleep(10000);

                    bool processstatus = globalFunctions.DisplayProcessStatus(myProcess, 1);
                    if (processstatus == true)
                    {
                        exist = globalFunctions.GetWindowByNameAndClassNameWithoutTitleExist("Skype for Business ", "CommunicatorMainWindowClass", 1);
                    }
                    else
                    {
                        myProcess = System.Diagnostics.Process.Start(Info);
                        System.Threading.Thread.Sleep(10000);
                        exist = globalFunctions.GetWindowByNameAndClassNameWithoutTitleExist("Skype for Business ", "CommunicatorMainWindowClass", 1);
                    }

                    bool success = true;
                    //Sign in
                    //Launch skype
                    SkypeForBusiness.SetFocus();

                 if (globalFunctions.WinTextNotExist(SkypeForBusiness, Finalusername, iteration))
                 {
                        // Check whether it is in sign out page or not
                        string signintext = globalFunctions.WinGetFriendlNameEditByNameIfExist(SkypeForBusiness, "Sign-in address:", "Sign-in address:", iteration, "Sign-in address");
                    //if sign in text box is not present
                    if (signintext == "")
                    {
                        // sign out
                        SkypeForBusiness.SetFocus();
                        //Click on File in the menu bar
                        globalFunctions.WinDropDownButtonByName(SkypeForBusiness, "File", iteration);
                        //Sign Out
                        if ("2015" == Skypeversion)
                        {
                            globalFunctions.WinMenuItemClickByClassName(SkypeForBusiness, "Net UI Tool Window", "Sign Out", iteration);
                        }
                        else
                        {
                            globalFunctions.WinMenuItemClickByName("Net UI Tool Window", "Sign Out", iteration);
                        }

                        //Check for Sign out dialog
                        globalFunctions.GetWindowByNameAndClassNameIfExist("Skype for Business", "NUIDialog", "Skype for Business", "Yes", iteration);

                       //Minimize
                       globalFunctions.WinButtonClickByParentClassNameAndButtonName(SkypeForBusiness, "NetUIHWND", "Minimize", iteration);
                    }

                    // Check for the hyperlink Change
                    WinHyperlink hyperlink = new WinHyperlink(SkypeForBusiness);
                    hyperlink.SearchProperties.Add(WinHyperlink.PropertyNames.Name, "Change");
                    if (hyperlink.Exists)
                    {
                        Mouse.Click(hyperlink);
                        WinWindow window = globalFunctions.GetWindowByNameAndClassName("Skype for Business - Options", "#32770", iteration);
                        WinWindow subwindow = globalFunctions.GetWindowByParentControlId(window, "14", iteration);
                        // Pass the email id
                        globalFunctions.WinEditByNameWithOutPassingEnterKey(subwindow, "Sign-in address:", email, iteration, "Sign-in address");

                        //Click on OK Button
                        WinWindow subwindowOk = globalFunctions.GetWindowByParentControlId(window, "1", iteration);
                        globalFunctions.WinButtonClickByName(subwindowOk, "OK", iteration);
                    }
                    else
                    {
                        // Pass the email id
                        globalFunctions.WinEditByNameWithOutPassingEnterKey(SkypeForBusiness, "Sign-in address:", email, iteration, "Sign-in address");
                    }

                    // Click on Sign in
                    globalFunctions.ClickOnButtonByNameIfExist(SkypeForBusiness, "Sign In", iteration);

                    //// wait untill Contacting server and signing in... disappears
                    WinText wintext = new WinText(SkypeForBusiness);
                    wintext.SearchProperties.Add(WinText.PropertyNames.Name, "Contacting server and signing in...");
                    wintext.SearchProperties.Add(WinText.PropertyNames.ControlType, "Text");
                    if (wintext.Exists)
                    {
                        wintext.WaitForControlNotExist(40000);
                    }

                    globalFunctions.GetWindowByNameAndClassNameIfExist("Can't sign in to Skype for Business", "NUIDialog", "Can't sign in to Skype for Business", "OK", iteration);

                   //Do you want to Save
                    globalFunctions.GetWindowByAccessibleNameAndClassNameIfExist("Do you want us to save your Skype for Business sign-in info?", "NetUIHWND", "Do you want us to save your Skype for Business sign-in info?", "No", iteration);

                        if (globalFunctions.WinTextNotExist(SkypeForBusiness, Finalusername, iteration))
                    {
                        globalFunctions.WinEditByNameIfExistPassword(SkypeForBusiness, "Password:", password, iteration, "Password");

                        // Click on Sign in
                        globalFunctions.ClickOnButtonByNameIfExist(SkypeForBusiness, "Sign In", iteration);

                        //// wait untill Contacting server and signing in... disappears
                        if (wintext.Exists)
                        {
                            wintext.WaitForControlNotExist(20000);
                        }

                        System.Threading.Thread.Sleep(10000);
                        globalFunctions.GetWindowByNameAndClassNameIfExist("Can't sign in to Skype for Business", "NUIDialog", "Can't sign in to Skype for Business", "OK", iteration);
                            //Do you want to Save
                            globalFunctions.GetWindowByAccessibleNameAndClassNameIfExist("Do you want us to save your Skype for Business sign-in info?", "NetUIHWND", "Do you want us to save your Skype for Business sign-in info?", "No", iteration);

                            success = true;


                        if (globalFunctions.WinTextNotExist(SkypeForBusiness, Finalusername, iteration))
                        {        // user name and password
                                 //Pass username if exist'
                            globalFunctions.WinEditByNameIfExistWithOutPassingEnterKey(SkypeForBusiness, "User name:", username, iteration, "Username");

                            //Pass password if exist
                            globalFunctions.WinEditByNameIfExistPassword(SkypeForBusiness, "Password:", password, iteration, "Password");

                            // Click on Sign in
                            globalFunctions.ClickOnButtonByNameIfExist(SkypeForBusiness, "Sign In", iteration);
                            success = true;
                        }
                        else
                        {
                            success = false;
                        }
                        if (globalFunctions.GetWindowByNameAndClassNameIfExist("Can't sign in to Skype for Business", "NUIDialog", "Can't sign in to Skype for Business", "OK", iteration))
                        {
                            globalFunctions.Result("Can't sign in to Skype for Business, Check the sign in address and logon credentials", "fail", iteration);
                            //Minimize
                            globalFunctions.WinButtonClickByParentClassNameAndButtonName(SkypeForBusiness, "NUIDialog", "Minimize", 1);
                            success = false;
                        }
                            //Do you want to Save
                            globalFunctions.GetWindowByAccessibleNameAndClassNameIfExist("Do you want us to save your Skype for Business sign-in info?", "NetUIHWND", "Do you want us to save your Skype for Business sign-in info?", "No", iteration);

                            if (success == true)
                            { // wait untill Contacting server and signing in... disappears

                                if (wintext.Exists)
                                {
                                    wintext.WaitForControlNotExist(40000);
                                }

                                // check for server pop ups
                                if (globalFunctions.GetWindowByNameAndClassNameIfExist("Can't sign in to Skype for Business", "NUIDialog", "Can't sign in to Skype for Business", "OK", iteration))
                                {
                                    globalFunctions.Result("Can't sign in to Skype for Business, The server is temporarily unavailable", "fail", iteration);
                                }

                                globalFunctions.WinTextNotExist(SkypeForBusiness, Finalusername, iteration);
                            }
                        }
                    }
                }
                else
                {
                    int count = 0;
                    while (globalFunctions.WinTextNotExist(SkypeForBusiness, Finalusername, 1) && count < 2)
                    {
                        System.Threading.Thread.Sleep(10000);
                        count++;
                    }
                    if (count == 2)
                    {
                        bool notexist = globalFunctions.WinTextNotExist(SkypeForBusiness, Finalusername, 1);
                        if (notexist)
                        {
                            bool success = true;
                            //Sign in
                            //Launch skype
                            SkypeForBusiness.SetFocus();

                            // Check whether it is in sign out page or not
                            string signintext = globalFunctions.WinGetFriendlNameEditByNameIfExist(SkypeForBusiness, "Sign-in address:", "Sign-in address:", iteration, "Sign-in address");
                            //if sign in text box is not present
                            if (signintext == "")
                            {
                                // sign out
                                SkypeForBusiness.SetFocus();
                                //Click on File in the menu bar
                                globalFunctions.WinDropDownButtonByName(SkypeForBusiness, "File", iteration);

                                //Sign Out
                                if ("2015" == Skypeversion)
                                {
                                    globalFunctions.WinMenuItemClickByClassName(SkypeForBusiness, "Net UI Tool Window", "Sign Out", iteration);
                                }
                                else
                                {
                                    globalFunctions.WinMenuItemClickByName("Net UI Tool Window", "Sign Out", iteration);
                                }

                                //Check for Sign out dialog
                                globalFunctions.GetWindowByNameAndClassNameIfExist("Skype for Business", "NUIDialog", "Skype for Business", "Yes", iteration);

                                //Minimize
                                globalFunctions.WinButtonClickByParentClassNameAndButtonName(SkypeForBusiness, "NetUIHWND", "Minimize", iteration);
                            }

                            // Check for the hyperlink Change
                            WinHyperlink hyperlink = new WinHyperlink(SkypeForBusiness);
                            hyperlink.SearchProperties.Add(WinHyperlink.PropertyNames.Name, "Change");
                            if (hyperlink.Exists)
                            {
                                Mouse.Click(hyperlink);
                                WinWindow window = globalFunctions.GetWindowByNameAndClassName("Skype for Business - Options", "#32770", iteration);
                                WinWindow subwindow = globalFunctions.GetWindowByParentControlId(window, "14", iteration);
                                // Pass the email id
                                globalFunctions.WinEditByNameWithOutPassingEnterKey(subwindow, "Sign-in address:", email, iteration, "Sign-in address");

                                //Click on OK Button
                                WinWindow subwindowOk = globalFunctions.GetWindowByParentControlId(window, "1", iteration);
                                globalFunctions.WinButtonClickByName(subwindowOk, "OK", iteration);
                            }
                            else
                            {
                                // Pass the email id
                                globalFunctions.WinEditByNameWithOutPassingEnterKey(SkypeForBusiness, "Sign-in address:", email, iteration, "Sign-in address");
                            }

                            // Click on Sign in
                            globalFunctions.ClickOnButtonByNameIfExist(SkypeForBusiness, "Sign In", iteration);

                            //// wait untill Contacting server and signing in... disappears
                            WinText wintext = new WinText(SkypeForBusiness);
                            wintext.SearchProperties.Add(WinText.PropertyNames.Name, "Contacting server and signing in...");
                            wintext.SearchProperties.Add(WinText.PropertyNames.ControlType, "Text");
                            if (wintext.Exists)
                            {
                                wintext.WaitForControlNotExist(40000);
                            }

                            globalFunctions.GetWindowByNameAndClassNameIfExist("Can't sign in to Skype for Business", "NUIDialog", "Can't sign in to Skype for Business", "OK", iteration);
                            //Do you want to Save
                            globalFunctions.GetWindowByAccessibleNameAndClassNameIfExist("Do you want us to save your Skype for Business sign-in info?", "NetUIHWND", "Do you want us to save your Skype for Business sign-in info?", "No", iteration);

                            if (globalFunctions.WinTextNotExist(SkypeForBusiness, Finalusername, iteration))
                            {
                                globalFunctions.WinEditByNameIfExistPassword(SkypeForBusiness, "Password:", password, iteration, "Password");

                                // Click on Sign in
                                globalFunctions.ClickOnButtonByNameIfExist(SkypeForBusiness, "Sign In", iteration);

                                //// wait untill Contacting server and signing in... disappears
                                if (wintext.Exists)
                                {
                                    wintext.WaitForControlNotExist(20000);
                                }

                                System.Threading.Thread.Sleep(10000);
                                globalFunctions.GetWindowByNameAndClassNameIfExist("Can't sign in to Skype for Business", "NUIDialog", "Can't sign in to Skype for Business", "OK", iteration);
                                success = true;
                                //Do you want to Save
                                globalFunctions.GetWindowByAccessibleNameAndClassNameIfExist("Do you want us to save your Skype for Business sign-in info?", "NetUIHWND", "Do you want us to save your Skype for Business sign-in info?", "No", iteration);


                                if (globalFunctions.WinTextNotExist(SkypeForBusiness, Finalusername, iteration))
                                {        // user name and password
                                         //Pass username if exist'
                                    globalFunctions.WinEditByNameIfExistWithOutPassingEnterKey(SkypeForBusiness, "User name:", username, iteration, "Username");

                                    //Pass password if exist
                                    globalFunctions.WinEditByNameIfExistPassword(SkypeForBusiness, "Password:", password, iteration, "Password");

                                    // Click on Sign in
                                    globalFunctions.ClickOnButtonByNameIfExist(SkypeForBusiness, "Sign In", iteration);
                                    success = true;
                                }
                                else
                                {
                                    success = false;
                                }
                                if (globalFunctions.GetWindowByNameAndClassNameIfExist("Can't sign in to Skype for Business", "NUIDialog", "Can't sign in to Skype for Business", "OK", iteration))
                                {
                                    globalFunctions.Result("Can't sign in to Skype for Business, Check the sign in address and logon credentials", "fail", iteration);
                                    //Minimize
                                    globalFunctions.WinButtonClickByParentClassNameAndButtonName(SkypeForBusiness, "NUIDialog", "Minimize", 1);
                                    success = false;
                                }
                                //Do you want to Save
                                globalFunctions.GetWindowByAccessibleNameAndClassNameIfExist("Do you want us to save your Skype for Business sign-in info?", "NetUIHWND", "Do you want us to save your Skype for Business sign-in info?", "No", iteration);

                                if (success == true)
                                { // wait untill Contacting server and signing in... disappears

                                    if (wintext.Exists)
                                    {
                                        wintext.WaitForControlNotExist(40000);
                                    }

                                    // check for server pop ups
                                    if (globalFunctions.GetWindowByNameAndClassNameIfExist("Can't sign in to Skype for Business", "NUIDialog", "Can't sign in to Skype for Business", "OK", iteration))
                                    {
                                        globalFunctions.Result("Can't sign in to Skype for Business, The server is temporarily unavailable", "fail", iteration);
                                    }

                                    globalFunctions.WinTextNotExist(SkypeForBusiness, Finalusername, iteration);
                                 }
                              }
                          }
                        }
                }
            }
            catch (System.Exception ex)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat(ex.Message);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), iteration.ToString()));
            }
        }

        public static bool GetWindowByNameAndClassNameWithoutTitleExist(string _id, string _classname, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                bool exist = false;
                WinWindow WindowObj = new WinWindow();
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.Name, _id, PropertyExpressionOperator.Contains);
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _classname);

                if (WindowObj.WaitForControlExist())
                {
                    _logMessage = String.Concat("Window: " + _id + " is found");
                    exist = true;
                    _methodStatus = _pass;
                }
                else
                {
                    _logMessage = String.Concat("Window: " + _id + " is not found");
                    exist = false;
                    _methodStatus = _pass;
                }
                return exist;
            }
            catch (System.Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Window: " + _id + " is not found");
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void StopProcess(string process, [CallerMemberName] string _callerName = null)
        {
            try
            {
                foreach (System.Diagnostics.Process proc in System.Diagnostics.Process.GetProcessesByName(process))
                {

                    proc.Kill();
                    proc.Refresh();
                    _methodStatus = _pass;
                    _logMessage = string.Concat("StopProcess method has killed the "+process+" process");
                }
            }
            catch (System.Exception ex)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat(ex.Message);
                throw;
            }
           
        }

        public static bool GetWindowByAccessibleNameAndClassNameIfExist(string _id, string _classname, string _DialogName, string _ButtonName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                bool success = false;
                WinWindow WindowObj = new WinWindow();
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.AccessibleName, _id, PropertyExpressionOperator.Contains);
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _classname);
                WindowObj.WindowTitles.Add(_id);

                System.Threading.Thread.Sleep(500);

                if (WindowObj.Exists)
                {
                    WinControl SkypeBusiness = new WinControl(WindowObj);
                    SkypeBusiness.SearchProperties.Add("ControlType", "Dialog");
                    SkypeBusiness.SearchProperties.Add(WinControl.PropertyNames.Name, _DialogName);
                    SkypeBusiness.WaitForControlReady();
                    //   SkypeBusiness.DrawHighlight();

                    WinButton winClick = new WinButton(SkypeBusiness);
                    winClick.SearchProperties.Add(WinButton.PropertyNames.Name, _ButtonName);
                    winClick.WaitForControlReady();

                    //   winClick.DrawHighlight();
                    while (winClick.Exists)
                    {
                        Mouse.Click(winClick);
                        _logMessage = string.Concat("Clicked on " + _ButtonName);
                        _methodStatus = _pass;
                    }
                    success = true;
                }
                else
                {
                    _logMessage = String.Concat("Window: " + _id + " is not found");
                    _methodStatus = _pass;
                    success = false;
                }
                return success;
            }
            catch (System.Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Window: " + _id + " is not found");
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static bool SearchInCitrixWindow(string CitrixWindowName,string OriginalImagePath, float accuracy, int flag, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                bool success = false;
                int count = 0;
                System.Drawing.Bitmap sourceImage = (Bitmap)Bitmap.FromFile(OriginalImagePath);
                System.Threading.Thread.Sleep(1000);


                string Verificationsnapshotpath = String.Concat(NGW_SharePoint.Utility.Constants.globalResultsPath + "\\" + "VerificationImage");

                if (!Directory.Exists(Verificationsnapshotpath))
                {

                    Directory.CreateDirectory(Verificationsnapshotpath);

                }

                // WinWindow capturewindow = globalFunctions.GetWindowByNameAndClassName(CitrixWindowName, "WindowsForms10.Window.8.app.0.2b89eaa_r16_ad1", _iteration);
                WinWindow capturewindow = globalFunctions.GetWindowByName(CitrixWindowName, _iteration);
               // capturewindow.DrawHighlight();
                System.Drawing.Image MyImage = capturewindow.CaptureImage();
                System.Drawing.Rectangle Rectangle = capturewindow.BoundingRectangle;


                string capturedpath = Path.Combine(Verificationsnapshotpath, _callerName + count + "_" + _iteration);
                MyImage.Save(capturedpath);

                // System.Drawing.Bitmap template = (Bitmap)Bitmap.FromFile(@"C:\SFBTesting\Images\2.jpg");
                // create template matching algorithm's instance
                // (set similarity threshold to 92.1%)
                Bitmap img1 = (Bitmap)System.Drawing.Image.FromFile(capturedpath);
                ExhaustiveTemplateMatching tm = new ExhaustiveTemplateMatching(accuracy);
                // find all matchings with specified above similarity

                Bitmap view = globalFunctions.ConvertToFormat(sourceImage, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
                Bitmap view2 = globalFunctions.ConvertToFormat(img1, System.Drawing.Imaging.PixelFormat.Format24bppRgb);

                TemplateMatch[] matchings = tm.ProcessImage(view2, view);
                //Rectangle = {X = 14 Y = 699 Width = 42 Height = 37}
                // highlight found matchings

                BitmapData data = sourceImage.LockBits(
                     new Rectangle(0, 0, sourceImage.Width, sourceImage.Height),
                     ImageLockMode.ReadWrite, sourceImage.PixelFormat);
                Point pt = new Point();
                foreach (TemplateMatch m in matchings)
                {
                    pt.X = m.Rectangle.X + Rectangle.X;
                    pt.Y = m.Rectangle.Y + Rectangle.Y;

                    break;
                    // do something else with matching
                }
                sourceImage.UnlockBits(data);
                if (matchings.Length >= 1)
                {
                    System.Threading.Thread.Sleep(10000);
                    Mouse.Click(new Point(pt.X, pt.Y));
                    globalFunctions.Result("Matching Image is found " + _callerName, "pass", _iteration);
                    success = true;
                }
                else
                {
                    //flag bit to find alternative image matching, alternative image - 0, else 1, 2 for image exist then click
                    if (flag == 2)
                    {
                        globalFunctions.Result(_callerName + "doesn't exist on the desktop screen", "pass", _iteration);
                    }
                    else if (flag == 1)
                    {
                        globalFunctions.Result("Image matching is not found " + _callerName, "fail", _iteration);
                    }
                    else
                    {
                        globalFunctions.Result("Matching Image is not found, try alternative image" + _callerName, "pass", _iteration);
                    }
                }
                count++;
                return success;
            }
            catch (System.Exception ex)
            {
                //Minimize remote window
                System.Windows.Forms.SendKeys.SendWait("^%{BREAK}");
                globalFunctions.ShowDesktop();
                _methodStatus = _fail;
                _logMessage = string.Concat(ex.Message);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static bool SearchInCitrixWindowAndDoubleClick(string CitrixWindowName, string OriginalImagePath, float accuracy, int flag, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                bool success = false;
                int count = 0;
                System.Drawing.Bitmap sourceImage = (Bitmap)Bitmap.FromFile(OriginalImagePath);
                System.Threading.Thread.Sleep(1000);


                string Verificationsnapshotpath = String.Concat(NGW_SharePoint.Utility.Constants.globalResultsPath + "\\" + "VerificationImage");

                if (!Directory.Exists(Verificationsnapshotpath))
                {

                    Directory.CreateDirectory(Verificationsnapshotpath);

                }

                // WinWindow capturewindow = globalFunctions.GetWindowByNameAndClassName(CitrixWindowName, "WindowsForms10.Window.8.app.0.2b89eaa_r16_ad1", _iteration);
                WinWindow capturewindow = globalFunctions.GetWindowByName(CitrixWindowName, _iteration);
              //  capturewindow.DrawHighlight();
                System.Drawing.Image MyImage = capturewindow.CaptureImage();
                System.Drawing.Rectangle Rectangle = capturewindow.BoundingRectangle;


                string capturedpath = Path.Combine(Verificationsnapshotpath, _callerName + count + "_" + _iteration);
                MyImage.Save(capturedpath);

                // System.Drawing.Bitmap template = (Bitmap)Bitmap.FromFile(@"C:\SFBTesting\Images\2.jpg");
                // create template matching algorithm's instance
                // (set similarity threshold to 92.1%)
                Bitmap img1 = (Bitmap)System.Drawing.Image.FromFile(capturedpath);
                ExhaustiveTemplateMatching tm = new ExhaustiveTemplateMatching(accuracy);
                // find all matchings with specified above similarity

                Bitmap view = globalFunctions.ConvertToFormat(sourceImage, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
                Bitmap view2 = globalFunctions.ConvertToFormat(img1, System.Drawing.Imaging.PixelFormat.Format24bppRgb);

                TemplateMatch[] matchings = tm.ProcessImage(view2, view);
                //Rectangle = {X = 14 Y = 699 Width = 42 Height = 37}
                // highlight found matchings

                BitmapData data = sourceImage.LockBits(
                     new Rectangle(0, 0, sourceImage.Width, sourceImage.Height),
                     ImageLockMode.ReadWrite, sourceImage.PixelFormat);
                Point pt = new Point();
                foreach (TemplateMatch m in matchings)
                {
                    pt.X = m.Rectangle.X + Rectangle.X;
                    pt.Y = m.Rectangle.Y + Rectangle.Y;

                    break;
                    // do something else with matching
                }
                sourceImage.UnlockBits(data);
                if (matchings.Length >= 1)
                {
                    
                    Mouse.DoubleClick(new Point(pt.X, pt.Y));
                    Thread.Sleep(10000);
                    globalFunctions.Result("Matching Image is found " + _callerName, "pass", _iteration);
                    success = true;
                }
                else
                {
                    //flag bit to find alternative image matching, alternative image - 0, else 1, 2 for image exist then click
                    if (flag == 2)
                    {
                        globalFunctions.Result(_callerName + "doesn't exist on the desktop screen", "pass", _iteration);
                    }
                    else if (flag == 1)
                    {
                        globalFunctions.Result("Image matching is not found " + _callerName, "fail", _iteration);
                    }
                    else
                    {
                        globalFunctions.Result("Matching Image is not found, try alternative image" + _callerName, "pass", _iteration);
                    }
                }
                count++;
                return success;
            }
            catch (System.Exception ex)
            {
                //Minimize remote window
                System.Windows.Forms.SendKeys.SendWait("^%{BREAK}");
                globalFunctions.ShowDesktop();
                _methodStatus = _fail;
                _logMessage = string.Concat(ex.Message);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void WinMenuItemClickByClassNameIfEnabled(UITestControl Parent, string _ClassName, string _MenuItem, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinWindow WinMenuItemWindow = new WinWindow();
                WinMenuItemWindow.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _ClassName, "ControlType", "Window");
                if (_MenuItem.Contains("_"))
                {
                    _MenuItem = _MenuItem.Replace("_", " ");
                }

                WinMenuItem MenuItem = new WinMenuItem(WinMenuItemWindow);
                MenuItem.SearchProperties.Add(WinMenuItem.PropertyNames.Name, _MenuItem);

                if (MenuItem.Enabled)
                {
                    MenuItem.WaitForControlReady();
                    _logMessage = string.Concat("Clicked on " + _MenuItem);
                    _methodStatus = _pass;
                    Mouse.Click(MenuItem);
                }
                else
                {
                    _logMessage = string.Concat(_MenuItem + " is disabled");
                    _methodStatus = _pass;
                }

            }
            catch (System.Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _MenuItem);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static object GetComboBoxSelecetedItemSelectComboBox(UITestControl Parent, string _name, string title, string value, int _iteration, [CallerMemberName] string _callerName = null)
        {

            try
            {

                WinComboBox winComboBox = new WinComboBox(Parent);
                // winComboBox.SearchProperties.Add(WinComboBox.PropertyNames.ClassName, _classname);
                winComboBox.SearchProperties.Add(WinComboBox.PropertyNames.Name, _name);
                //     winComboBox.DrawHighlight();
                winComboBox.WindowTitles.Add(title);
                UITestControlCollection uic = winComboBox.FindMatchingControls();
                //foreach(UITestControl UI in uic)
                //{
                //    object ab = UI.GetProperty("SelectedItem");
                //}
                object selecteditem = winComboBox.GetProperty("SelectedItem");
                if (selecteditem.ToString() != value)
                {
                    //Click on Open
                    WinButton button = new WinButton(winComboBox);
                    button.SearchProperties.Add(WinButton.PropertyNames.Name, "Open");
                    button.WindowTitles.Add(title);
                    //  button.DrawHighlight();
                    // button.SearchProperties.Add(WinButton.PropertyNames.ClassName, "ComboBox");
                    Mouse.Click(button);

                    //Select Day Calendar
                    WinListItem winlistitem = new WinListItem(winComboBox);
                    winlistitem.SearchProperties.Add(WinListItem.PropertyNames.Name, "Day Calendar");
                    winlistitem.WindowTitles.Add(title);
                    // winlistitem.SearchProperties.Add(WinListItem.PropertyNames.ClassName, "ComboLBox");
                    Mouse.Click(winlistitem);
                    _logMessage = String.Concat("Combo box is selected: " + _callerName);

                }

                return selecteditem;

            }
            catch (System.Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Validation : Failed to find the value in the combo box is ");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, globalFunctions._logMessage, globalFunctions._methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static Bitmap ConvertToFormat(System.Drawing.Bitmap image, PixelFormat format)
        {
            try
            {
                Bitmap copy = new Bitmap(image.Width, image.Height, format);
                using (Graphics gr = Graphics.FromImage(copy))
                {
                    gr.DrawImage(image, new Rectangle(0, 0, copy.Width, copy.Height));
                }
                return copy;
            }
            catch (System.Exception ex)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat(ex.Message);
                throw;
            }
        }

        public static bool SearchUsingUtilities(string OriginalImagePath, float accuracy, int flag, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                bool success = false;
                int count = 0;
                System.Drawing.Bitmap sourceImage = (Bitmap)Bitmap.FromFile(OriginalImagePath);
                System.Threading.Thread.Sleep(1000);


                string Verificationsnapshotpath = String.Concat(NGW_SharePoint.Utility.Constants.globalResultsPath + "\\" + "VerificationImage");

                if (!Directory.Exists(Verificationsnapshotpath))
                {

                    Directory.CreateDirectory(Verificationsnapshotpath);

                }

                string capturedpath = Path.Combine(Verificationsnapshotpath, _callerName + count + "_" + _iteration);

                Utility.Screeshot(capturedpath);
                // create template matching algorithm's instance
                // (set similarity threshold to 92.1%)
                Bitmap img1 = (Bitmap)System.Drawing.Image.FromFile(capturedpath);
                ExhaustiveTemplateMatching tm = new ExhaustiveTemplateMatching(accuracy);
                // find all matchings with specified above similarity

                Bitmap view = globalFunctions.ConvertToFormat(sourceImage, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
                Bitmap view2 = globalFunctions.ConvertToFormat(img1, System.Drawing.Imaging.PixelFormat.Format24bppRgb);

                TemplateMatch[] matchings = tm.ProcessImage(view2, view);
                // highlight found matchings

                BitmapData data = sourceImage.LockBits(
                     new Rectangle(0, 0, sourceImage.Width, sourceImage.Height),
                     ImageLockMode.ReadWrite, sourceImage.PixelFormat);
                Point pt = new Point();

                foreach (TemplateMatch m in matchings)
                {
                    pt = m.Rectangle.Location;

                    // do something else with matching
                }
                sourceImage.UnlockBits(data);
                if (matchings.Length >= 1)
                {
                    Mouse.Click(new Point(pt.X, pt.Y));
                    globalFunctions.Result("Matching Image is found " + _callerName, "pass", _iteration);
                    success = true;
                }
                else
                {
                    //flag bit to find alternative image matching, alternative image - 0, else 1, 2 for image exist then click
                    if (flag == 2)
                    {
                        globalFunctions.Result(_callerName + "doesn't exist on the desktop screen", "pass", _iteration);
                    }
                    else if (flag == 1)
                    {
                        globalFunctions.Result("Image matching is not found " + _callerName, "fail", _iteration);
                    }
                    else
                    {
                        globalFunctions.Result("Matching Image is not found, try alternative image" + _callerName, "pass", _iteration);
                    }
                }
                count++;
                return success;
            }
            catch (System.Exception ex)
            {
                //Minimize remote window
                System.Windows.Forms.SendKeys.SendWait("^%{BREAK}");
                globalFunctions.ShowDesktop();
                _methodStatus = _fail;
                _logMessage = string.Concat(ex.Message);
                throw;
            }
        }
        public static bool SearchRightClickInCitrixWindow(string CitrixWindowName,string OriginalImagePath, float accuracy, int flag, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                bool success = false;
                int count = 0;
                System.Drawing.Bitmap sourceImage = (Bitmap)Bitmap.FromFile(OriginalImagePath);
                System.Threading.Thread.Sleep(1000);


                string Verificationsnapshotpath = String.Concat(NGW_SharePoint.Utility.Constants.globalResultsPath + "\\" + "VerificationImage");

                if (!Directory.Exists(Verificationsnapshotpath))
                {

                    Directory.CreateDirectory(Verificationsnapshotpath);

                }

                WinWindow capturewindow = globalFunctions.GetWindowByName(CitrixWindowName, _iteration);

                // WinWindow capturewindow = globalFunctions.GetWindowByNameAndClassName("SFBTesting_S - Desktop Viewer", "WindowsForms10.Window.8.app.0.2b89eaa_r16_ad1", _iteration);
                System.Drawing.Image MyImage = capturewindow.CaptureImage();
                System.Drawing.Rectangle Rectangle = capturewindow.BoundingRectangle;


                string capturedpath = Path.Combine(Verificationsnapshotpath, _callerName + count + "_" + _iteration);
                MyImage.Save(capturedpath);

                // create template matching algorithm's instance
                // (set similarity threshold to 92.1%)
                Bitmap img1 = (Bitmap)System.Drawing.Image.FromFile(capturedpath);
                ExhaustiveTemplateMatching tm = new ExhaustiveTemplateMatching(accuracy);
                // find all matchings with specified above similarity

                Bitmap view = globalFunctions.ConvertToFormat(sourceImage, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
                Bitmap view2 = globalFunctions.ConvertToFormat(img1, System.Drawing.Imaging.PixelFormat.Format24bppRgb);

                TemplateMatch[] matchings = tm.ProcessImage(view2, view);
               // highlight found matchings

                BitmapData data = sourceImage.LockBits(
                     new Rectangle(0, 0, sourceImage.Width, sourceImage.Height),
                     ImageLockMode.ReadWrite, sourceImage.PixelFormat);
                Point pt = new Point();
                foreach (TemplateMatch m in matchings)
                {
                    pt.X = m.Rectangle.X + Rectangle.X;
                    pt.Y = m.Rectangle.Y + Rectangle.Y;

                    break;
                    // do something else with matching
                }
                sourceImage.UnlockBits(data);
                if (matchings.Length >= 1)
                {
                    Mouse.Click(System.Windows.Forms.MouseButtons.Right, System.Windows.Input.ModifierKeys.None, new Point(pt.X, pt.Y));
                    globalFunctions.Result("Matching Image is found " + _callerName, "pass", _iteration);
                    success = true;
                }
                else
                {
                    //flag bit to find alternative image matching, alternative image - 0, else 1, 2 for image exist then click
                    if (flag == 2)
                    {
                        globalFunctions.Result(_callerName + "doesn't exist on the desktop screen", "pass", _iteration);
                    }
                    else if (flag == 1)
                    {
                        globalFunctions.Result("Image matching is not found " + _callerName, "fail", _iteration);
                    }
                    else
                    {
                        globalFunctions.Result("Matching Image is not found, try alternative image" + _callerName, "pass", _iteration);
                    }
                }
                count++;
                return success;
            }
            catch (System.Exception ex)
            {
                //Minimize remote window
                System.Windows.Forms.SendKeys.SendWait("^%{BREAK}");
                globalFunctions.ShowDesktop();
                _methodStatus = _fail;
                _logMessage = string.Concat(ex.Message);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static bool SearchBottomLeftInCitrixWindow(string CitrixWindowName,string OriginalImagePath, float accuracy,int flag,int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                bool success = false;
                int count = 0;
                System.Drawing.Bitmap sourceImage = (Bitmap)Bitmap.FromFile(OriginalImagePath);
                System.Threading.Thread.Sleep(1000);


                string Verificationsnapshotpath = String.Concat(NGW_SharePoint.Utility.Constants.globalResultsPath + "\\" + "VerificationImage");

                if (!Directory.Exists(Verificationsnapshotpath))
                {

                    Directory.CreateDirectory(Verificationsnapshotpath);

                }


                WinWindow capturewindow = globalFunctions.GetWindowByName(CitrixWindowName, _iteration);
                //  WinWindow capturewindow = globalFunctions.GetWindowByNameAndClassName("SFBTesting_S - Desktop Viewer", "WindowsForms10.Window.8.app.0.2b89eaa_r16_ad1", _iteration);
                System.Drawing.Image MyImage = capturewindow.CaptureImage();
                System.Drawing.Rectangle Rectangle = capturewindow.BoundingRectangle;


                string capturedpath = Path.Combine(Verificationsnapshotpath, _callerName + count + "_" + _iteration);
                MyImage.Save(capturedpath);

                Bitmap originalImage = new Bitmap(System.Drawing.Image.FromFile(capturedpath));
                Rectangle rect = new Rectangle(originalImage.Width / 2, originalImage.Height / 2, originalImage.Width / 2, originalImage.Height / 2);
                //{X = 800 Y = 450 Width = 800 Height = 450}

                Bitmap secondHalf = originalImage.Clone(rect, originalImage.PixelFormat);
                string bottomleftimagepath = Path.Combine(capturedpath + "bottomleft");
                secondHalf.Save(bottomleftimagepath);



                // create template matching algorithm's instance
                // (set similarity threshold to 92.1%)
                Bitmap img1 = (Bitmap)System.Drawing.Image.FromFile(bottomleftimagepath);
                // create template matching algorithm's instance
                // (set similarity threshold to 92.1%)
                ExhaustiveTemplateMatching tm = new ExhaustiveTemplateMatching(accuracy);
                // find all matchings with specified above similarity

                Bitmap view = globalFunctions.ConvertToFormat(sourceImage, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
                Bitmap view2 = globalFunctions.ConvertToFormat(img1, System.Drawing.Imaging.PixelFormat.Format24bppRgb);

                TemplateMatch[] matchings = tm.ProcessImage(view2, view);
                // highlight found matchings

                BitmapData data = sourceImage.LockBits(
                     new Rectangle(0, 0, sourceImage.Width, sourceImage.Height),
                     ImageLockMode.ReadWrite, sourceImage.PixelFormat);
                Point pt = new Point();
                foreach (TemplateMatch m in matchings)
                {
                    pt.X = m.Rectangle.X + Rectangle.X + rect.X;
                    pt.Y = m.Rectangle.Y + Rectangle.Y + rect.Y;

                    break;
                    // do something else with matching
                }
                sourceImage.UnlockBits(data);
                if (matchings.Length >= 1)
                {
                    Mouse.Click(new Point(pt.X, pt.Y));
                    globalFunctions.Result("Matching Image is found " + _callerName, "pass", _iteration);
                    success = true;
                }
                else
                {
                    //flag bit to find alternative image matching, alternative image - 0, else 1, 2 for image exist then click
                    if (flag == 2)
                    {
                        globalFunctions.Result(_callerName + "doesn't exist on the desktop screen", "pass", _iteration);
                    }
                    else if (flag == 1)
                    {
                        globalFunctions.Result("Image matching is not found " + _callerName, "fail", _iteration);
                    }
                    else
                    {
                        globalFunctions.Result("Matching Image is not found, try alternative image" + _callerName, "pass", _iteration);
                    }
                }
                count++;
                return success;
            }
            catch (System.Exception ex)
            {
                //Minimize remote window
                System.Windows.Forms.SendKeys.SendWait("^%{BREAK}");
                globalFunctions.ShowDesktop();

                _methodStatus = _fail;
                _logMessage = string.Concat(ex.Message);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static bool DisplayProcessStatus(Process process, int iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                process.Refresh();  // Important

                bool running = false;

                if (process.HasExited)
                {
                    //Console.WriteLine("Exited.");
                    running = false;
                }
                else
                {
                    // Console.WriteLine("Running.");
                    running = true;
                }
                return running;
            }
            catch (System.Exception ex)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat(ex.Message);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), iteration.ToString()));
            }
        }

        public static bool SearchOnDesktop(string OriginalImagePath, float accuracy, int flag, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                bool success = false;
                int count = 0;
                System.Drawing.Bitmap sourceImage = (Bitmap)Bitmap.FromFile(OriginalImagePath);
                System.Threading.Thread.Sleep(1000);


                string Verificationsnapshotpath = String.Concat(NGW_SharePoint.Utility.Constants.globalResultsPath + "\\" + "VerificationImage");

                if (!Directory.Exists(Verificationsnapshotpath))
                {

                    Directory.CreateDirectory(Verificationsnapshotpath);

                }

                string capturedpath = Path.Combine(Verificationsnapshotpath, _callerName + count + "_" + _iteration);

                System.Drawing.Image MyImage = UITestControl.Desktop.CaptureImage();
                MyImage.Save(capturedpath);

                // create template matching algorithm's instance
                // (set similarity threshold to 92.1%)
                Bitmap img1 = (Bitmap)System.Drawing.Image.FromFile(capturedpath);
                ExhaustiveTemplateMatching tm = new ExhaustiveTemplateMatching(accuracy);
                // find all matchings with specified above similarity

                Bitmap view = globalFunctions.ConvertToFormat(sourceImage, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
                Bitmap view2 = globalFunctions.ConvertToFormat(img1, System.Drawing.Imaging.PixelFormat.Format24bppRgb);

                TemplateMatch[] matchings = tm.ProcessImage(view2, view);
                // highlight found matchings

                BitmapData data = sourceImage.LockBits(
                     new Rectangle(0, 0, sourceImage.Width, sourceImage.Height),
                     ImageLockMode.ReadWrite, sourceImage.PixelFormat);
                Point pt = new Point();

                foreach (TemplateMatch m in matchings)
                {
                    pt = m.Rectangle.Location;

                    // do something else with matching
                }
                sourceImage.UnlockBits(data);
                if (matchings.Length >= 1)
                {
                    Mouse.Click(new Point(pt.X, pt.Y));
                    globalFunctions.Result("Matching Image is found " + _callerName, "pass", _iteration);
                    success = true;
                }
                else
                {
                    //flag bit to find alternative image matching, alternative image - 0, else 1, 2 for image exist then click
                    if (flag == 2)
                    {
                        globalFunctions.Result(_callerName + "doesn't exist on the desktop screen", "pass", _iteration);
                    }
                    else if (flag == 1)
                    {
                        globalFunctions.Result("Image matching is not found " + _callerName, "fail", _iteration);
                    }
                    else
                    {
                        globalFunctions.Result("Matching Image is not found, try alternative image" + _callerName, "pass", _iteration);
                    }
                }
                count++;
                return success;
            }
            catch (System.Exception ex)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat(ex.Message);
                throw;
            }
        }

        public static bool SearchOnDesktopDoubleClick(string OriginalImagePath, float accuracy, int flag, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                bool success = false;
                int count = 0;
                System.Drawing.Bitmap sourceImage = (Bitmap)Bitmap.FromFile(OriginalImagePath);
                System.Threading.Thread.Sleep(1000);


                string Verificationsnapshotpath = String.Concat(NGW_SharePoint.Utility.Constants.globalResultsPath + "\\" + "VerificationImage");

                if (!Directory.Exists(Verificationsnapshotpath))
                {

                    Directory.CreateDirectory(Verificationsnapshotpath);

                }

                string capturedpath = Path.Combine(Verificationsnapshotpath, _callerName + count + "_" + _iteration);

                System.Drawing.Image MyImage = UITestControl.Desktop.CaptureImage();
                MyImage.Save(capturedpath);

                // create template matching algorithm's instance
                // (set similarity threshold to 92.1%)

                Bitmap img1 = (Bitmap)System.Drawing.Image.FromFile(capturedpath);
                ExhaustiveTemplateMatching tm = new ExhaustiveTemplateMatching(accuracy);
                // find all matchings with specified above similarity

                Bitmap view = globalFunctions.ConvertToFormat(sourceImage, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
                Bitmap view2 = globalFunctions.ConvertToFormat(img1, System.Drawing.Imaging.PixelFormat.Format24bppRgb);

                TemplateMatch[] matchings = tm.ProcessImage(view2, view);
                // highlight found matchings

                BitmapData data = sourceImage.LockBits(
                     new Rectangle(0, 0, sourceImage.Width, sourceImage.Height),
                     ImageLockMode.ReadWrite, sourceImage.PixelFormat);
                Point pt = new Point();

                foreach (TemplateMatch m in matchings)
                {
                    pt = m.Rectangle.Location;
                    break;

                    // do something else with matching
                }
                sourceImage.UnlockBits(data);
                if (matchings.Length >= 1)
                {
                    Mouse.Click(new Point(pt.X, pt.Y));
                    globalFunctions.Result("Matching Image is found " + _callerName, "pass", _iteration);
                    success = true;
                }
                else
                {
                    //flag bit to find alternative image matching, alternative image - 0, else 1, 2 for image exist then click
                    if (flag == 2)
                    {
                        globalFunctions.Result(_callerName + "doesn't exist on the desktop screen", "pass", _iteration);
                    }
                    else if (flag == 1)
                    {
                        globalFunctions.Result("Image matching is not found " + _callerName, "fail", _iteration);
                    }
                    else
                    {
                        globalFunctions.Result("Matching Image is not found, try alternative image" + _callerName, "pass", _iteration);
                    }
                }
                count++;
                return success;
            }
            catch (System.Exception ex)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat(ex.Message);
                throw;
            }
        }
        public static bool SearchRightClickOnDesktop(string OriginalImagePath, float accuracy, int flag, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                bool success = false;
                int count = 0;
                System.Drawing.Bitmap sourceImage = (Bitmap)Bitmap.FromFile(OriginalImagePath);
                System.Threading.Thread.Sleep(1000);


                string Verificationsnapshotpath = String.Concat(NGW_SharePoint.Utility.Constants.globalResultsPath + "\\" + "VerificationImage");

                if (!Directory.Exists(Verificationsnapshotpath))
                {

                    Directory.CreateDirectory(Verificationsnapshotpath);

                }

                System.Drawing.Image MyImage = UITestControl.Desktop.CaptureImage();
                string capturedpath = Path.Combine(Verificationsnapshotpath, _callerName + count + "_" + _iteration);
                MyImage.Save(capturedpath);

                // System.Drawing.Bitmap template = (Bitmap)Bitmap.FromFile(@"C:\SFBTesting\Images\2.jpg");
                // create template matching algorithm's instance
                // (set similarity threshold to 92.1%)
                Bitmap img1 = (Bitmap)System.Drawing.Image.FromFile(capturedpath);
                ExhaustiveTemplateMatching tm = new ExhaustiveTemplateMatching(accuracy);
                // find all matchings with specified above similarity

                Bitmap view = globalFunctions.ConvertToFormat(sourceImage, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
                Bitmap view2 = globalFunctions.ConvertToFormat(img1, System.Drawing.Imaging.PixelFormat.Format24bppRgb);

                TemplateMatch[] matchings = tm.ProcessImage(view2, view);
                //Rectangle = {X = 14 Y = 699 Width = 42 Height = 37}
                // highlight found matchings

                BitmapData data = sourceImage.LockBits(
                     new Rectangle(0, 0, sourceImage.Width, sourceImage.Height),
                     ImageLockMode.ReadWrite, sourceImage.PixelFormat);
                Point pt = new Point();
                foreach (TemplateMatch m in matchings)
                {
                    pt = m.Rectangle.Location;

                    break;
                    // do something else with matching
                }
                sourceImage.UnlockBits(data);
                if (matchings.Length >= 1)
                {
                    Mouse.Click(System.Windows.Forms.MouseButtons.Right, System.Windows.Input.ModifierKeys.None, new Point(pt.X, pt.Y));
                    globalFunctions.Result("Matching Image is found " + _callerName, "pass", _iteration);
                    success = true;
                }
                else
                {
                    //flag bit to find alternative image matching, alternative image - 0, else 1, 2 for image exist then click
                    if (flag == 2)
                    {
                        globalFunctions.Result(_callerName + "doesn't exist on the desktop screen", "pass", _iteration);
                    }
                    else if (flag == 1)
                    {
                        globalFunctions.Result("Image matching is not found " + _callerName, "fail", _iteration);
                    }
                    else
                    {
                        globalFunctions.Result("Matching Image is not found, try alternative image" + _callerName, "pass", _iteration);
                    }
                }
                count++;
                return success;
            }
            catch (System.Exception ex)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat(ex.Message);
                throw;
            }
        }
        public static bool SearchGiveControlClick(string OriginalImagePath, float accuracy, int flag, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                bool success = false;
                int count = 0;
                System.Drawing.Bitmap sourceImage = (Bitmap)Bitmap.FromFile(OriginalImagePath);
                System.Threading.Thread.Sleep(8000);


                string Verificationsnapshotpath = String.Concat(NGW_SharePoint.Utility.Constants.globalResultsPath + "\\" + "VerificationImage");

                if (!Directory.Exists(Verificationsnapshotpath))
                {

                    Directory.CreateDirectory(Verificationsnapshotpath);

                }
                string capturedpath = Path.Combine(Verificationsnapshotpath, _callerName + count + "_" + _iteration);


                WinWindow AllowControlWindow = globalFunctions.GetWindowByNameAndClassNameAndTitle("NUIDocumentWindow", "NetUINativeHWNDHost", "NUIDocumentWindow", _iteration, "Allow Control Main Window");
                //newly added
                AllowControlWindow.WaitForControlExist();
                AllowControlWindow.WaitForControlReady();
                WinWindow SubwindowAllowControlWindow = globalFunctions.GetWindowByParentClassName(AllowControlWindow, "NUIDocumentWindow", "NetUIHWND", _iteration, "SubwindowAllowControlWindow");
                WinClient client = new WinClient(SubwindowAllowControlWindow);
                client.SearchProperties.Add(WinClient.PropertyNames.ControlType, "Client");

                WinControl AllowControlImage = new WinControl(client);
                AllowControlImage.SearchProperties.Add(WinControl.PropertyNames.ControlType, "Image");
                //newly added
                AllowControlImage.WindowTitles.Add("NUIDocumentWindow");
                AllowControlImage.WaitForControlReady();
                //AllowControlImage.WindowTitles.Add("NUIDocumentWindow");

                Point Imagelocation = new Point();
                UITestControlCollection uic1 = AllowControlWindow.FindMatchingControls();
                foreach (UITestControl ui in uic1)
                {
                    if (ui.BoundingRectangle.Width > 0)
                    {
                        System.Drawing.Image MyImage = ui.CaptureImage();
                        Imagelocation = ui.BoundingRectangle.Location;
                        MyImage.Save(capturedpath);
                        break;

                    }
                }


                // create template matching algorithm's instance
                // (set similarity threshold to 92.1%)
                Bitmap img1 = (Bitmap)System.Drawing.Image.FromFile(capturedpath);
                ExhaustiveTemplateMatching tm = new ExhaustiveTemplateMatching(accuracy);
                // find all matchings with specified above similarity

                Bitmap view = globalFunctions.ConvertToFormat(sourceImage, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
                Bitmap view2 = globalFunctions.ConvertToFormat(img1, System.Drawing.Imaging.PixelFormat.Format24bppRgb);

                TemplateMatch[] matchings = tm.ProcessImage(view2, view);
                // highlight found matchings

                BitmapData data = sourceImage.LockBits(
                     new Rectangle(0, 0, sourceImage.Width, sourceImage.Height),
                     ImageLockMode.ReadWrite, sourceImage.PixelFormat);
                Point pt = new Point();

                foreach (TemplateMatch m in matchings)
                {
                    pt = m.Rectangle.Location;
                    break;
                    // do something else with matching
                }
                sourceImage.UnlockBits(data);

                if (matchings.Length >= 1)
                {
                    pt.X = pt.X + Imagelocation.X;
                    pt.Y = pt.Y + Imagelocation.Y;
                    Mouse.Click(new Point(pt.X, pt.Y));
                    globalFunctions.Result("Matching Image is found " + _callerName, "pass", _iteration);
                    success = true;
                }
                else
                {
                    //flag bit to find alternative image matching, alternative image - 0, else 1, 2 for image exist then click
                    if (flag == 2)
                    {
                        globalFunctions.Result(_callerName + "doesn't exist on the desktop screen", "pass", _iteration);
                    }
                    if (flag == 1)
                    {
                        globalFunctions.Result("Image matching is not found " + _callerName, "fail", _iteration);
                    }
                    else
                    {
                        globalFunctions.Result("Matching Image is not found, try alternative image" + _callerName, "pass", _iteration);
                    }
                }
                count++;
                return success;
            }
            catch (System.Exception ex)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat(ex.Message);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static WpfWindow GetWPFWindowByName(string _id, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WpfWindow WindowObj = new WpfWindow();
                WindowObj.SearchProperties.Add(WpfWindow.PropertyNames.Name, _id, WpfWindow.PropertyNames.ControlType, "Window");
                _logMessage = String.Concat("Window: " + _id + " is found");
                //WindowObj.DrawHighlight();
                WindowObj.WaitForControlReady();
                _methodStatus = _pass;

                WpfText text = new WpfText(WindowObj);
                text.SearchProperties.Add(WpfText.PropertyNames.AutomationId, "textBlockSupportAccess");
                if (text.DisplayText.Contains("Disabled") || text.DisplayText.Contains("Inaktiv"))
                {
                    WpfButton _grantButton = new WpfButton(WindowObj);
                    _grantButton.SearchProperties.Add(WpfButton.PropertyNames.AutomationId, "buttonEnableSupportAccess");
                    if (_grantButton.Exists)
                    {
                        Mouse.Click(_grantButton);
                        while (text.DisplayText.Contains("Disabled") || text.DisplayText.Contains("Inaktiv"))
                        {
                            text.WaitForControlReady();
                        }
                        WinButtonClickByName(WindowObj, "Close", 0);
                        _methodStatus = _pass;
                    }
                }

                else if (text.DisplayText.Contains("Enabled") || text.DisplayText.Contains("Aktiv"))
                {
                    WinButtonClickByName(WindowObj, "Close", 0);
                }

                return WindowObj;

            }

            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Window: " + _id + " is not found");
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static WinButton WinButtonDoubleClickByName(UITestControl Parent, string _objectName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinButton _winClick = new WinButton(Parent);
                _winClick.SearchProperties.Add(WinButton.PropertyNames.Name, _objectName, PropertyExpressionOperator.Contains);
                _winClick.WaitForControlReady();
                // _winClick.DrawHighlight();
                Mouse.DoubleClick(_winClick);
                Mouse.DoubleClick(_winClick);
                _logMessage = string.Concat("Clicked on " + _objectName);
                //

                _methodStatus = _pass;
                return _winClick;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _objectName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void CheckGrantSupportAccess(string objParentToolbarName, string objToolbarName, string objIWTButton, string objIWTwindow)
        {
            WinWindow _parentToolBar = GetWindowByName(objParentToolbarName, 0);
            WinToolBar _toolbar = WinToolBarByName(_parentToolBar, objToolbarName, 0);
            WinButton _IWTbutton = WinButtonDoubleClickByName(_toolbar, objIWTButton, 0);
            Thread.Sleep(20000);
            WpfWindow _IWTwindow = GetWPFWindowByName(objIWTwindow, 0);
        }
        public static WinToolBar WinToolBarByName(UITestControl _parent, string _id, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinToolBar WindowObj = new WinToolBar(_parent);
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.Name, _id, WinWindow.PropertyNames.ControlType, "ToolBar");
                _logMessage = String.Concat("ToolBar: " + _id + " is found");
                WindowObj.WaitForControlReady();
                _methodStatus = _pass;
                return WindowObj;
            }

            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("ToolBar: " + _id + " is not found");
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }


        public static void WinListItemHoverByNameAndGetChildTextName(UITestControl Parent, string _objectName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinListItem _WinListItem = new WinListItem(Parent);
                _WinListItem.SearchProperties.Add(WinListItem.PropertyNames.Name, _objectName, PropertyExpressionOperator.Contains);
                Mouse.HoverDuration = 100;
                Mouse.Hover(_WinListItem);

                WinText wincontrol = new WinText(_WinListItem);
                wincontrol.SearchProperties.Add(WinText.PropertyNames.ControlType, "Text");
                object name = wincontrol.GetProperty("Name");

                _logMessage = string.Concat(_objectName + _callerName + "is" + name);

            }
            catch (System.Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat(_objectName + " is not found " + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        //public static void RemoteConnection(string IPAddress, string Username, string Password, string Machine_Name, int _iteration, [CallerMemberName] string _callerName = null)
        //{
        //    try
        //    {
        //        //Remote machine
        //        Process rdcProcess = new Process();
        //        string executable = Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\mstsc.exe");

        //        if (executable != null)
        //        {
        //            WinWindow Running_applications = globalFunctions.GetWindowByNameAndClassName("Running applications", "MSTaskSwWClass", 1);
        //            //Check in task bar
        //            if (globalFunctions.HoverOnToolBarButtonOrMenuButton(Running_applications, "Running applications", "IPAddress - Remote Desktop Connection", "", 1))
        //            {

        //            }
        //            else
        //            {
        //                //if icon does not exist
        //                Process[] processRM = System.Diagnostics.Process.GetProcessesByName("mstsc");
        //                //Kill mstsc 
        //                if (processRM.Length > 0)
        //                {
        //                    foreach (System.Diagnostics.Process proc in System.Diagnostics.Process.GetProcessesByName("mstsc"))
        //                    {
        //                        proc.Kill();
        //                    }
        //                }
        //                //start mstsc
        //                rdcProcess.StartInfo.FileName = Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\cmdkey.exe");
        //                rdcProcess.StartInfo.Arguments = "/generic:TERMSRV/" + IPAddress + "/user:" + Username + " /pass:" + Password;
        //                rdcProcess.Start();

        //                rdcProcess.StartInfo.FileName = Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\mstsc.exe");
        //                rdcProcess.StartInfo.Arguments = "/v " + Machine_Name; // ip or name of computer to connect
        //                rdcProcess.Start();
        //                System.Threading.Thread.Sleep(5000);

        //                //Security certificate if exist
        //                if (globalFunctions.GetWindowByExactNameAndClassNameIfWindowExist("Remote Desktop Connection", "#32770", 1, "Security window"))
        //                {
        //                    //Click on Don't ask me again for connections to this computer
        //                    WinWindow securitywindow = globalFunctions.GetWindowByNameAndClassName("Remote Desktop Connection", "#32770", 1, "Security window");
        //                    WinWindow ctrlwindow1 = globalFunctions.GetWindowByParentControlId(securitywindow, "14002", 1);
        //                    WinCheckBox checkbox = new WinCheckBox(ctrlwindow1);
        //                    checkbox.SearchProperties.Add(WinCheckBox.PropertyNames.Name, "Don't ask me again for connections to this computer");
        //                    checkbox.SearchProperties.Add(WinCheckBox.PropertyNames.ControlType, "CheckBox");
        //                    if (checkbox.Exists)
        //                    {
        //                        checkbox.Checked = true;

        //                        //Click on Yes button
        //                        WinWindow ctrlwindow = globalFunctions.GetWindowByParentControlId(securitywindow, "14004", 1);
        //                        globalFunctions.WinButtonClickByName(ctrlwindow, "Yes", 1, "Yes button");
        //                    }
        //                }

        //                //Enter username and password
        //                WinWindow window = globalFunctions.GetWindowByNameAndClassName("Windows Security", "#32770", 1, "login");
        //                window.WaitForControlExist(20000);

        //                /* WinListItem list = new WinListItem(window);
        //                 list.SearchProperties.Add(WinListItem.PropertyNames.Name, Username);
        //                 list.DrawHighlight();

        //                 WinEdit edit = new WinEdit(list);
        //                 edit.SearchProperties.Add(WinEdit.PropertyNames.Name, "Password");
        //                 Keyboard.SendKeys(edit, Password);
        //                 Keyboard.SendKeys("{ENTER}");*/

        //                WinListItem list = new WinListItem(window);
        //                list.SearchProperties.Add(WinListItem.PropertyNames.Name, Username);
        //                if (list.Exists)
        //                {
        //                    WinEdit edit = new WinEdit(list);
        //                    edit.SearchProperties.Add(WinEdit.PropertyNames.Name, "Password");
        //                    Keyboard.SendKeys(edit, Password);
        //                    Keyboard.SendKeys("{ENTER}");
        //                }
        //                else
        //                {
        //                    WinListItem list1 = new WinListItem(window);
        //                    list1.SearchProperties.Add(WinListItem.PropertyNames.Name, "Use another account");
        //                    Mouse.Click(list1);

        //                    WinEdit edit1 = new WinEdit(list1);
        //                    edit1.SearchProperties.Add(WinEdit.PropertyNames.Name, "User name");
        //                    Keyboard.SendKeys(edit1, Username);

        //                    WinEdit edit = new WinEdit(list1);
        //                    edit.SearchProperties.Add(WinEdit.PropertyNames.Name, "Password");
        //                    Keyboard.SendKeys(edit, Password);
        //                    Keyboard.SendKeys("{ENTER}");
        //                }


        //                int count = 0;

        //                //check for the task bar remote machine icon in local machine
        //                while (!globalFunctions.HoverOnToolBarButtonOrMenuButton(Running_applications, "Running applications", IPAddress + " - Remote Desktop Connection", "", 1) && count <= 2)
        //                {
        //                    count++;
        //                    System.Threading.Thread.Sleep(20000);
        //                }
        //                System.Threading.Thread.Sleep(20000);
        //            }
        //        }
        //        else
        //        {
        //            _methodStatus = _fail;
        //            _logMessage = string.Concat("Remote executable is null");
        //            Assert.Fail("Remote executable is null");
        //        }
        //    }
        //    catch (System.Exception)
        //    {
        //        _methodStatus = _fail;
        //        _logMessage = string.Concat("Remote execution failed ");
        //        throw;
        //    }
        //    finally
        //    {
        //        listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
        //    }
        //}

        public static void RemoteConnection(string IPAddress, string Username, string Password, string Machine_Name,string skypeVersion,string remoteMachineIcon, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                //Remote machine
                Process rdcProcess = new Process();
                string executable = Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\mstsc.exe");

                if (executable != null)
                {
                    bool exist = false;
                    WinWindow Running_applications = globalFunctions.GetWindowByNameAndClassName("Running applications", "MSTaskSwWClass", 1);
                    //Check in task bar for 2016
                    if (skypeVersion == "2016")
                    {
                        exist = globalFunctions.HoverOnToolBarButtonOrMenuButton(Running_applications, "Running applications", "Remote Desktop Connection", "", 1);

                        if (exist)
                        {
                           
                            globalFunctions.ClickOnToolBarButton(Running_applications, "Running applications", "Remote Desktop Connection", 1);
                            WinWindow maximizeWindow = globalFunctions.GetWindowByNameAndClassName(remoteMachineIcon, "TscShellContainerClass", 1);
                            Thread.Sleep(2000);
                            maximizeWindow.SetFocus();

                            globalFunctions.ClickOnWindowButtonIfExist(maximizeWindow, "Maximize", 1);

                        }

                    }
                    else
                    {
                        exist = globalFunctions.HoverOnToolBarButtonOrMenuButton(Running_applications, "Running applications", remoteMachineIcon, "", 1);
                        if (exist)
                        {
                            // click on the icon 
                            globalFunctions.ClickOnToolBarButton(Running_applications, "Running applications", remoteMachineIcon,1);
                            WinWindow maximizeWindow = globalFunctions.GetWindowByNameAndClassName(remoteMachineIcon, "TscShellContainerClass", 1);
                            globalFunctions.ClickOnWindowButtonIfExist(maximizeWindow, "Maximize", 1);
                        }

                    }

                    if (exist == false)
                    {
                            //if icon does not exist
                            Process[] processRM = System.Diagnostics.Process.GetProcessesByName("mstsc");
                            //Kill mstsc 
                            if (processRM.Length > 0)
                            {
                                foreach (System.Diagnostics.Process proc in System.Diagnostics.Process.GetProcessesByName("mstsc"))
                                {
                                    proc.Kill();
                                }
                            }
                            //start mstsc
                            //rdcProcess.StartInfo.FileName = Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\cmdkey.exe");
                            //rdcProcess.StartInfo.Arguments = "/generic:TERMSRV/" + IPAddress + "/user:" + Username + " /pass:" + Password;
                            //rdcProcess.Start();

                            rdcProcess.StartInfo.FileName = Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\mstsc.exe");
                            rdcProcess.StartInfo.Arguments = "/v " + Machine_Name; // ip or name of computer to connect
                            rdcProcess.Start();
                            System.Threading.Thread.Sleep(5000);

                            //Security certificate if exist
                            if (globalFunctions.GetWindowByExactNameAndClassNameIfWindowExist("Remote Desktop Connection", "#32770", 1, "Security window"))
                            {
                                //Click on Don't ask me again for connections to this computer
                                WinWindow securitywindow = globalFunctions.GetWindowByNameAndClassName("Remote Desktop Connection", "#32770", 1, "Security window");
                                WinWindow ctrlwindow1 = globalFunctions.GetWindowByParentControlId(securitywindow, "14002", 1);
                                WinCheckBox checkbox = new WinCheckBox(ctrlwindow1);
                                checkbox.SearchProperties.Add(WinCheckBox.PropertyNames.Name, "Don't ask me again for connections to this computer");
                                checkbox.SearchProperties.Add(WinCheckBox.PropertyNames.ControlType, "CheckBox");
                                if (checkbox.Exists)
                                {
                                    checkbox.Checked = true;

                                    //Click on Yes button
                                    WinWindow ctrlwindow = globalFunctions.GetWindowByParentControlId(securitywindow, "14004", 1);
                                    globalFunctions.WinButtonClickByName(ctrlwindow, "Yes", 1, "Yes button");
                                }
                            }

                            //Enter username and password
                            if (skypeVersion == "2016")
                            {

                                WinWindow window1 = globalFunctions.GetWindowByNameAndClassName("Windows Security", "Credential Dialog Xaml Host", 1, "login");
                                window1.WaitForControlExist(20000);

                                WinPane wp = new WinPane();
                                wp.SearchProperties.Add(WinPane.PropertyNames.ClassName, "Credential Dialog Xaml Host");

                                WinText wt = new WinText(wp);
                                wt.SearchProperties.Add(WinText.PropertyNames.Name, Username);

                                if (wt.Exists)
                                {
                                    WinEdit we = new WinEdit(wp);
                                    we.SearchProperties.Add(WinEdit.PropertyNames.Name, "Password");
                                    Mouse.Click(we);
                                    Keyboard.SendKeys(we, Password);
                                    Keyboard.SendKeys("{ENTER}");

                                }
                                else
                                {
                                    WinText wt1 = new WinText(window1);
                                    wt1.SearchProperties.Add(WinText.PropertyNames.Name, "More choices");
                                    Mouse.Click(wt1);

                                    WinListItem list1 = new WinListItem(window1);
                                    list1.SearchProperties.Add(WinListItem.PropertyNames.Name, "Switch to Local or domain account password");
                                    Mouse.Click(list1);

                                    WinPane wp1 = new WinPane(window1);
                                    wp1.SearchProperties.Add(WinPane.PropertyNames.ClassName, "Credential Dialog Xaml Host");

                                    WinEdit we = new WinEdit(wp1);
                                    we.SearchProperties.Add(WinEdit.PropertyNames.Name, "User name");
                                    Keyboard.SendKeys(we, Username);

                                    WinEdit we1 = new WinEdit(wp1);
                                    we1.SearchProperties.Add(WinEdit.PropertyNames.Name, "Password");
                                    Keyboard.SendKeys(we1, Password);
                                    Keyboard.SendKeys("{ENTER}");
                                }
                            }

                        else
                        {
                            WinWindow window = globalFunctions.GetWindowByNameAndClassName("Windows Security", "#32770", 1, "login");
                            window.WaitForControlExist(20000);

                            /* WinListItem list = new WinListItem(window);
                             list.SearchProperties.Add(WinListItem.PropertyNames.Name, Username);
                             list.DrawHighlight();

                             WinEdit edit = new WinEdit(list);
                             edit.SearchProperties.Add(WinEdit.PropertyNames.Name, "Password");
                             Keyboard.SendKeys(edit, Password);
                             Keyboard.SendKeys("{ENTER}");*/

                            WinListItem list = new WinListItem(window);
                            list.SearchProperties.Add(WinListItem.PropertyNames.Name, Username);
                            if (list.Exists)
                            {
                                WinEdit edit = new WinEdit(list);
                                edit.SearchProperties.Add(WinEdit.PropertyNames.Name, "Password");
                                Keyboard.SendKeys(edit, Password);
                                Keyboard.SendKeys("{ENTER}");
                            }
                            else
                            {
                                WinListItem list1 = new WinListItem(window);
                                list1.SearchProperties.Add(WinListItem.PropertyNames.Name, "Use another account");
                                Mouse.Click(list1);

                                WinEdit edit1 = new WinEdit(list1);
                                edit1.SearchProperties.Add(WinEdit.PropertyNames.Name, "User name");
                                Keyboard.SendKeys(edit1, Username);

                                WinEdit edit = new WinEdit(list1);
                                edit.SearchProperties.Add(WinEdit.PropertyNames.Name, "Password");
                                Keyboard.SendKeys(edit, Password);
                                Keyboard.SendKeys("{ENTER}");
                            }

                        }


                        int count = 0;

                        if (skypeVersion == "2016")
                        {
                            while (!globalFunctions.HoverOnToolBarButtonOrMenuButton(Running_applications, "Running applications", "Remote Desktop Connection", "", 1) && count <= 2)
                            {
                                count++;
                                System.Threading.Thread.Sleep(20000);
                            }
                        }
                        else
                        {
                            //check for the task bar remote machine icon in local machine
                            while (!globalFunctions.HoverOnToolBarButtonOrMenuButton(Running_applications, "Running applications", remoteMachineIcon, "", 1) && count <= 2)
                            {
                                count++;
                                System.Threading.Thread.Sleep(20000);
                            }
                        }
                        System.Threading.Thread.Sleep(20000);
                    }
                }
                else
                {
                    _methodStatus = _fail;
                    _logMessage = string.Concat("Remote executable is null");
                    Assert.Fail("Remote executable is null");
                }
            }
            catch (System.Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Remote execution failed ");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static bool GetWindowByExactNameAndClassNameIfWindowExist(string _id, string _classname, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                bool success = false;
                WinWindow WindowObj = new WinWindow();
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.Name, _id);
                WindowObj.SearchProperties.Add(WinWindow.PropertyNames.ClassName, _classname);
                WindowObj.WindowTitles.Add(_id);
                //changed for schedule meeting from sfb // 
                if (WindowObj.Exists)
                {
                    WindowObj.WaitForControlReady();
                    _logMessage = String.Concat("Window: " + _id + " is found");
                    success = true;
                }
                else
                {
                    _logMessage = String.Concat("Window: " + _id + " is not found");
                    success = false;
                }

                _methodStatus = _pass;
                return success;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Window: " + _id + " is not found");
                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void RunPowershell(String filepath, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                //To execute powershell
                PowerShellExec ps = new PowerShellExec();
                var scriptFullpath = filepath;
                string errors = string.Empty;
                string output = string.Empty;
                var success = ps.RunPowerShellScript(scriptFullpath, out output, out errors);
                if (success)
                {
                    _methodStatus = _pass;
                    _logMessage = string.Concat("Powershell script is executed successfully");
                }
                else
                {
                    _logMessage = string.Concat(errors);
                    Assert.Fail("Powershell script  failed to execute");
                }

            }
            catch (System.Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to run Powershell script");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void WinClientHyperLinkGetName(UITestControl _parent, string _id, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinClient client = new WinClient(_parent);
                client.SearchProperties.Add(WinClient.PropertyNames.ControlType, "Client");

                WinHyperlink uIWinEditObject = new WinHyperlink(client);
                uIWinEditObject.SearchProperties.Add(WinHyperlink.PropertyNames.Name, _id, WinHyperlink.PropertyNames.ControlType, "Hyperlink");

                UITestControlCollection uic = uIWinEditObject.FindMatchingControls();
                int c = uic.Count;
                while (c > 0)
                {
                    _logMessage = string.Concat("Clicked on ", _callerName);
                    Mouse.DoubleClick(uIWinEditObject);
                    break;
                }

                _methodStatus = _pass;

            }

            catch (Exception e)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on : " + _callerName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        //public static void ReadAllLinesFromTextAndReplaceFirstThreeLines(String filepath, string _firstLine, string _secondLine, string _thirdLine, int _iteration, [CallerMemberName] string _callerName = null)
        //{
        //    try
        //    {
        //        String PowershellPath = @"C:\Users\uhb5kor\Documents\CleanupLync.ps1";
        //        for (int lineToRead = 0; lineToRead <= 2; lineToRead++)
        //        {
        //            //read specific line
        //            string[] lines = File.ReadAllLines(PowershellPath);
        //            string requiredLine = lines[lineToRead];

        //            //replace that line with required string
        //            string text = File.ReadAllText(PowershellPath);
        //            if (lineToRead == 0)
        //            {
        //                text = text.Replace(requiredLine, _firstLine);
        //            }
        //            else if (lineToRead == 1)
        //            {
        //                text = text.Replace(requiredLine, _secondLine);
        //            }
        //            else
        //            {
        //                text = text.Replace(requiredLine, _thirdLine);
        //            }
        //            File.WriteAllText(PowershellPath, text);

        //            _logMessage = string.Concat("The Server, UserName and Password is updated");


        //            _methodStatus = _pass;
        //        }

        //    }
        //    catch (System.Exception)
        //    {
        //        _methodStatus = _fail;
        //        _logMessage = string.Concat("The Server, UserName and Password is updation failed ");
        //        throw;
        //    }
        //    finally
        //    {
        //        listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
        //    }
        //}

        public static void ReplaceTheFirstOccurances(String filepath, List<string> wordToBeReplaced, List<string> newWords, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                String PowershellPath = filepath;
                //List<string> wordToBeReplaced = new List<string>();
                //wordToBeReplaced.Add("$server");
                //wordToBeReplaced.Add("$Username");
                //wordToBeReplaced.Add("$Password");

                for (int wordToRead = 0; wordToRead < wordToBeReplaced.Count; wordToRead++)
                {
                    int counter = 0;
                    //read specific line
                    string[] lines = File.ReadAllLines(PowershellPath);

                    foreach (string line in lines)
                    {
                            if (line.Contains(wordToBeReplaced[wordToRead]))
                            {
                                string requiredLine = lines[counter];
                                string text = File.ReadAllText(PowershellPath);
                                if (wordToBeReplaced[wordToRead] == "$server")
                                {
                                    text = text.Replace(requiredLine, "$server = '" + newWords[0] +"'");
                                    File.WriteAllText(PowershellPath, text);
                                }
                                else if (wordToBeReplaced[wordToRead] == "$Username")
                                {
                                    text = text.Replace(requiredLine, "$Username = '" + newWords[1] + "'");
                                    File.WriteAllText(PowershellPath, text);
                                }
                                else if (wordToBeReplaced[wordToRead] == "$Password")
                                {
                                    text = text.Replace(requiredLine, "$Password = '" + newWords[2] + "'");
                                    File.WriteAllText(PowershellPath, text);
                                }
                                else if (wordToBeReplaced[wordToRead] == "$testName")
                                {
                                    text = text.Replace(requiredLine, "$testName = '" + newWords[3] + "'");
                                    File.WriteAllText(PowershellPath, text);
                                }
                                break;
                                }
                        counter++;
                    }
                }
            }
            catch (System.Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("The Server, UserName and Password is updation failed ");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }


        public static void WinAlertImageByName(UITestControl _parent, string value, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {

                WinControl CallAlert = new WinControl(_parent);
                CallAlert.SearchProperties.Add(WinControl.PropertyNames.ControlType, "Alert");
                CallAlert.SearchProperties.Add(WinControl.PropertyNames.Name, "Incoming Skype for Business call from " + value + ". Press Windows+Shift+O to accept, Windows+Escape to decline.", PropertyExpressionOperator.Contains);
              //  CallAlert.WaitForControlReady();

                WinControl control = new WinControl(CallAlert);
                control.SearchProperties.Add(WinControl.PropertyNames.ControlType, "Image");
                Mouse.Click(control);
                _methodStatus = _pass;
                _logMessage = string.Concat("Clicked on Incoming Skype for business call");
            }

            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to Clicked on Incoming Skype for business call");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
      

        
        public static void CreateFileAndWrite(String filepath,int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                FileStream fs = null;
                if (!File.Exists(filepath))
                {
                    using (fs = File.Create(filepath))
                    {
                        using (StreamWriter sw = new StreamWriter(filepath))
                        {
                            sw.Write(_iteration);
                        }
                    }
                }
                else
                {
                    using (StreamWriter sw = new StreamWriter(filepath))
                    {
                        sw.Write(_iteration);
                    }
                }

            }
            catch (System.Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to create and write file");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void GetAllProcesses(int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                Process[] processlist = Process.GetProcesses();

                foreach (Process theprocess in processlist)
                {
                    if (theprocess.ProcessName == "ONENOTE")
                    {
                        _logMessage = string.Concat("Process: {0} ID: {1}", theprocess.ProcessName, theprocess.Id);
                        theprocess.Kill();
                        theprocess.Refresh();
                    }
                }

            }
            catch (System.Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to get all processes");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }
        public static void CreateFileAndWrite1(String filepath, string name, int _iteration,[CallerMemberName] string _callerName = null)
        {
            try
            {
                FileStream fs = null;
                if (!File.Exists(filepath))
                {
                    using (fs = File.Create(filepath))
                    {
                        using (StreamWriter sw = new StreamWriter(filepath))
                        {
                            sw.Write(name);
                        }
                    }
                }
                else
                {
                    using (StreamWriter sw = new StreamWriter(filepath))
                    {
                        sw.Write(name);
                    }
                }

            }
            catch (System.Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to create and write file");
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static string ReadFileAndDelete(String filepath, [CallerMemberName] string _callerName = null)
        {
            try
            {
                string value = "";
                if (File.Exists(filepath))
                {
                    using (TextReader tr = new StreamReader(filepath))
                    {
                         value = tr.ReadLine();
                        File.Delete(filepath);
                    }
                }
                return value;

            }
            catch (System.Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to read and delete file ");
                throw;
            }
            
        }
        public static void WinMenuItemAndClickOnWinText(UITestControl Parent, string _MenuItemName, string _TextName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinMenuItem MenuItem = new WinMenuItem(Parent);
                MenuItem.SearchProperties.Add(WinMenuItem.PropertyNames.Name, _MenuItemName);

                WinText WinTextByName = new WinText(MenuItem);
                WinTextByName.SearchProperties.Add(WinText.PropertyNames.Name, _TextName, "ControlType", "Text");
                WinTextByName.SearchConfigurations.Add(SearchConfiguration.ExpandWhileSearching);
                WinTextByName.WaitForControlReady();
                Mouse.Click(WinTextByName);
                _logMessage = string.Concat("Clicked on " + _TextName);


                _methodStatus = _pass;
            }
            catch (System.Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Process failed to click on " + _TextName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static SqlConnection Connect(string username, string password, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                SqlConnection connObj;
                string currentUser = WindowsIdentity.GetCurrent().Name;
                if (!currentUser.Equals(username))
                {
                    connObj = ImpersonationDemo.impersonate(username, password);
                    _methodStatus = _pass;
                    _logMessage = string.Concat("The User is impersonated and the DataBase connection is established");
                }
                else
                {
                    connObj = DatabaseConnection();
                    _methodStatus = _pass;
                    _logMessage = string.Concat("The DataBase connection is established");
                }
                return connObj;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("The DataBase connection is establishment failed");
                return null;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }

        }

        public static SqlConnection DatabaseConnection()
        {
            SqlConnection con = new SqlConnection("Data Source=FE0PLYMD03.de.bosch.com;Initial Catalog=DE_LcsCDR;Persist Security Info=False;Integrated Security=SSPI");
            
                try
                {
                    con.Open();

                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception occurred. " + ex.Message);
                }
            
            return con;
        }

        public static List<DatabaseValues> QoE_DatabaseQuery(SqlConnection con, string CallTime, string Caller, string CallType, int _iteration, string Callee=null,[CallerMemberName] string _callerName = null)
        {
            SqlCommand command;
            string sqlCreateDBQuery = "";
            //SqlDataReader reader;
            List <DatabaseValues> active = new List<DatabaseValues>();
            try
            {
                if (CallType.Equals("P2P"))
                {
                    sqlCreateDBQuery = "SELECT TOP 2 EndTime,SessionType,Caller,CallQuality,Callee,AudioDirection,ConferenceDateTime,DialogID FROM [DE_QoEMetrics].[dbo].[TestResult_View] where [Caller] like " + "'%" + Caller + "%'" + "and [Callee] like " + "'%" + Callee + "%'" + "and[SessionType] = 'P2P'" +" and ConferenceDateTime >= " + "'" + CallTime + "'" + " Order By ConferenceDateTime ASC ";
                }
                else if(CallType.Equals("Conference"))
                {
                    sqlCreateDBQuery = "SELECT TOP 2 EndTime,SessionType,Caller,CallQuality,Callee,AudioDirection,ConferenceDateTime,DialogID FROM [DE_QoEMetrics].[dbo].[TestResult_View] where [Caller] like " + "'%" + Caller + "%'" + "and [SessionType] = 'Conference' " + " and ConferenceDateTime >= " + "'" + CallTime + "'" + " Order By ConferenceDateTime ASC ";

                }
                int i = 0;
                do
                {
                    i++;
                    command = new SqlCommand(sqlCreateDBQuery, con);
                    //Default time of command is 30 seconds
                    command.CommandTimeout = 300;
                    reader = command.ExecuteReader();
                    Thread.Sleep(60000);
                } while (reader == null && i <= 15);

                if (reader != null)
                {
                    while (reader.Read())
                    {
                        var DataBaseItem = new DatabaseValues();
                        DataBaseItem.CallEndTime = reader[0].ToString();
                        DataBaseItem.SessionType = reader[1].ToString();
                        DataBaseItem.Caller = reader[2].ToString();
                        DataBaseItem.CallQuality = reader[3].ToString();
                        DataBaseItem.Callee = reader[4].ToString();
                        DataBaseItem.AudioDirection = reader[5].ToString();
                        DataBaseItem.ConferenceCallTime = reader[6].ToString();
                        DataBaseItem.DialogId = reader[7].ToString();
                        active.Add(DataBaseItem);
                    }
                    return active;
                }
                else
                {
                    throw new Exception();
                }
               
            }
            catch (Exception ex)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("No records found in DE_QoEMetrics Database after querying for 15 times");
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
                throw;

                //return null;
            }
            finally
            {
                reader.Close();
            }
        }


        public static List<DatabaseValues> SplunkP2P_DatabaseQuery(SqlConnection con, string DialogID, int _iteration, [CallerMemberName] string _callerName = null)
        {
            SqlCommand command;

            List<DatabaseValues> active = new List<DatabaseValues>();
            try
            {
                //string sqlCreateDBQuery = "SELECT TOP 1 *FROM [DE_LcsCDR].[dbo].[SplunkP2PView] where DialogId = '"+ DialogID +"'";
                string sqlCreateDBQuery = "SELECT TOP 1 MsDiagHeader,FromUri,ToUri FROM [DE_LcsCDR].[dbo].[SplunkP2PView] where DialogId = '" + DialogID + "'";

                int i = 0;
                do
                {
                    i++;
                    command = new SqlCommand(sqlCreateDBQuery, con);
                    //Default time of command is 30 seconds
                    command.CommandTimeout = 300;
                    reader = command.ExecuteReader();
                    Thread.Sleep(60000);
                } while (reader == null && i <= 15);

                if (reader != null)
                {
                    while (reader.Read())
                    {
                        var DataBaseItem = new DatabaseValues();
                        DataBaseItem.ErrorMessage = reader[0].ToString();
                        DataBaseItem.FromUri = reader[1].ToString();
                        DataBaseItem.ToUri = reader[2].ToString();
                        active.Add(DataBaseItem);
                        // active = DataBaseItem.ErrorMessage;
                    }
                    return active;
                }
                else
                {
                    throw new Exception();
                }
              
            }
            catch (Exception ex)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("No records found in SplunkP2PView Database after querying for 15 times");
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));

                throw;
            }
            finally
            {
                reader.Close();
            }
        }

        public static List<DatabaseValues> McuConference_DatabaseQuery(SqlConnection con, string DialogID, int _iteration, [CallerMemberName] string _callerName = null)
        {
            SqlCommand command;

            List<DatabaseValues> active = new List<DatabaseValues>();
            try
            {
                //string sqlCreateDBQuery = "SELECT TOP 1 *FROM [DE_LcsCDR].[dbo].[SplunkP2PView] where DialogId = '"+ DialogID +"'";
                string sqlCreateDBQuery = "SELECT TOP 1 MsDiagHeader,ConferenceUri,UserUri FROM [DE_LcsCDR].[dbo].[SplunkMcuJoinLeaveView] where DialogId = '" + DialogID + "'";

                int i = 0;
                do
                {
                    i++;
                    command = new SqlCommand(sqlCreateDBQuery, con);
                    //Default time of command is 30 seconds
                    command.CommandTimeout = 300;
                    reader = command.ExecuteReader();
                    Thread.Sleep(60000);
                } while (reader == null && i <= 15);

                if (reader != null)
                {
                    while (reader.Read())
                    {
                        var DataBaseItem = new DatabaseValues();
                        DataBaseItem.ErrorMessage = reader[0].ToString();
                        DataBaseItem.ConferenceUri = reader[1].ToString();
                        DataBaseItem.UserUri = reader[2].ToString();
                        active.Add(DataBaseItem);
                    }
                    return active;
                }
                else
                {
                    throw new Exception();
                }

            }
            catch (Exception ex)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("No records found in SplunkMcuJoinLeaveView Database after querying for 15 times");
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
                throw;
            }
            finally
            {
                reader.Close();
            }
        }

        public class DatabaseValues
        {
            public string CallEndTime { get; set; }
            public string SessionType { get; set; }
            public string Caller { get; set; }
            public string CallQuality { get; set; }
            public string Callee { get; set; }
            public string AudioDirection { get; set; }
            public string ConferenceCallTime { get; set; }
            public string DialogId { get; set; }

            public string ErrorMessage { get; set; }
            public string ToUri { get; set; }
            public string FromUri { get; set; }

            public string UserUri { get; set; }

            public string ConferenceUri { get; set; }
            
        }

        public static void VerifyDatabaseCallQualtiy(string audioDirection, string CallQuality, string Caller,string Callee,string ConferenceCallTime,string DialogId, string sessiontype,SqlConnection con,int _iteration, [CallerMemberName] string _callerName = null)
        {
            List<DatabaseValues> message=new List<DatabaseValues>(); ;
            //string P2PQuery = "[DE_LcsCDR].[dbo].[SplunkP2PView]";
            //string McuConferenceQuery = "[DE_LcsCDR].[dbo].[SplunkMcuJoinLeaveView]";
            try
            {
                if (audioDirection.Equals("Caller_to_Callee") || audioDirection.Equals("Callee_to_Caller") || audioDirection.Equals("Server_to_Participant") || audioDirection.Equals("Participant_to_Server"))
                {
                    if (CallQuality.Equals("Good"))
                    {
                        _methodStatus = _pass;
                        _logMessage = string.Concat(sessiontype +" Call Quality from " +Caller +" to "+Callee+" for "+audioDirection + " at "+ ConferenceCallTime + " is found " + CallQuality);
                    }
                    else
                    {
                        if (sessiontype.Equals("P2P"))
                        {
                            message = SplunkP2P_DatabaseQuery(con, DialogId,1);
                            _logMessage = string.Concat("P2P Audio Call Quality from " + message[0].FromUri + " to " + message[0].ToUri + " for " + audioDirection + " at " + ConferenceCallTime + " is found " + CallQuality + "   Error Message:  " + message[0].ErrorMessage);

                        }
                        else if (sessiontype.Equals("Conference") || sessiontype.Equals("Conference_Mcu"))
                        {
                             message = McuConference_DatabaseQuery(con, DialogId,1);
                            _logMessage = string.Concat("Conference Audio Call Quality from " + message[0].ConferenceUri + " to " + message[0].UserUri + " for " + audioDirection + " at " + ConferenceCallTime + " is found " + CallQuality + "   Error Message:  " + message[0].ErrorMessage);

                        }

                        _methodStatus = _fail;
                        throw new Exception();

                    }
                }
            }
            catch (Exception ex )
            {
                _methodStatus = _fail;
                //_logMessage = string.Concat("The DataBase connection is establishment failed");
                throw;

            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static void InsertDateTimeToDataFile(string _filePath, int _column, int _row, [CallerMemberName] string _callerName = null)
        {
            try
            {
                FileInfo fi = new FileInfo(_filePath);
                if (fi.Exists)
                {

                    TimeZoneInfo timeZoneInfo;
                    DateTime dateTime;
                    //Set the time zone information to US Mountain Standard Time 
                    timeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Greenwich Standard Time");
                    //Get date and time in US Mountain Standard Time 
                    dateTime = TimeZoneInfo.ConvertTime(DateTime.Now, timeZoneInfo);

                    string format = "yyyy-M-d HH:mm:ss.fff";
                    string time = dateTime.ToString(format);

                    //total number of records should be same 
                    List<string> lines = File.ReadAllLines(_filePath).ToList();

                    //row
                    string row_line = lines[_row];
                    string[] split = row_line.Split(',');

                    string format1 = "yyyy";
                    //4th column - 3rd index
                    if (split[_column].Contains(dateTime.ToString(format1)))
                    {
                        List<String> list = new List<String>(split);

                        //Remove string at column 
                        list.RemoveAt(_column);

                        // Add string at column
                        list.Insert(_column, time);

                        //Clear thespecified index row
                        lines.RemoveAt(_row);

                        //insert the complete row 
                        string finalrow = string.Join(",", list.ToArray());
                        lines.Insert(_row, finalrow);
                    }
                    else
                    {
                        //append to the end of the row
                        lines[_row] += time + ",";
                    }

                    //write the new content
                    File.WriteAllLines(_filePath, lines);

                    _methodStatus = _pass;
                }
                else
                {
                    //file doesn't exist
                    _logMessage = String.Concat(_filePath, " data file doesn't exist");
                    _methodStatus = _fail;
                }

            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Inserting the data and time to the " + _filePath + " data file failed");
                throw;
            }

        }

       

        public static void ClickOnPresentButtonByName(UITestControl Parent, string _Buttonname, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {
                WinClient client = new WinClient(Parent);
                client.SearchProperties.Add(WinClient.PropertyNames.ControlType, "Client");
                client.WaitForControlReady();

                WinButton winButton = new WinButton(client);
                winButton.SearchProperties[UITestControl.PropertyNames.ControlType] = "Button";
                winButton.SearchProperties[UITestControl.PropertyNames.Name] = _Buttonname;
                winButton.WaitForControlReady();
                Point Imagelocation = new Point();
                if (winButton.Exists)
                {
                    winButton.SetFocus();
                    //Mouse.Click(winButton);

                    Imagelocation = winButton.BoundingRectangle.Location;
                    Imagelocation.Offset(winButton.BoundingRectangle.Width / 2, winButton.BoundingRectangle.Height / 2);
                    Mouse.Click(Imagelocation);

                    _logMessage = String.Concat("Clicked on : " + _Buttonname);
                    _methodStatus = _pass;
                }
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on : " + _Buttonname);

                throw;
            }
            finally
            {

                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static bool WinDialogClickByNameIfEnabled(UITestControl Parent, string _DialogName, string _ButtonName, int _iteration, [CallerMemberName] string _callerName = null)
        {
            bool enable = false;
            try
            {
                //WinWindow Skype_for_Business = new WinWindow();
                //Skype_for_Business.SearchProperties.Add(WinControl.PropertyNames.ClassName, _WindowClassName);
                //Skype_for_Business.SearchProperties.Add(WinControl.PropertyNames.ControlType, "Window");

                WinControl SkypeBusiness = new WinControl(Parent);
                SkypeBusiness.SearchProperties.Add("ControlType", "Dialog");
                SkypeBusiness.SearchProperties.Add(WinControl.PropertyNames.Name, _DialogName);
                SkypeBusiness.WaitForControlReady();
                // SkypeBusiness.DrawHighlight();

                WinButton winClick = new WinButton(SkypeBusiness);
                winClick.SearchProperties.Add(WinButton.PropertyNames.Name, _ButtonName);
                winClick.WaitForControlReady();

                //  winClick.DrawHighlight();
                if (winClick.Enabled)
                {
                    enable = true;
                    Mouse.Click(winClick);
                    _logMessage = string.Concat("Clicked on " + _ButtonName);
                    _methodStatus = _pass;
                }
                else
                {
                    enable = false;
                }
                return enable;
            }
            catch (Exception)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat("Failed to click on " + _ButtonName);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
        }

        public static string GetOCRText(string PicturePath, int _iteration, [CallerMemberName] string _callerName = null)
        {
            try
            {

                System.Drawing.Image image = System.Drawing.Image.FromFile(PicturePath);
                if (!image.Equals(null))
                {
                    Tesseract_OCR ocrObject = new Tesseract_OCR();
                    string OCR_Text = ocrObject.ImageExtraction(PicturePath);
                    if (OCR_Text != null)
                    {
                        _logMessage = string.Concat("The text recognization of the image is completed");
                        return OCR_Text;
                    }
                    else
                    {
                        _logMessage = string.Concat("Text is not found in an image");
                        Assert.Fail("Text is not found in an image");
                    }
                }
                else
                {
                    _logMessage = string.Concat("Image is not found");
                    Assert.Fail("Image is not found");
                }

            }
            catch (Exception ex)
            {
                _methodStatus = _fail;
                _logMessage = string.Concat(ex.Message);
                throw;
            }
            finally
            {
                listOfTuples.Add(new Tuple<String, String, String, String, String, String>(System.Reflection.MethodBase.GetCurrentMethod().Name, _callerName, _logMessage, _methodStatus, NGW_SharePoint.Utility.Utility.GetCurrentDateTime(), _iteration.ToString()));
            }
            return "Error";
        }
             
        #endregion'
    }
}
