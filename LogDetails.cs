using System;
using System.Text;
using System.Data;
using System.IO;
using System.Collections.Generic;

namespace NGW_SharePoint.Utility
{
    class LogDetails
    {
        public static int count = 0;
        static string HTMLTEXT = null;

        public static void Log(List<Tuple<String, String, String, String, String, String>> l, string datarow)
        {
            // Here we create a DataTable with 8 columns.
            DataTable _table = new DataTable();

            //Adding columns to DataTable
            _table.Columns.Add("ProjectName", typeof(string));
            _table.Columns.Add("CodedUIMethod", typeof(string));
            _table.Columns.Add("TestMethod", typeof(string));
            _table.Columns.Add("Results", typeof(string));
            _table.Columns.Add("TestUser", typeof(string));
            _table.Columns.Add("TimeStamp", typeof(string));
            _table.Columns.Add("DataRow", typeof(string));
            _table.Columns.Add("Message", typeof(string));

            // Here we add all the DataRows
            for (int i = 0; i < l.Count; i++)
            {
                _table.Rows.Add(Constants.currentProjectName, l[i].Item2, l[i].Item1, l[i].Item4, Utility.GetCurrentAdUserName(), l[i].Item5, datarow, l[i].Item3);
            }


            //for (int i = 0; i < l.Count; i++) // Loop through List with for
            //{
            //    Constants.logTuple.Add(l[i]);
            //}

            //Stringbuilder class to convert from datatable to HTML
            StringBuilder _builder = new StringBuilder();
            _builder.Append("<html>");
            _builder.Append("<header>");
            _builder.Append("<head>");
            _builder.Append("<title>");
            _builder.Append("Log Report");
            _builder.Append("</title>");
            _builder.Append("</head>");
            _builder.Append("<body>");

            //Table Alignment    
            _builder.Append("<table cellspacing='0' cellpadding='4' border= 'thin'  width='100' style ='border-style:groove;border-top-width:thin;font-size:x-small;table-layout:fixed;font-family:arial,sans-serif;font-weight: bold;'>");

            //Fixing the width of each column
            _builder.Append("<col width='100'>");
            _builder.Append("<col width='200'>");
            _builder.Append("<col width='240'>");
            _builder.Append("<col width='140'>");
            _builder.Append("<col width='140'>");
            _builder.Append("<col width='150'>");
            _builder.Append("<col width='70'>");
            _builder.Append("<col width='540'>");

            //Below if condition is used to avoid adding headers each and every time this function is called. Header will be added only during the first Instance
            if (count == 0)
            {
                _builder.Append("<tr bgcolor='SteelBlue' style='font-family:arial,sans-serif' align='left' valign='top'>");
                foreach (DataColumn c in _table.Columns)
                {
                    _builder.Append("<td align='left' valign='top' style='color:#FFFFFF'><b>");
                    _builder.Append(c.ColumnName);
                    _builder.Append("</b></td>");
                }
                count = 1;
                goto tag;
            }
            else
            {
                goto tag;
            }
            tag:
            _builder.Append("</tr>");
            foreach (DataRow r in _table.Rows)
            {
                _builder.Append("<tr align='left' valign='top' >");
                foreach (DataColumn c in _table.Columns)
                {
                    //Below If condition is used to check whether the Status column value is Failed If so then Red colour will be given to that particular cell
                    if (r[c.ColumnName].ToString() == "Failed")
                    {
                        _builder.Append("<td align='left' valign='top'>");
                        _builder.Append("<font color='#660000'>");
                        _builder.Append(r[c.ColumnName]);
                        _builder.Append("</font>");
                        goto build;
                    }
                    else if (r[c.ColumnName].ToString() == "Inconclusive")
                    {
                        _builder.Append("<td align='left' valign='top'>");
                        _builder.Append("<font color='663300'>");
                        _builder.Append(r[c.ColumnName]);
                        _builder.Append("</font>");
                        goto build;
                    }
                    else if (r[c.ColumnName].ToString() == "Passed")
                    {
                        _builder.Append("<td align='left' valign='top'>");
                        _builder.Append("<font color='#006600'>");
                        _builder.Append(r[c.ColumnName]);
                        _builder.Append("</font>");
                        goto build;
                    }
                    else
                    {
                        _builder.Append("<td align='left' valign='top'>");
                        goto Passed;
                    }
                    Passed:
                    _builder.Append(r[c.ColumnName]);
                    build:
                    _builder.Append("</td>");
                }
                _builder.Append("</tr>");
            }
            //Closing Tags
            _builder.Append("</table>");
            _builder.Append("</body>");
            _builder.Append("</html>");
            HTMLTEXT = _builder.ToString();

            //Path is combined along with the file name
            string path = Path.Combine(Constants.globalResultsPath, "Log" + ".htm");
            //Creating the directory to store Log files
            Utility.createDirectory();
            if (!File.Exists(path))
            {
                var resultfile = File.Create(path);
                resultfile.Close();
                goto write;
            }
            else
            {
                goto write;
            }
            write:
            File.AppendAllText(@path, HTMLTEXT);
            Constants.globalLogPath = path;
        }

        //public static void LogListToDataTable(List<Tuple<String, String, String, String, String>> logTuple)
        //{
        //    DataTable dataTableShort = new DataTable();
        //    dataTableShort.Columns.Add("ProjectName", typeof(string));
        //    dataTableShort.Columns.Add("CodedUIMethod", typeof(string));
        //    dataTableShort.Columns.Add("TestMethod", typeof(string));
        //    dataTableShort.Columns.Add("Results", typeof(string));
        //    dataTableShort.Columns.Add("TestUser", typeof(string));
        //    dataTableShort.Columns.Add("TimeStamp", typeof(string));
        //   // dataTableShort.Columns.Add("DataRow", typeof(string));
        //    dataTableShort.Columns.Add("Message", typeof(string));

        //    // Here we add all the DataRows
        //    for (int i = 0; i <logTuple.Count; i++)
        //    {
        //        dataTableShort.Rows.Add(logTuple[i].Item1,logTuple[i].Item2, logTuple[i].Item3, logTuple[i].Item4, logTuple[i].Item5);
        //    }
        //    string htmlText = LogFileHtmlBuilder(dataTableShort);
        //    //Path is combined along with the file name 
        //    string path = Path.Combine(Constants.globalResultsPath, "Log1" + ".html");
        //    //Creating the directory to store Log files
        //    Utility.createDirectory();

        //    if (!File.Exists(path))
        //    {
        //        var resultfile = File.Create(path);
        //        resultfile.Close();
        //        goto write;
        //    }
        //    else
        //    {
        //        goto write;
        //    }
        //    write:
        //    File.AppendAllText(@path, HTMLTEXT);
        //    Constants.globalLogPath = path;
        //}
        //public static string LogFileHtmlBuilder(DataTable dt)
        //{
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
        //    builder.Append("<img src = 'C:\\AutomationTesting\\AutomationTestScripts\\NGW_SharePoint\\NGW_SharePoint\\BoschLogo.png'/>");
        //    //builder.Append("<img src='data: image/png; base64,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' alt='BoschLogo.png'/>");
        //    //builder.Append("<div style='border - top:3px solid #22BCE5'>&nbsp;</div>");
        //    //builder.Append("<span style = 'font-family:Arial;font-size:10pt'>");
          
        //    // builder.Append("<table cellpadding='5' cellspacing='0' style='border: solid 1px Silver; font-size:small;table-layout:fixed;font-family:arial,sans-serif;'> ");
        //    builder.Append("<table style = 'width:100%' border = '1'>");
        //    // builder.Append("<tr>");
        //    builder.Append("<th>" + "ProjectName" + "</th>");
        //    builder.Append("<th>" + "CodedUIMethod" + "</th>");
        //    builder.Append("<th>" + "TestMethod" + "</th>");
        //    builder.Append("<th>" + "Results" + "</th>");
        //    builder.Append("<th>" + "TestUser" + "</th>");
        //    builder.Append("<th>" + "TimeStamp" + "</th>");
        //    builder.Append("<th>" + "Message" + "</th>");

        //    foreach (DataRow r in dt.Rows)
        //    {
        //        builder.Append("<tr>");

        //        foreach (DataColumn c in dt.Columns)
        //        {
        //            //Below If condition is used to check whether the Status column value is Failed If so then Red colour will be given to that particular cell
        //            if (r[c.ColumnName].ToString() == "Failed")
        //            {
        //                builder.Append("<td style='word -break:break-all'>");
        //                builder.Append("<font color='#660000'>");
        //                builder.Append(r[c.ColumnName]);
        //                builder.Append("</font>");
        //                goto build;
        //            }
        //            else if (r[c.ColumnName].ToString() == "Inconclusive")
        //            {

        //                builder.Append("<td style='word -break:break-all'>");
        //                builder.Append("<font color='663300'>");
        //                builder.Append(r[c.ColumnName]);

        //                builder.Append("</font>");
        //                goto build;

        //            }
        //            else if (r[c.ColumnName].ToString() == "Passed")
        //            {

        //                builder.Append("<td style='word -break:break-all'>");
        //                builder.Append("<font color='#006600'>");
        //                builder.Append(r[c.ColumnName]);

        //                builder.Append("</font>");
        //                goto build;

        //            }
        //            else
        //            {
        //                builder.Append("<td style='word -break:break-all'>");
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

    }

    }

