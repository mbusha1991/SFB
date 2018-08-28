using System.Collections.Generic;
using SFB.LibraryFunctions;
namespace NGW_SharePoint.Variables
{
    class Variables
    {
        //Chart variables
        public static double chartDoubleLeft = 100;
        public static double chartDoubleTop = 20;
        public static double chartDoubleWidth = 300;
        public static double chartDoubleHeight = 200;


        public static int wsCellsMergeStartObjectRowIndex = 3;
        public static int wsCellsMergeStartObjectColumnIndex = 12;
        public static int wsCellsMergeEndObjectRowIndex = wsCellsMergeStartObjectRowIndex;
        public static int wsCellsMergeEndObjectColumnIndex = wsCellsMergeStartObjectColumnIndex + 1;

        public static string headerRange = "L3:L3";
        public static string headerRangeValue = "Test Run Status";
        public static System.Drawing.Color headerRangeColor = System.Drawing.Color.White;
        public static System.Drawing.Color headerRangeInteriorColor = System.Drawing.Color.OrangeRed;

        public static int wsCellsTableRow1Column1StartObjectRowIndex = 4;
        public static int wsCellsTableRow1Column1StartObjectColumnIndex = 12;
        public static int wsCellsTableRow2Column1StartObjectRowIndex = wsCellsTableRow1Column1StartObjectRowIndex + 1;
        public static int wsCellsTableRow3Column1StartObjectRowIndex = wsCellsTableRow2Column1StartObjectRowIndex + 1;
        public static int wsCellsTableRow4Column1StartObjectRowIndex = wsCellsTableRow3Column1StartObjectRowIndex + 1;

        public static int wsCellsTableRow1Column2EndObjectRowIndex = wsCellsTableRow1Column1StartObjectRowIndex;
        public static int wsCellsTableRow2Column2EndObjectRowIndex = wsCellsTableRow2Column1StartObjectRowIndex;
        public static int wsCellsTableRow3Column2EndObjectRowIndex = wsCellsTableRow3Column1StartObjectRowIndex;
        public static int wsCellsTableRow4Column2EndObjectRowIndex = wsCellsTableRow4Column1StartObjectRowIndex;
        public static int wsCellsTableRow1Column2EndObjectColumnIndex = wsCellsTableRow1Column1StartObjectColumnIndex + 1;

        public static string rowHeader1 = "Total";
        public static string rowHeader2 = "Skipped";
        public static string rowHeader3 = globalFunctions._fail;
        public static string rowHeader4 = globalFunctions._pass;

        public static string fullTableRangeStart = "L4";
        public static string fullTableRangeEnd = "M7";
        public static string fullTableRangeFontName = "Arial";
        public static int fullTableRangeFontSize = 9;
        public static string pieChartTableContentRangeStart = "L5";
        public static string pieChartTableContentRangeEnd = fullTableRangeEnd;

        //Class Email variable
        public static string mailSubject = "Test Automation Results- " + Utility.Utility.GetCurrentDate();
        public static string smtpServerName = "rb-smtp-int.bosch.com";
        public static string mailFrom = "manojkumar.munuswamy@in.bosch.com";

    }
}
