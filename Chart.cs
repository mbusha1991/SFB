using System;
using ExcelQ = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.IO;
namespace NGW_SharePoint.Utility
{
    public partial class TestRunChart
    {
        private object missing = Type.Missing;

        static Application xla;
        public static void Chart_Creation(int totalrun, int totalpass, int totalfail, int errors, int skipped )
        {
            xla = new Application();
            Workbook wb = xla.Workbooks.Open(Constants.mREPORTPATH, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Worksheet ws = (Worksheet)xla.ActiveSheet;
            Range range;

            // Now create the chart.
            ChartObjects chartObjs = (ChartObjects)ws.ChartObjects(Type.Missing);            
            ChartObject chartObj = chartObjs.Add(Variables.Variables.chartDoubleLeft, Variables.Variables.chartDoubleTop, Variables.Variables.chartDoubleWidth, Variables.Variables.chartDoubleHeight);
            ExcelQ.Chart xlChart = chartObj.Chart;

            ws.Range[ws.Cells[Variables.Variables.wsCellsMergeStartObjectRowIndex, Variables.Variables.wsCellsMergeStartObjectColumnIndex], ws.Cells[Variables.Variables.wsCellsMergeEndObjectRowIndex, Variables.Variables.wsCellsMergeEndObjectColumnIndex]].Merge();
            Range headerRange = ws.get_Range(Variables.Variables.headerRange);
            headerRange.Value = Variables.Variables.headerRangeValue;
            headerRange.Font.Color = Variables.Variables.headerRangeColor;
            headerRange.Font.Bold = true;
            headerRange.Cells.HorizontalAlignment= XlHAlign.xlHAlignCenter;
            headerRange.Interior.Color = Variables.Variables.headerRangeInteriorColor;
            //Adding table row and column
            ws.Cells[Variables.Variables.wsCellsTableRow1Column1StartObjectRowIndex, Variables.Variables.wsCellsTableRow1Column1StartObjectColumnIndex] = Variables.Variables.rowHeader1;
            ws.Cells[Variables.Variables.wsCellsTableRow1Column1StartObjectRowIndex, Variables.Variables.wsCellsTableRow1Column1StartObjectColumnIndex].Font.Bold = true;
            ws.Cells[Variables.Variables.wsCellsTableRow2Column1StartObjectRowIndex, Variables.Variables.wsCellsTableRow1Column1StartObjectColumnIndex] = Variables.Variables.rowHeader2;
            ws.Cells[Variables.Variables.wsCellsTableRow2Column1StartObjectRowIndex, Variables.Variables.wsCellsTableRow1Column1StartObjectColumnIndex].Font.Bold = true;
            ws.Cells[Variables.Variables.wsCellsTableRow3Column1StartObjectRowIndex, Variables.Variables.wsCellsTableRow1Column1StartObjectColumnIndex] = Variables.Variables.rowHeader3;
            ws.Cells[Variables.Variables.wsCellsTableRow3Column1StartObjectRowIndex, Variables.Variables.wsCellsTableRow1Column1StartObjectColumnIndex].Font.Bold = true;
            ws.Cells[Variables.Variables.wsCellsTableRow4Column1StartObjectRowIndex, Variables.Variables.wsCellsTableRow1Column1StartObjectColumnIndex] = Variables.Variables.rowHeader4;
            ws.Cells[Variables.Variables.wsCellsTableRow4Column1StartObjectRowIndex, Variables.Variables.wsCellsTableRow1Column1StartObjectColumnIndex].Font.Bold = true;
            ws.Cells[Variables.Variables.wsCellsTableRow1Column2EndObjectRowIndex, Variables.Variables.wsCellsTableRow1Column2EndObjectColumnIndex] = totalrun;
            ws.Cells[Variables.Variables.wsCellsTableRow2Column2EndObjectRowIndex, Variables.Variables.wsCellsTableRow1Column2EndObjectColumnIndex] = skipped;
            ws.Cells[Variables.Variables.wsCellsTableRow3Column2EndObjectRowIndex, Variables.Variables.wsCellsTableRow1Column2EndObjectColumnIndex] = totalfail;
            ws.Cells[Variables.Variables.wsCellsTableRow4Column2EndObjectRowIndex, Variables.Variables.wsCellsTableRow1Column2EndObjectColumnIndex] = totalpass;
      
            range = ws.UsedRange;
            Range fullTableRange = ws.get_Range(Variables.Variables.fullTableRangeStart, Variables.Variables.fullTableRangeEnd);
            Range pieChartContentTableRange = ws.get_Range(Variables.Variables.pieChartTableContentRangeStart, Variables.Variables.pieChartTableContentRangeEnd);         
            fullTableRange.Style.Font.Name = Variables.Variables.fullTableRangeFontName;
            fullTableRange.Style.Font.Size = Variables.Variables.fullTableRangeFontSize;
            fullTableRange.Style.Interior.Pattern = XlPattern.xlPatternSolid;
            BorderAround(fullTableRange, 000000);
            xlChart.ChartType = XlChartType.xlPie;
            xlChart.ChartArea.Interior.Color = Color.WhiteSmoke;
            xlChart.ApplyLayout(1);
            xlChart.ChartTitle.Text = Variables.Variables.headerRangeValue;
            xlChart.SetSourceData(pieChartContentTableRange, Type.Missing);
            xlChart.SeriesCollection(1).Format.Fill.Transparency = 1;

            xlChart.ApplyDataLabels(ExcelQ.XlDataLabelsType.xlDataLabelsShowLabelAndPercent,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing );

            ws.Shapes.Item("Chart 1").Top = 130;
            ws.Shapes.Item("Chart 1").Left = 1150;
            wb.Save();
            wb.Close(true, Type.Missing, Type.Missing);
            xla.Quit();
            releaseObject(ws);
            releaseObject(wb);
            releaseObject(xla);
            string path = Path.Combine(Constants.globalRecentResultsPath, " Automation_Test_Report" + ".xlsm");
            //Creating the directory to store Log files
            if (!Directory.Exists(Constants.globalRecentResultsPath))
            {

                Directory.CreateDirectory(Constants.globalRecentResultsPath);

            }
            File.Copy(Constants.mREPORTPATH, Constants.RecentREPORTPATH, true);
        }
        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);

                obj = null;
            }

            catch (Exception ex)
            {
                obj = null;
               Console.WriteLine("Exception Occured while releasing object " +ex.ToString());
            }

            finally
            {
                GC.Collect();
            }
        }
        private static void BorderAround(ExcelQ.Range range, int colour)
        {
            ExcelQ.Borders borders = range.Borders;
            borders[ExcelQ.XlBordersIndex.xlEdgeLeft].LineStyle = ExcelQ.XlLineStyle.xlContinuous;
            borders[ExcelQ.XlBordersIndex.xlEdgeTop].LineStyle = ExcelQ.XlLineStyle.xlContinuous;
            borders[ExcelQ.XlBordersIndex.xlEdgeBottom].LineStyle = ExcelQ.XlLineStyle.xlContinuous;
            borders[ExcelQ.XlBordersIndex.xlEdgeRight].LineStyle = ExcelQ.XlLineStyle.xlContinuous;
            borders.Color = colour;           
            borders = null;
        }
        public static object ChartType { get; private set; }
    }
}
