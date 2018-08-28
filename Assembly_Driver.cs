using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NGW_SharePoint.Utility;
using System.Drawing;

namespace NGW_SharePoint.DriverFunctions.Assembly
{
    /// <summary>
    /// Summary description for CodedUITest1
    /// </summary>
    [CodedUITest]
    public class Assembly_Driver
    {   
        [AssemblyCleanup]
        public static void Test()
        {

            TestRunChart.Chart_Creation(Constants.totalRun, Constants.passCount, Constants.failCount, Constants.errorCount, Constants.othersCount);
            Utility.Utility.excelAddButtonWithVBA();
            Utility.Utility.ListToDataTableConverter();
            Reports.CreatePieChart(Constants.passCount, Constants.failCount, Constants.othersCount);
            Utility.Utility.ConvertExceltoImage();
            Utility.Utility.ConvertExceltoHtml();
            Reports.WriteResultToNotepad();
            Utility.Utility.DirectoryCopyToNewDirectory(Constants.globalResultsPath, Constants.globalRecentResultsPath, true);
            //Email.sendMail();
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
        //public void MyTestCleanup()
        //{        
        //    // To generate code for this test, select "Generate Code for Coded UI Test" from the shortcut menu and select one of the menu items.
        //}

        #endregion

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
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
    }
}
