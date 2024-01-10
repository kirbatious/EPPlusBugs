using System;
using System.IO;
using EPPlusShared;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusNewTests
{
    [TestClass]
    public class CalculationTests
    {
        [ClassInitialize]
        public static void ClassInitialize(TestContext _)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        /// <summary>
        /// This test passes in 4.5.3.2 but not in 7.0.5
        /// </summary>
        [TestMethod]
        public void OpenFileWithConditionalFormattingTest()
        {
            var fi = Helper.GetFileInfo("YourResultsPart.xlsm");
            Assert.IsTrue(fi.Exists);

            // Open the file
            var package = new ExcelPackage(fi);
            Assert.IsNotNull(package.Workbook);
        }
        
        /// <summary>
        /// This test passes in 4.5.3.2 but not in 7.0.5
        /// </summary>
        [TestMethod]
        public void CalculationTest()
        {
            // Get hold of ExcelApplicationTestCalculations.xlsx in the shared project
            var fi = Helper.GetFileInfo("ExcelApplicationTestCalculations.xlsx");
            Assert.IsTrue(fi.Exists);

            // Open the file
            var package = new ExcelPackage(fi);
            var wbk = package.Workbook;
            var sht = wbk.Worksheets["TestSheet"];

            // Call calculate
            wbk.Calculate();

            // Check everything is initially in order
            Assert.AreEqual(1.0, sht.Cells["B3"].Value);
            Assert.AreEqual(2.0, sht.Cells["C3"].Value);
            Assert.AreEqual(2.0, sht.Cells["B4"].Value);
            Assert.AreEqual(4.0, sht.Cells["C4"].Value);

            // Update the value of two cells
            sht.Cells["B3"].Value = 500.0;
            sht.Cells["B4"].Value = 500.0;

            // Call calculate again
            wbk.Calculate();

            // C3 and C4 have formulae in them which double the value to their left
            Assert.AreEqual(1000.0, sht.Cells["C3"].Value);
            Assert.AreEqual(1000.0, sht.Cells["C4"].Value);
        }

        /// <summary>
        /// This test passes in both versions - I only hit the issue in existing spreadsheets created in regular MS Excel.
        /// </summary>
        [TestMethod]
        public void NewExcelFileCalculationTest()
        {
            // Open the file
            var package = new ExcelPackage();
            var wbk = package.Workbook;
            var sht = wbk.Worksheets.Add("TestSheet");
            
            // Set up
            sht.Cells["B3"].Value = 1.0;
            sht.Cells["B4"].Value = 2.0;

            sht.Cells["C3"].Formula = "B3*2";
            sht.Cells["C4"].Formula = "B4*2";

            // Call calculate
            wbk.Calculate();

            var path = $"{Environment.CurrentDirectory}\\Test.xlsx";
            if(File.Exists(path)) File.Delete(path);
            var fi = new FileInfo(path);
            package.SaveAs(fi);

            Assert.IsTrue(fi.Exists);

            package = new ExcelPackage(fi);
            wbk = package.Workbook;
            sht = wbk.Worksheets["TestSheet"];

            // Call calculate
            wbk.Calculate();

            Assert.AreEqual(1.0, sht.Cells["B3"].Value);
            Assert.AreEqual(2.0, sht.Cells["C3"].Value);
            Assert.AreEqual(2.0, sht.Cells["B4"].Value);
            Assert.AreEqual(4.0, sht.Cells["C4"].Value);

            // Update the value of two cells
            sht.Cells["B3"].Value = 500.0;
            sht.Cells["B4"].Value = 500.0;

            // Call calculate again
            wbk.Calculate();

            // C3 and C4 have formulae in them which double the value to their left
            Assert.AreEqual(1000.0, sht.Cells["C3"].Value);
            Assert.AreEqual(1000.0, sht.Cells["C4"].Value);
        }
    }
}