using System.IO;
using OfficeOpenXml;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EPPlusCRIssueTest
{
    [TestClass]
    public class CrIssueReproducer
    {
        [TestMethod]
        public void CrIssueTest()
        {
            const string lfLineEndingExpected = "Test LF line ending \n",
                         crLfLineEndingExpected = "Test CR LF line ending \r\n", 
                         crLineEndingExpected = "Test CR line ending \r";

            FileInfo newFile = new FileInfo("CRIssue.xlsx");
            if (newFile.Exists)
            {
                newFile.Delete();  // ensures we create a new workbook
                newFile = new FileInfo("CRIssue.xlsx");
            }
            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Inventory");
                worksheet.Cells["A1"].Value = lfLineEndingExpected;
                worksheet.Cells["A2"].Value = crLfLineEndingExpected;
                worksheet.Cells["A3"].Value = crLineEndingExpected;
                package.Save();
            }

            newFile = new FileInfo("CRIssue.xlsx");
            string lfLineEnding, crLfLineEnding, crLineEnding;

            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                var worksheet = package.Workbook.Worksheets["Inventory"];
                lfLineEnding = worksheet.Cells["A1"].Value.ToString();
                crLfLineEnding = worksheet.Cells["A2"].Value.ToString();
                crLineEnding = worksheet.Cells["A3"].Value.ToString();
            }
            newFile.Delete();

            Assert.AreEqual(lfLineEndingExpected, lfLineEnding);
            Assert.AreEqual(crLfLineEndingExpected, crLfLineEnding);
            Assert.AreEqual(crLineEndingExpected, crLineEnding);
        }
    }
}
