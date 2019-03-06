using System.Diagnostics;
using System.IO;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlus.Html.Test
{
    [TestClass]
    public class HtmlTests
    {
        [TestMethod]
        public void ExcelToHtml()
        {
            FileInfo test001 = new FileInfo(@"C:\Users\adkerti\Downloads\2019 HDCC ECM Cummins V1a Round 2 as of 03-06-2019 2.46.30 PM.xlsx");
            var package = new ExcelPackage(test001);
            var worksheet = package.Workbook.Worksheets[1];

            var html = worksheet.ToHtml(new ExcelAddress(2, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column));

            Show(html);
        }

        private void Show(string html)
        {
            var tmpFile = Path.GetTempFileName() + ".html";
            File.WriteAllText(tmpFile, html);
            Process.Start(tmpFile);
        }
    }
}