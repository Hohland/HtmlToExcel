using HtmlAgilityPack;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.Linq;
using TowerSoft.HtmlToExcel;

namespace HtmlToExcelTests.Utilities
{
    [TestClass]
    public class EpplusUtilitiesTest
    {
        private const string SheetName = "sheetName";

        [TestMethod]
        public void AddBackgroundToCellNode()
        {
            var htmlWithTable = "<div><table><thead></thead><tbody><tr><td style=\"background-color: #bbbbbb;\">Test</td></tr></tbody></table></div>";
            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(htmlWithTable);

            HtmlNode node = new HtmlAgilityUtilities().GetHtmlTableNode(htmlDoc);
            var package = new ExcelPackage();
            new EPPlusUtilities(new HtmlToExcelSettings()).CreateSheet(package, SheetName, node);
            var sheet = package.Workbook.Worksheets[SheetName];
            var cell = sheet.Cells.First();
            if (cell.Style.Fill.BackgroundColor.Rgb != "01BBBBBB") 
            {
                Assert.Fail("The first cell is not filled.");
            }
        }
    }
}
