// See https://aka.ms/new-console-template for more information

using ClosedXML.Excel;
using XlsTypewriter;

var workbook = new Workbook();
var worksheet = workbook.AddWorksheet("aaa");

worksheet = workbook.GetWorksheet("aaa");

worksheet.Print("Hello");
worksheet.Print("World");
worksheet.NewRow();

var style = worksheet.GetStyle();
style.Font.Bold = true;
worksheet.SetStyle(style, 2, 1);
worksheet.Print("Hello2");
worksheet.Print("World2");
worksheet.NewRow();

worksheet.Merge(2, 1);
worksheet.Print("Hello3 World3");
worksheet.NewRow();

worksheet.SetBoxBorders(XLBorderStyleValues.Thick);
worksheet.SetBordersColor(XLColor.Blue);
worksheet.Print("Hello4");

worksheet.SetVerticalBorders(XLBorderStyleValues.Thin);
worksheet.Print("World4");

worksheet.SetHorizontalBorders(XLBorderStyleValues.Medium);
worksheet.SetFontColor(XLColor.Amber);
worksheet.SetBackgroundColor(XLColor.Bistre);
worksheet.Print("World4");

workbook.SaveAs("./HelloWorld.xlsx");