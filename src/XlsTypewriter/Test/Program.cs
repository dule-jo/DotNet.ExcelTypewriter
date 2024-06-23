// See https://aka.ms/new-console-template for more information

using ClosedXML.Excel;
using XlsTypewriter;

using var workbook = new Workbook();
workbook.Properties.Title = "Hello World";
workbook.CustomProperties.Add("Author2", "John Doe");
var worksheet = workbook.AddWorksheet("aaa");

worksheet = workbook.GetWorksheet("aaa");

var wsstyle = worksheet.GetStyle();
wsstyle.Font.Bold = true;
worksheet.SetStyleToWorksheet(wsstyle);

worksheet.Print("Hello");

var nextStyle = worksheet.GetStyle();
nextStyle.Font.Bold = false;
worksheet.SetStyle(nextStyle);
worksheet.Print("World");
worksheet.FreezeColumns();
worksheet.NewRow();
worksheet.FreezeRows();


// worksheet.HideColumn();
// worksheet.HideRow();

wsstyle.Font.Bold = false;
worksheet.SetStyleToWorksheet(wsstyle);

var style = worksheet.GetStyle();
style.Font.Bold = true;
worksheet.SetStyle(style, 1, 2);
worksheet.Print("Hello2");
worksheet.Print("World2");
worksheet.NewRow();

worksheet.Merge(1, 2);
worksheet.Print("Hello3 World3");
worksheet.NewRow();
worksheet.PageSetup.Header.Left.AddText("Created with XlsTypewriter");


worksheet.SetBoxBorders(XLBorderStyleValues.Thick);
worksheet.SetBordersColor(XLColor.Blue);
worksheet.Print("Hello4");

worksheet.SetVerticalBorders(XLBorderStyleValues.Thin);
worksheet.Print("World4");

worksheet.SetHorizontalBorders(XLBorderStyleValues.Medium);
worksheet.SetFontColor(XLColor.Amber);
worksheet.SetBackgroundColor(XLColor.Bistre);
worksheet.Print("World4");

worksheet = workbook.AddWorksheet("Sheet2");
worksheet.Print("Hello");
worksheet.Print("World");
var style2 = worksheet.GetStyle();
style2.Fill.BackgroundColor = XLColor.Blue;
worksheet.SetRowStyle(style2);
worksheet.NewRow();

worksheet.Print("Hello");
worksheet.Print("world");
var style3 = worksheet.GetStyle();
style3.Fill.BackgroundColor = XLColor.Red;
worksheet.SetRowStyle(style3, 2);
worksheet.NewRow();

worksheet.Print("Hello");
worksheet.Print("world");

worksheet = workbook.AddWorksheet("Sheet3");

var style4 = worksheet.GetStyle();
style4.Fill.BackgroundColor = XLColor.Blue;
worksheet.SetColumnStyle(style4);
worksheet.Print("Hello");

var style5 = worksheet.GetStyle();
style5.Fill.BackgroundColor = XLColor.Red;
worksheet.SetColumnStyle(style5, 2);
worksheet.Print("Whole");
worksheet.Print("World");

worksheet.NewRow();

worksheet.Print("Hello");
worksheet.Print("Whole");
worksheet.Print("World, this is very long column");

worksheet.AdjustColumnWidth();
worksheet.AdjustRowHeight();

workbook.SaveAs("./HelloWorld.xlsx");