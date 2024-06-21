
# XlsTypewriter
ExcelTypewriter is .NET library for writing Excel files built on top of [ClosedXML](https://github.com/ClosedXML/ClosedXML). It aims to provide intuitive and user friendly interface, simulating typewriter. User doesn't have to worry about cell names. Instead of writing cell names him/herself (e.g. "C3"), user navigates excel file using functions Print, Skip and NextRow.
## Installation
Package can be installed as Nuget package.

```
dotnet add package XlsTypewriter
```
## What can you do with this?
XlsTypewriter allows you to write Excel file without keeping track of cell names. 

**Example:**
```  
using XlsTypewriter;  
  
var workbook = new Workbook();  
var worksheet = workbook.AddWorksheet("Sheet1");

for (var i = 1; i < 4; i++) {
	for (var j = 1; j < 3; j++) { 
		worksheet.Print(i * j);
	}
	worksheet.NewRow();
}

workbook.SaveAs("Example.xlsx");
```
## Merging cells
```
worksheet.Merge(columns:2, rows:1); // merge current cell with next cell in row, getting 2*1
worksheet.Merge(2, 2); // merge current cell with cells in next row and in next column, getting 2x2 cell
```
## Style
Style can be set for current cell or for range of cells.
```
var style = worksheet.GetStyle();
style.Font.Bold = true;
worksheet.SetStyle(style, 2, 1); // set style for cell in 2nd column and 1st row
```
Style can be set using predefined styles.
```
worksheet.SetBoxBorders(XLBorderStyleValues.Thick); // set thick border around current cell
worksheet.SetBordersColor(XLColor.Blue, 2, 1); // set blue color for borders current cell and next cell in row
worksheet.SetVerticalBorders(XLBorderStyleValues.Thin, 1, 2); // set thin vertical borders for current cell and next cell in column
worksheet.SetHorizontalBorders(XLBorderStyleValues.Medium);
worksheet.SetFontColor(XLColor.Amber);
worksheet.SetBackgroundColor(XLColor.Bistre);
worksheet.SetFontName("Arial");
worksheet.SetFontSize(12);
worksheet.SetFontItalic();
worksheet.SetFontUnderline();
worksheet.SetVerticalAlignment(XLAlignmentVerticalValues.Center);
worksheet.SetHorizontalAlignment(XLAlignmentHorizontalValues.Center);
worksheet.SetBordersColor(XLColor.Blue);
worksheet.SetTopBorder(XLBorderStyleValues.Thin);
worksheet.SetBottomBorder(XLBorderStyleValues.Thin);
worksheet.SetLeftBorder(XLBorderStyleValues.Thin);
worksheet.SetRightBorder(XLBorderStyleValues.Thin);
```
**TODO**:: Add style for whole row and column

## Navigation
Beside regular Print, Skip and NewRow navigating through file, you can use 
```
GoTo(int column, int row) // go to cell with given column and row
GoToStart() // go to begin of file, or "A1" cell
GoToEnd() // go to last used cell in file
 ```

## Formula

*TO DO*