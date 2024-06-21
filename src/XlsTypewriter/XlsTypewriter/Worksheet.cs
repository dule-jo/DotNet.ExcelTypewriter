using ClosedXML.Excel;
using XlsTypewriter.Common;

namespace XlsTypewriter;

public class Worksheet(IXLWorksheet worksheet)
{
    private int _rowIndex = 1;

    private int _columnIndex = 1;

    private IXLCell CurrentCell => worksheet.Cell(_rowIndex, _columnIndex);

    public void GoTo(int columnIndex, int rowIndex)
    {
        _rowIndex = rowIndex;
        _columnIndex = columnIndex;
    }

    public void GoToStart() => GoTo(1, 1);

    public void GoToEnd() => GoTo(worksheet.LastRowUsed().RowNumber(), worksheet.LastColumnUsed().ColumnNumber());

    public void Print(object value)
    {
        CurrentCell.Value = XLCellValue.FromObject(value);
        _columnIndex++;
    }

    public void NewRow()
    {
        _rowIndex++;
        _columnIndex = 1;
    }

    public void SkipColumn(int count = 1) => _columnIndex += count;

    private IXLRange GetRange(int rows, int columns) => worksheet.Range(
        CurrentCell,
        worksheet.Cell(_rowIndex + rows - 1, _columnIndex + columns - 1)
    );

    public IXLStyle GetStyleFromWorksheet() => worksheet.Style;

    public void SetStyleToWorksheet(IXLStyle style) => worksheet.Style = style;

    public void SetRowStyle(IXLStyle style, int rows = 1)
    {
        for (var i = 0; i < rows; i++) worksheet.Row(_rowIndex + i).Style = style;
    }

    public void SetColumnStyle(IXLStyle style, int columns = 1)
    {
        for (var i = 0; i < columns; i++) worksheet.Column(_columnIndex + i).Style = style;
    }

    public void Merge(int rows, int columns) => GetRange(rows, columns).Merge();

    public IXLStyle GetStyle() => CurrentCell.Style;

    public void SetStyle(IXLStyle style, int rows = 1, int columns = 1) => GetRange(rows, columns).Style = style;

    public void SetBoldStyle(int rows = 1, int columns = 1)
    {
        var style = GetStyle();
        style.Font.Bold = true;
        GetRange(rows, columns).Style = style;
    }

    public void SetFontName(string fontName, int rows = 1, int columns = 1)
    {
        var style = GetStyle();
        style.Font.FontName = fontName;
        GetRange(rows, columns).Style = style;
    }

    public void SetBoxBorders(XLBorderStyleValues borderStyle, int rows = 1, int columns = 1)
    {
        var style = GetStyle();
        style.Border.BottomBorder = borderStyle;
        style.Border.TopBorder = borderStyle;
        style.Border.RightBorder = borderStyle;
        style.Border.LeftBorder = borderStyle;
        GetRange(rows, columns).Style = style;
    }

    public void SetVerticalBorders(XLBorderStyleValues borderStyle, int rows = 1, int columns = 1)
    {
        var style = GetStyle();
        style.Border.BottomBorder = borderStyle;
        style.Border.TopBorder = borderStyle;
        GetRange(rows, columns).Style = style;
    }

    public void SetHorizontalBorders(XLBorderStyleValues borderStyle, int rows = 1, int columns = 1)
    {
        var style = GetStyle();
        style.Border.RightBorder = borderStyle;
        style.Border.LeftBorder = borderStyle;
        GetRange(rows, columns).Style = style;
    }

    public void SetTopBorder(XLBorderStyleValues borderStyle, int rows = 1, int columns = 1)
    {
        var style = GetStyle();
        style.Border.TopBorder = borderStyle;
        GetRange(rows, columns).Style = style;
    }

    public void SetBottomBorder(XLBorderStyleValues borderStyle, int rows = 1, int columns = 1)
    {
        var style = GetStyle();
        style.Border.BottomBorder = borderStyle;
        GetRange(rows, columns).Style = style;
    }

    public void SetLeftBorder(XLBorderStyleValues borderStyle, int rows = 1, int columns = 1)
    {
        var style = GetStyle();
        style.Border.LeftBorder = borderStyle;
        GetRange(rows, columns).Style = style;
    }

    public void SetRightBorder(XLBorderStyleValues borderStyle, int rows = 1, int columns = 1)
    {
        var style = GetStyle();
        style.Border.RightBorder = borderStyle;
        GetRange(rows, columns).Style = style;
    }

    public void SetBackgroundColor(XLColor color, int rows = 1, int columns = 1)
    {
        var style = GetStyle();
        style.Fill.BackgroundColor = color;
        GetRange(rows, columns).Style = style;
    }

    public void SetFontColor(XLColor color, int rows = 1, int columns = 1)
    {
        var style = GetStyle();
        style.Font.FontColor = color;
        GetRange(rows, columns).Style = style;
    }

    public void SetFontSize(double size, int rows = 1, int columns = 1)
    {
        var style = GetStyle();
        style.Font.FontSize = size;
        GetRange(rows, columns).Style = style;
    }

    public void SetFontUnderline(XLFontUnderlineValues underline, int rows = 1, int columns = 1)
    {
        var style = GetStyle();
        style.Font.Underline = underline;
        GetRange(rows, columns).Style = style;
    }

    public void SetFontItalic(bool italic, int rows = 1, int columns = 1)
    {
        var style = GetStyle();
        style.Font.Italic = italic;
        GetRange(rows, columns).Style = style;
    }

    public void SetVerticalAlignment(XLAlignmentVerticalValues alignment, int rows = 1, int columns = 1)
    {
        var style = GetStyle();
        style.Alignment.Vertical = alignment;
        GetRange(rows, columns).Style = style;
    }

    public void SetHorizontalAlignment(XLAlignmentHorizontalValues alignment, int rows = 1, int columns = 1)
    {
        var style = GetStyle();
        style.Alignment.Horizontal = alignment;
        GetRange(rows, columns).Style = style;
    }

    public void SetBordersColor(XLColor color, int rows = 1, int columns = 1)
    {
        var style = GetStyle();
        style.Border.BottomBorderColor = color;
        style.Border.TopBorderColor = color;
        style.Border.RightBorderColor = color;
        style.Border.LeftBorderColor = color;
        GetRange(rows, columns).Style = style;
    }

    public void AdjustColumnWidth()
    {
        worksheet.ColumnsUsed().AdjustToContents();
    }
    
    public void AdjustCurrentColumnWidth()
    {
        worksheet.Column(_columnIndex).AdjustToContents();
    }
    
    public void AdjustRowHeight()
    {
        worksheet.RowsUsed().AdjustToContents();
    }
    
    public void AdjustCurrentRowHeight()
    {
        worksheet.Row(_rowIndex).AdjustToContents();
    }

    public string Name => worksheet.Name;
}