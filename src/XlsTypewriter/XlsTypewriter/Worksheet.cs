using ClosedXML.Excel;
using XlsTypewriter.Common;

namespace XlsTypewriter;

public class Worksheet
{
    private IXLWorksheet _worksheet;

    private int _rowIndex = 1;

    private int _columnIndex = 1;

    private string ColumnName() => XlsHelper.IntToColumnName(_columnIndex);

    private static string ColumnName(int index) => XlsHelper.IntToColumnName(index);

    private string CurrentCellName() => $"{ColumnName()}{_rowIndex}";

    private static string CellName(int columnIndex, int rowIndex) => $"{ColumnName(columnIndex)}{rowIndex}";

    public Worksheet(IXLWorksheet worksheet)
    {
        _worksheet = worksheet;
    }

    public void GoTo(int columnIndex, int rowIndex)
    {
        _rowIndex = rowIndex;
        _columnIndex = columnIndex;
    }

    public void GoToStart() => GoTo(1, 1);

    public void GoToEnd() => GoTo(_worksheet.LastRowUsed().RowNumber(), _worksheet.LastColumnUsed().ColumnNumber());

    public void Print(object value)
    {
        _worksheet.Cell(CurrentCellName()).Value = XLCellValue.FromObject(value);
        _columnIndex++;
    }

    public void NewRow()
    {
        _rowIndex++;
        _columnIndex = 1;
    }

    public void SkipColumn(int count = 1) => _columnIndex += count;

    private IXLRange GetRange(int rows, int columns) => _worksheet.Range(
        CurrentCellName(),
        CellName(_columnIndex + (columns - 1), _rowIndex + (rows - 1))
    );

    public IXLStyle GetStyleFromWorksheet() => _worksheet.Style;

    public void SetStyleToWorksheet(IXLStyle style) => _worksheet.Style = style;

    public void SetRowStyle(IXLStyle style, int rows = 1)
    {
        for (var i = 0; i < rows; i++) _worksheet.Row(_rowIndex + i).Style = style;
    }
    
    public void SetRowStyle(IXLStyle style, int[] rows)
    {
        foreach (var row in rows) _worksheet.Row(row).Style = style;
    }

    public void SetColumnStyle(IXLStyle style, int columns = 1)
    {
        for (var i = 0; i < columns; i++) _worksheet.Column(_columnIndex + i).Style = style;
    }
    
    public void SetColumnStyle(IXLStyle style, int[] columns)
    {
        foreach (var column in columns) _worksheet.Column(column).Style = style;
    }

    public void Merge(int columns, int rows) => GetRange(rows, columns).Merge();

    public IXLStyle GetStyle() => _worksheet.Cell(CurrentCellName()).Style;

    public void SetStyle(IXLStyle style, int columns = 1, int rows = 1) => GetRange(rows, columns).Style = style;

    public void SetBoldStyle(int columns = 1, int rows = 1)
    {
        var style = GetStyle();
        style.Font.Bold = true;
        GetRange(rows, columns).Style = style;
    }

    public void SetFontName(string fontName, int columns = 1, int rows = 1)
    {
        var style = GetStyle();
        style.Font.FontName = fontName;
        GetRange(rows, columns).Style = style;
    }

    public void SetBoxBorders(XLBorderStyleValues borderStyle, int columns = 1, int rows = 1)
    {
        var style = GetStyle();
        style.Border.BottomBorder = borderStyle;
        style.Border.TopBorder = borderStyle;
        style.Border.RightBorder = borderStyle;
        style.Border.LeftBorder = borderStyle;
        GetRange(rows, columns).Style = style;
    }

    public void SetVerticalBorders(XLBorderStyleValues borderStyle, int columns = 1, int rows = 1)
    {
        var style = GetStyle();
        style.Border.BottomBorder = borderStyle;
        style.Border.TopBorder = borderStyle;
        GetRange(rows, columns).Style = style;
    }

    public void SetHorizontalBorders(XLBorderStyleValues borderStyle, int columns = 1, int rows = 1)
    {
        var style = GetStyle();
        style.Border.RightBorder = borderStyle;
        style.Border.LeftBorder = borderStyle;
        GetRange(rows, columns).Style = style;
    }

    public void SetTopBorder(XLBorderStyleValues borderStyle, int columns = 1, int rows = 1)
    {
        var style = GetStyle();
        style.Border.TopBorder = borderStyle;
        GetRange(rows, columns).Style = style;
    }

    public void SetBottomBorder(XLBorderStyleValues borderStyle, int columns = 1, int rows = 1)
    {
        var style = GetStyle();
        style.Border.BottomBorder = borderStyle;
        GetRange(rows, columns).Style = style;
    }

    public void SetLeftBorder(XLBorderStyleValues borderStyle, int columns = 1, int rows = 1)
    {
        var style = GetStyle();
        style.Border.LeftBorder = borderStyle;
        GetRange(rows, columns).Style = style;
    }

    public void SetRightBorder(XLBorderStyleValues borderStyle, int columns = 1, int rows = 1)
    {
        var style = GetStyle();
        style.Border.RightBorder = borderStyle;
        GetRange(rows, columns).Style = style;
    }

    public void SetBackgroundColor(XLColor color, int columns = 1, int rows = 1)
    {
        var style = GetStyle();
        style.Fill.BackgroundColor = color;
        GetRange(rows, columns).Style = style;
    }

    public void SetFontColor(XLColor color, int columns = 1, int rows = 1)
    {
        var style = GetStyle();
        style.Font.FontColor = color;
        GetRange(rows, columns).Style = style;
    }

    public void SetFontSize(double size, int columns = 1, int rows = 1)
    {
        var style = GetStyle();
        style.Font.FontSize = size;
        GetRange(rows, columns).Style = style;
    }

    public void SetFontUnderline(XLFontUnderlineValues underline, int columns = 1, int rows = 1)
    {
        var style = GetStyle();
        style.Font.Underline = underline;
        GetRange(rows, columns).Style = style;
    }

    public void SetFontItalic(bool italic, int columns = 1, int rows = 1)
    {
        var style = GetStyle();
        style.Font.Italic = italic;
        GetRange(rows, columns).Style = style;
    }

    public void SetVerticalAlignment(XLAlignmentVerticalValues alignment, int columns = 1, int rows = 1)
    {
        var style = GetStyle();
        style.Alignment.Vertical = alignment;
        GetRange(rows, columns).Style = style;
    }

    public void SetHorizontalAlignment(XLAlignmentHorizontalValues alignment, int columns = 1, int rows = 1)
    {
        var style = GetStyle();
        style.Alignment.Horizontal = alignment;
        GetRange(rows, columns).Style = style;
    }

    public void SetBordersColor(XLColor color, int columns = 1, int rows = 1)
    {
        var style = GetStyle();
        style.Border.BottomBorderColor = color;
        style.Border.TopBorderColor = color;
        style.Border.RightBorderColor = color;
        style.Border.LeftBorderColor = color;
        GetRange(rows, columns).Style = style;
    }
    
    public void AdjustColumnWidth(int[]? columns = default)
    {
        if (columns == null) _worksheet.ColumnsUsed().AdjustToContents();
        else foreach (var column in columns) _worksheet.Column(column).AdjustToContents();
    }

    public string Name => _worksheet.Name;
}