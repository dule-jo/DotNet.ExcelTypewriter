using ClosedXML.Excel;
using DotNet.XlsTypewriter.Common;

namespace DotNet.XlsTypewriter;

public class Worksheet
{
    private IXLWorksheet _worksheet;

    private int _rowIndex = 1;

    private int _columnIndex = 1;

    private string RowName() => XlsHelper.IntToColumnName(_rowIndex);

    private static string RowName(int index) => XlsHelper.IntToColumnName(index);

    private string CurrentCellName() => $"{RowName()}{_columnIndex}";

    private static string CellName(int rowIndex, int columnIndex) => $"{RowName(rowIndex)}{columnIndex}";

    public Worksheet(IXLWorksheet worksheet)
    {
        _worksheet = worksheet;
    }

    public void Print(object value)
    {
        _worksheet.Cell(CurrentCellName()).Value = value.ToString();
        _columnIndex++;
    }

    public void NewRow()
    {
        _rowIndex++;
        _columnIndex = 1;
    }

    public void SkipColumn(int count = 1) => _columnIndex += count;

    public IXLRange GetRange(int rows, int columns) => _worksheet.Range(
        CurrentCellName(),
        CellName(_rowIndex + rows, _columnIndex + columns)
    );

    public void Merge(int rows, int columns) => GetRange(rows, columns).Merge();

    public IXLStyle GetStyle() => _worksheet.Cell(CurrentCellName()).Style;

    public void SetStyle(IXLStyle style, int rows = 1, int columns = 1) => GetRange(rows, columns).Style = style;

    public string Name => _worksheet.Name;
}