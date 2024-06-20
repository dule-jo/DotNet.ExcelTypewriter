using ClosedXML.Excel;

namespace XlsTypewriter;

public class Workbook : IDisposable
{
    private readonly List<Worksheet> _worksheets = new();

    private readonly XLWorkbook _workbook;

    public Workbook()
    {
        _workbook = new XLWorkbook();
    }

    public void SaveAs(string path) => _workbook.SaveAs(path);

    public Worksheet AddWorksheet(string name)
    {
        var worksheet = new Worksheet(_workbook.AddWorksheet(name));

        _worksheets.Add(worksheet);

        return worksheet;
    }

    public Worksheet GetWorksheet(string name)
    {
        return _worksheets.FirstOrDefault(w => w.Name == name);
    }

    public void Dispose()
    {
        _workbook.Dispose();
    }
}