using ClosedXML.Excel;

namespace DotNet.XlsTypewriter;

public class Workbook : IDisposable
{
    private List<Worksheet> _worksheets;

    private XLWorkbook _workbook;

    public Workbook(string path)
    {
        _workbook = new XLWorkbook(path);
        _worksheets = new List<Worksheet>();
    }

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