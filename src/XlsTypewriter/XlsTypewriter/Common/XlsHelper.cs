namespace XlsTypewriter.Common;

public static class XlsHelper
{
    public static string IntToColumnName(int columnNumber)
    {
        var columnName = string.Empty;
        int modulo;

        while (columnNumber > 0)
        {
            modulo = (columnNumber - 1) % 26;
            columnName = Convert.ToChar(65 + modulo) + columnName;
            columnNumber = (columnNumber - modulo) / 26;
        }

        return columnName;
    }
}