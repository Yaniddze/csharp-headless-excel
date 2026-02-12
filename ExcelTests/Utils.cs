namespace ExcelTests;

public static class DataGridCalc
{
    public static string GetAlphaColumnHeader(int i)
    {
        return GetAlphaColumnHeader(i, string.Empty);
    }

    private static string GetAlphaColumnHeader(int i, string s)
    {
        var rem = i % 26;
        s = (char)('A' + rem) + s;
        i = i / 26 - 1;
        return i < 0 ? s : GetAlphaColumnHeader(i, s);
    }
}
