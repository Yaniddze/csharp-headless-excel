using System.Text.RegularExpressions;
using ExcelTests.Calculators.AdditionalFunctions;
using ExcelTests.Calculators.Models;
using ExcelTests.Models;

namespace ExcelTests.Calculators;

public partial class ExcelConfigCalculator : CalcEngine.CalcEngine
{
    // SheetName -> RowIndex -> ColIndex -> Value
    private Dictionary<string, Dictionary<int, Dictionary<int, object>>> cache = [];

    public ExcelConfigCalculator(ExcelConfig config)
    {
        IdentifierChars = "$:!";
        ProcessConfig(config);

        this.RegisterFunctions();
    }

    public void ProcessConfig(ExcelConfig config)
    {
        cache = config.Sheets.ToDictionary(
            x => x.Title,
            x =>
                x.Cells.Select((row, rowIndex) => new { row, rowIndex })
                    .ToDictionary(
                        y => y.rowIndex,
                        y =>
                            y.row.Select((cell, colIndex) => new { cell, colIndex })
                                .ToDictionary(
                                    z => z.colIndex,
                                    z => ProcessCellValue(z.cell.Value, x.Title)
                                )
                    )
        );
    }

    private object ProcessCellValue(object value, string sheetName)
    {
        if (value is string strVal && strVal.StartsWith("="))
        {
            var regex = GetRegex();
            var result = regex.Replace(strVal, match => $"{sheetName}!{match.Value}");
            return result;
        }

        return value;
    }

    public object Evaluate(string sheetToFind, int rowIndex, int colIndex)
    {
        var val = cache[sheetToFind][rowIndex][colIndex];

        var text = val as string;
        if (!string.IsNullOrEmpty(text) && text.StartsWith("="))
        {
            return Evaluate(text);
        }
        return val;
    }

    /// <summary>
    /// Парсинг ренжей в конкретные значения
    /// </summary>
    /// <param name="identifier">String representing a cell range (e.g. "A1" or "A1:B12").</param>
    /// <returns>A <see cref="CellRange"/> object that represents the given range.</returns>
    public override object GetExternalObject(string identifier)
    {
        var cells = identifier.Split(':');
        if (cells.Length > 0 && cells.Length < 3)
        {
            var rng = GetRange(cells[0]);
            if (cells.Length > 1)
            {
                rng = MergeRanges(rng, GetRange(cells[1]));
            }
            if (rng.IsValid)
            {
                return new CellRangeReference(rng, Evaluate);
            }
        }

        // this doesn't look like a range
        return null;
    }

    private CellRange GetRange(string cell)
    {
        if (!cell.Contains('!'))
        {
            return new CellRange(null, -1, -1);
        }

        var parts = cell.Split('!');
        var currentSheet = parts[0].Replace("'", "").Replace("\"", "");
        var address = parts[1];

        int index = 0;

        // parse row
        int row = -1;
        var absCol = false;
        for (; index < address.Length; index++)
        {
            var c = address[index];
            if (c == '$' && !absCol)
            {
                absCol = true;
                continue;
            }
            if (!char.IsLetter(c))
            {
                break;
            }
            if (row < 0)
                row = 0;
            row = 26 * row + char.ToUpper(c) - 'A' + 1;
        }

        // parse column
        int col = -1;
        var absRow = false;
        for (; index < address.Length; index++)
        {
            var c = address[index];
            if (c == '$' && !absRow)
            {
                absRow = true;
                continue;
            }
            if (!char.IsDigit(c))
            {
                break;
            }
            if (col < 0)
                col = 0;
            col = 10 * col + (c - '0');
        }

        // sanity
        if (index < address.Length)
        {
            throw new Exception("Invalid cell reference.");
        }

        // done
        return new CellRange(currentSheet!, row - 1, col - 1);
    }

    private static CellRange MergeRanges(CellRange rng1, CellRange rng2) =>
        new(
            rng1.Sheet1,
            Math.Min(rng1.TopRow, rng2.TopRow),
            Math.Min(rng1.LeftCol, rng2.LeftCol),
            rng1.Sheet2,
            Math.Max(rng1.BottomRow, rng2.BottomRow),
            Math.Max(rng1.RightCol, rng2.RightCol)
        );

    [GeneratedRegex(@"(?<![A-Za-z0-9_!$])[A-Z]+\d+")]
    private static partial Regex GetRegex();
}
