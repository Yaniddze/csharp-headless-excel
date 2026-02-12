using ExcelTests.Calculators.Models;
using ExcelTests.Models;

namespace ExcelTests.Calculators;

public class ExcelConfigCalculator : CalcEngine.CalcEngine
{
    private readonly ExcelConfig config;
    private string? parentSheet;

    public ExcelConfigCalculator(ExcelConfig config)
    {
        this.config = config;
        IdentifierChars = "$:!";
    }

    public object Evaluate(string? sheetToFind, int rowIndex, int colIndex)
    {
        var sheet = config.Sheets.First(x => x.Title == sheetToFind);
        // get the value
        var val = sheet.Cells[rowIndex][colIndex].Value;
        var text = val as string;
        if (!string.IsNullOrEmpty(text) && text.StartsWith("="))
        {
            return Evaluate(text);
        }
        return val;
    }

    public object Evaluate(string sheet, string expression)
    {
        parentSheet = sheet;

        return Evaluate(expression);
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
        var address = cell;
        var currentSheet = parentSheet;

        if (cell.Contains('!'))
        {
            var parts = cell.Split('!');
            currentSheet = parts[0].Replace("'", "").Replace("\"", "");
            address = parts[1];
        }

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

    private CellRange MergeRanges(CellRange rng1, CellRange rng2)
    {
        return new CellRange(
            rng1.Sheet1,
            Math.Min(rng1.TopRow, rng2.TopRow),
            Math.Min(rng1.LeftCol, rng2.LeftCol),
            rng1.Sheet2,
            Math.Max(rng1.BottomRow, rng2.BottomRow),
            Math.Max(rng1.RightCol, rng2.RightCol)
        );
    }
}
