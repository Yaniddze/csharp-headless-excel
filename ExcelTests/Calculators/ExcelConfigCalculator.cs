using ExcelTests.Calculators.Models;
using ExcelTests.Models;

namespace ExcelTests.Calculators;

public class ExcelConfigCalculator : CalcEngine.CalcEngine
{
    private readonly ExcelConfig config;

    public ExcelConfigCalculator(ExcelConfig config)
    {
        this.config = config;
        IdentifierChars = "$:!";
    }

    public object Evaluate(int rowIndex, int colIndex)
    {
        var sheet = config.Sheets.First();
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
        // Find the sheet by title
        var targetSheet = config.Sheets.FirstOrDefault(s => s.Title == sheet);
        if (targetSheet == null)
        {
            throw new ArgumentException($"Sheet '{sheet}' not found");
        }

        // Set the current sheet context for cross-sheet references
        // We need to modify the expression to handle sheet references
        var modifiedExpression = expression;

        // Replace sheet references like "SomeTitle!A1" with just "A1"
        // and evaluate in the context of the referenced sheet
        foreach (var s in config.Sheets)
        {
            var sheetRef = $"{s.Title}!";
            if (expression.Contains(sheetRef))
            {
                // For now, we'll handle simple cases where we need to evaluate
                // the referenced cell from another sheet
                var cellRef = expression.Substring(expression.IndexOf(sheetRef) + sheetRef.Length);
                return EvaluateCellFromSheet(s, cellRef);
            }
        }

        return Evaluate(modifiedExpression);
    }

    private object EvaluateCellFromSheet(ExcelSheet sheet, string cellRef)
    {
        // Parse the cell reference (e.g., "A1", "B2")
        int colIndex = cellRef[0] - 'A';
        int rowIndex = int.Parse(cellRef.Substring(1)) - 1;

        if (
            rowIndex >= 0
            && rowIndex < sheet.Cells.Count
            && colIndex >= 0
            && colIndex < sheet.Cells[rowIndex].Count
        )
        {
            var cellValue = sheet.Cells[rowIndex][colIndex].Value;
            if (cellValue is string str && str.StartsWith("="))
            {
                // This is a formula, we need to evaluate it in the context of this sheet
                // For now, we'll use a simple approach
                return EvaluateInSheetContext(sheet, str);
            }
            return cellValue;
        }
        return null;
    }

    private object EvaluateInSheetContext(ExcelSheet sheet, string formula)
    {
        // Create a temporary calculator for this sheet
        var tempConfig = new ExcelConfig(new[] { sheet });
        var tempCalculator = new ExcelConfigCalculator(tempConfig);
        return tempCalculator.Evaluate(formula);
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
        int index = 0;

        // parse row
        int row = -1;
        var absCol = false;
        for (; index < cell.Length; index++)
        {
            var c = cell[index];
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
        for (; index < cell.Length; index++)
        {
            var c = cell[index];
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
        if (index < cell.Length)
        {
            throw new Exception("Invalid cell reference.");
        }

        // done
        return new CellRange(row - 1, col - 1);
    }

    private CellRange MergeRanges(CellRange rng1, CellRange rng2)
    {
        return new CellRange(
            Math.Min(rng1.TopRow, rng2.TopRow),
            Math.Min(rng1.LeftCol, rng2.LeftCol),
            Math.Max(rng1.BottomRow, rng2.BottomRow),
            Math.Max(rng1.RightCol, rng2.RightCol)
        );
    }
}
