using System.Collections;

namespace ExcelTests.Calculators.Models;

class CellRangeReference(CellRange rng, Func<string?, int, int, object> EvaluateCell)
    : CalcEngine.IValueObject,
        IEnumerable
{
    bool evaluating;

    // ** IValueObject
    public object GetValue()
    {
        return GetValue(rng);
    }

    // ** IEnumerable
    public IEnumerator GetEnumerator()
    {
        for (int r = rng.TopRow; r <= rng.BottomRow; r++)
        {
            for (int c = rng.LeftCol; c <= rng.RightCol; c++)
            {
                var localRng = new CellRange(rng.Sheet1, r, c);
                yield return GetValue(localRng);
            }
        }
    }

    private object GetValue(CellRange rng)
    {
        if (evaluating)
        {
            throw new Exception("Circular Reference");
        }
        try
        {
            evaluating = true;
            return EvaluateCell(rng.Sheet1, rng.Row1, rng.Col1);
        }
        finally
        {
            evaluating = false;
        }
    }
}
