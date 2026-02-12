namespace ExcelTests.Calculators.Models;

public record CellRange(int Row1, int Col1, int Row2, int Col2)
{
    public CellRange(int row, int col)
        : this(row, col, row, col) { }

    public int TopRow => Math.Min(Row1, Row2);
    public int BottomRow => Math.Max(Row1, Row2);
    public int LeftCol => Math.Min(Col1, Col2);
    public int RightCol => Math.Max(Col1, Col2);

    public bool IsValid => Row1 > -1 && Col1 > -1 && Row2 > -1 && Col2 > -1;
    public bool IsSingleCell => Row1 == Row2 && Col1 == Col2;
}
