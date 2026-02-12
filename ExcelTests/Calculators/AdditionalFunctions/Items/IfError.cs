using CalcEngine;

namespace ExcelTests.Calculators.AdditionalFunctions.Items;

public static class IfError
{
    public static ExcelConfigCalculator RegisterIfError(this ExcelConfigCalculator calculator)
    {
        calculator.RegisterFunction("IFERROR", 2, Ff);
        return calculator;
    }

    public static object Ff(List<Expression> p)
    {
        if (p[0] == double.PositiveInfinity)
        {
            return p[1].Evaluate();
        }
        return p[0].Evaluate();
    }
}
