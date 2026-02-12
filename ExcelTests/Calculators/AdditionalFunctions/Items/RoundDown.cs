using CalcEngine;

namespace ExcelTests.Calculators.AdditionalFunctions.Items;

public static class RoundDown
{
    public static ExcelConfigCalculator RegisterRoundDown(this ExcelConfigCalculator calculator)
    {
        calculator.RegisterFunction("ROUNDDOWN", 1, Ff);
        return calculator;
    }

    public static object Ff(List<Expression> p)
    {
        var decimals = p.Count == 2 ? int.Parse(p[1].Evaluate().ToString()) : 0;

        var power = Math.Pow(10, decimals);
        var p0 = p[0].Evaluate();

        var rounded = Math.Floor((double)p0 * power) / power;
        return rounded;
    }
}
