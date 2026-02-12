using ExcelTests.Calculators.AdditionalFunctions.Items;

namespace ExcelTests.Calculators.AdditionalFunctions;

public static class SetupFunctions
{
    public static void RegisterFunctions(this ExcelConfigCalculator calculator)
    {
        calculator.RegisterIfError().RegisterRoundDown();
    }
}
