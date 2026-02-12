using ExcelTests.Calculators;
using ExcelTests.Models;

namespace ExcelTests.Tests;

public class MathTests
{
    [Fact]
    public void Exponantion_Works()
    {
        // Arrange
        List<ExcelSheet> sheets =
        [
            new(
                "SomeTitle",
                [
                    [new("=3^3")],
                ]
            ),
        ];

        var calculator = new ExcelConfigCalculator(new ExcelConfig(sheets));

        // Act
        var result = (double)calculator.Evaluate("SomeTitle!A1");

        // Assert
        Assert.Equal(27d, result);
    }
}
