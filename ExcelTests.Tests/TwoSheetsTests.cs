using ExcelTests.Calculators;
using ExcelTests.Models;

namespace ExcelTests.Tests;

public class TwoSheetsTests
{
    [Fact]
    public void SimplestTest_Works()
    {
        // Arrange
        List<ExcelSheet> sheets =
        [
            new(
                "SomeTitle",
                [
                    [new(1), new(2), new("=A1+A2")],
                ]
            ),
            new(
                "SomeTitle2",
                [
                    [new(3), new(4), new("=SUM(A1:A2)+SomeTitle!A3")],
                ]
            ),
        ];

        var calculator = new ExcelConfigCalculator(new ExcelConfig(sheets));

        // Act
        var result = (double)calculator.Evaluate("SomeTitle2", "A3");

        // Assert
        Assert.Equal(10d, result);
    }
}
