using ExcelTests.Calculators;
using ExcelTests.Models;

namespace ExcelTests.Tests;

public class OneSheetTests
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
                    [new(1), new(2), new("=A1+B1")],
                ]
            ),
        ];

        var calculator = new ExcelConfigCalculator(new ExcelConfig(sheets));

        // Act
        var result = (double)calculator.Evaluate("SomeTitle!C1");

        // Assert
        Assert.Equal(3d, result);
    }

    [Fact]
    public void ManyRows_Works()
    {
        // Arrange
        List<ExcelSheet> sheets =
        [
            new(
                "SomeTitle",
                [
                    [new(1), new(2)],
                    [new(3), new(4), new("=A1+B2")],
                ]
            ),
        ];

        var calculator = new ExcelConfigCalculator(new ExcelConfig(sheets));

        // Act
        var result = (double)calculator.Evaluate("SomeTitle!C2");

        // Assert
        Assert.Equal(5d, result);
    }

    [Fact]
    public void Sum_OneRow_Works()
    {
        // Arrange
        List<ExcelSheet> sheets =
        [
            new(
                "SomeTitle",
                [
                    [new(1), new(2), new("=SUM(A1:B1)")],
                ]
            ),
        ];

        var calculator = new ExcelConfigCalculator(new ExcelConfig(sheets));

        // Act
        var result = (double)calculator.Evaluate("SomeTitle!C1");

        // Assert
        Assert.Equal(3d, result);
    }

    [Fact]
    public void Sum_ManyRows_Works()
    {
        // Arrange
        List<ExcelSheet> sheets =
        [
            new(
                "SomeTitle",
                [
                    [new(1), new(2)],
                    [new(3), new(4), new("=SUM(A1:B2)")],
                ]
            ),
        ];

        var calculator = new ExcelConfigCalculator(new ExcelConfig(sheets));

        // Act
        var result = (double)calculator.Evaluate("SomeTitle!C2");

        // Assert
        Assert.Equal(10d, result);
    }
}
