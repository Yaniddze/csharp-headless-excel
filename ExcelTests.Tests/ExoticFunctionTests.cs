using ExcelTests.Calculators;
using ExcelTests.Models;

namespace ExcelTests.Tests;

public class ExoticFunctionTests
{
    [Fact]
    public void SUM_Works()
    {
        // Arrange
        List<ExcelSheet> sheets =
        [
            new(
                "SomeTitle",
                [
                    [new(1), new(2), new(3), new("=SUM(A1,B1,C1)")],
                ]
            ),
        ];

        var calculator = new ExcelConfigCalculator(new ExcelConfig(sheets));

        // Act
        var result = (double)calculator.Evaluate("SomeTitle!D1");

        // Assert
        Assert.Equal(6d, result);
    }

    [Fact]
    public void SUMPRODUCT_Works()
    {
        // Arrange
        List<ExcelSheet> sheets =
        [
            new(
                "SomeTitle",
                [
                    [new(1), new(8), new(15)],
                    [new(2), new(9), new(16)],
                    [new(3), new(10), new(17)],
                    [new(4), new(11), new(18)],
                    [new(5), new(12), new(19)],
                    [new(6), new(13), new(20)],
                    [new(7), new(14), new(21)],
                    [new("=SUMPRODUCT(A1:A7,B1:B7)")],
                ]
            ),
        ];

        var calculator = new ExcelConfigCalculator(new ExcelConfig(sheets));

        // Act
        var result = (double)calculator.Evaluate("SomeTitle!A8");

        // Assert
        Assert.Equal(336d, result);
    }

    [Fact]
    public void ROUND_Works()
    {
        // Arrange
        List<ExcelSheet> sheets =
        [
            new(
                "SomeTitle",
                [
                    [new(1.23), new(2.25), new("=ROUND(SUM(A1,B1),1)")],
                ]
            ),
        ];

        var calculator = new ExcelConfigCalculator(new ExcelConfig(sheets));

        // Act
        var result = (double)calculator.Evaluate("SomeTitle!C1");

        // Assert
        Assert.Equal(3.5d, result);
    }

    [Fact]
    public void ROUNDDOWN_Works()
    {
        // Arrange
        List<ExcelSheet> sheets =
        [
            new(
                "SomeTitle",
                [
                    [new(1.23), new(2.25), new("=ROUNDDOWN(SUM(A1,B1))")],
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
    public void IF_Works()
    {
        // Arrange
        List<ExcelSheet> sheets =
        [
            new(
                "SomeTitle",
                [
                    [new(1), new(2), new(3), new("=IF(A1=1,B1,C1)")],
                ]
            ),
        ];

        var calculator = new ExcelConfigCalculator(new ExcelConfig(sheets));

        // Act
        var result = (int)calculator.Evaluate("SomeTitle!D1");

        // Assert
        Assert.Equal(2d, result);
    }

    [Fact]
    public void MIN_Works()
    {
        // Arrange
        List<ExcelSheet> sheets =
        [
            new(
                "SomeTitle",
                [
                    [new(1), new(2), new(3), new("=MIN(A1,B1,C1)")],
                ]
            ),
        ];

        var calculator = new ExcelConfigCalculator(new ExcelConfig(sheets));

        // Act
        var result = (double)calculator.Evaluate("SomeTitle!D1");

        // Assert
        Assert.Equal(1d, result);
    }

    [Fact]
    public void EXP_Works()
    {
        // Arrange
        List<ExcelSheet> sheets =
        [
            new(
                "SomeTitle",
                [
                    [new("=EXP(1)")],
                ]
            ),
        ];

        var calculator = new ExcelConfigCalculator(new ExcelConfig(sheets));

        // Act
        var result = (double)calculator.Evaluate("SomeTitle!A1");

        // Assert
        Assert.Equal(2.718281828459045, result);
    }

    [Fact]
    public void IFERROR_Works()
    {
        // Arrange
        List<ExcelSheet> sheets =
        [
            new(
                "SomeTitle",
                [
                    [new("=IFERROR(1/0,123)")],
                ]
            ),
        ];

        var calculator = new ExcelConfigCalculator(new ExcelConfig(sheets));

        // Act
        var result = calculator.Evaluate("SomeTitle!A1");

        // Assert
        Assert.Equal(123d, result);
    }
}
