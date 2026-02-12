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
                    [new(1), new(2), new(3), new("=SUM(A1,A2,A3)")],
                ]
            ),
        ];

        var calculator = new ExcelConfigCalculator(new ExcelConfig(sheets));

        // Act
        var result = (double)calculator.Evaluate("SomeTitle!A4");

        // Assert
        Assert.Equal(6d, result);
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
                    [new(1.23), new(2.25), new("=ROUND(SUM(A1,A2),1)")],
                ]
            ),
        ];

        var calculator = new ExcelConfigCalculator(new ExcelConfig(sheets));

        // Act
        var result = (double)calculator.Evaluate("SomeTitle!A3");

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
                    [new(1.23), new(2.25), new("=ROUNDDOWN(SUM(A1,A2))")],
                ]
            ),
        ];

        var calculator = new ExcelConfigCalculator(new ExcelConfig(sheets));

        // Act
        var result = (double)calculator.Evaluate("SomeTitle!A3");

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
                    [new(1), new(2), new(3), new("=IF(A1=1,A2,A3)")],
                ]
            ),
        ];

        var calculator = new ExcelConfigCalculator(new ExcelConfig(sheets));

        // Act
        var result = (int)calculator.Evaluate("SomeTitle!A4");

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
                    [new(1), new(2), new(3), new("=MIN(A1,A2,A3)")],
                ]
            ),
        ];

        var calculator = new ExcelConfigCalculator(new ExcelConfig(sheets));

        // Act
        var result = (double)calculator.Evaluate("SomeTitle!A4");

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
