using ExcelTests.Calculators;
using ExcelTests.Models;

List<ExcelSheet> sheets =
[
    new(
        "SomeTitle",
        [
            [new(1), new(2), new("=A1+A2")],
        ]
    ),
];

var calculator = new ExcelConfigCalculator(new ExcelConfig(sheets));
var result = calculator.Evaluate("SomeTitle", "A3");

Console.WriteLine(result);
