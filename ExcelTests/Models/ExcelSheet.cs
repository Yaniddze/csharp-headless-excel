namespace ExcelTests.Models;

public record ExcelSheet(string Title, List<List<ExcelCell>> Cells);
