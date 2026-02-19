using System.Collections;
using CalcEngine;
using CalcEngine.Functions;

namespace ExcelTests.Calculators.AdditionalFunctions.Items;

public static class Sumproduct
{
    public static ExcelConfigCalculator RegisterSumproduct(this ExcelConfigCalculator calculator)
    {
        calculator.RegisterFunction("SUMPRODUCT", 2, int.MaxValue, Ff);
        return calculator;
    }

    public static object Ff(List<Expression> p)
    {
        List<List<double>> nums = p.Select(
                (x, xIndex) =>
                {
                    List<double> localList = [];
                    var enumerableX = x as IEnumerable;

                    foreach (var item in enumerableX)
                    {
                        var parsed = double.Parse(item.ToString()!.Replace(",", "."));
                        localList.Add(parsed);
                    }

                    return localList;
                }
            )
            .ToList();

        var result = SumProduct(nums);

        return result;
    }

    public static double SumProduct(List<List<double>> columns)
    {
        if (columns == null || columns.Count == 0)
            return 0.0;

        // Проверка: все списки должны быть одной длины
        var rowCount = columns[0].Count;
        if (columns.Any(col => col.Count != rowCount))
            throw new ArgumentException("Все списки должны иметь одинаковую длину.");

        var totalSum = 0.0;

        // Проходим по индексам (строкам)
        for (var i = 0; i < rowCount; i++)
        {
            var rowProduct = 1.0;

            // Перемножаем значения i-го элемента из каждого списка
            for (var j = 0; j < columns.Count; j++)
            {
                rowProduct *= columns[j][i];
            }

            totalSum += rowProduct;
        }

        return totalSum;
    }
}
