using DocumentFormat.OpenXml.Drawing.Charts;

namespace PowerPointEditorLib;

public static class BarChartExtensions
{
    public static BarChart FindBarChart(this SlidePart slidePart, int chartNumber)
    {
        IEnumerable<ChartPart> chartPart = slidePart.ChartParts.Where(cp => cp.ChartSpace.Descendants<BarChart>().Any());

        if (chartPart.Any() == false)
            throw new NullReferenceException("Cannot find any bar charts on the slide");

        
        
        ChartPart? selectedChartPart = chartPart.ElementAtOrDefault(chartNumber - 1);

        if (selectedChartPart is null)
            throw new NullReferenceException($"Cannot find a bar chart at the given position ('{chartNumber}')");



        return selectedChartPart.ChartSpace.Descendants<BarChart>().First();
    }

    public static BarChartSeries FindSeriesByName(this BarChart barChart, string name)
    {
        BarChartSeries? series = barChart.Elements<BarChartSeries>().FirstOrDefault(s => s.Descendants<StringCache>().First().InnerText == name);

        if (series is null)
            throw new NullReferenceException($"No bar chart series has been found with that name ('{name}')");

        return series;
    }

    public static void ReplaceValuesWith(this BarChartSeries series, List<string> newValues)
    {
        var values = series.GetFirstChild<Values>() ?? throw new NullReferenceException("Cannot read chart elements (values)");
        var numRef = values.GetFirstChild<NumberReference>() ?? throw new NullReferenceException("Cannot read chart elements (numberReference)");
        var numCache = numRef.GetFirstChild<NumberingCache>() ?? throw new NullReferenceException("Cannot read chart elements (numberingCache)");

        int chartValuesCount = numCache.Elements<NumericPoint>().Count();
        if (chartValuesCount != newValues.Count())
            throw new ArgumentException($"The number of values provided for the chart edit ({newValues.Count()}) does not match the number of values used in the chart ({chartValuesCount})");

        int valuePosition = -1;
        foreach (var np in numCache.Elements<NumericPoint>())
        {
            valuePosition++;
            var nv = np.GetFirstChild<NumericValue>() ?? throw new NullReferenceException("Cannot read chart elements (numericValue)");
            nv.Text = newValues[valuePosition];
        }
    }
}
