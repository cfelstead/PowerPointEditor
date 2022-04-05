using DocumentFormat.OpenXml.Drawing.Charts;

namespace PowerPointEditorLib;

public static class PieChartExtensions
{
    public static PieChart FindPieChart(this SlidePart slidePart, int chartNumber)
    {
        IEnumerable<ChartPart> chartPart = slidePart.ChartParts.Where(cp => cp.ChartSpace.Descendants<PieChart>().Any());

        if (chartPart.Any() == false)
            throw new NullReferenceException("Cannot find any pie charts on the slide");

        
        
        ChartPart? selectedChartPart = chartPart.ElementAtOrDefault(chartNumber - 1);

        if (selectedChartPart is null)
            throw new NullReferenceException($"Cannot find a pie chart at the given position ('{chartNumber}')");



        return selectedChartPart.ChartSpace.Descendants<PieChart>().First();
    }

    public static PieChartSeries FindSeries(this PieChart pieChart)
    {
        PieChartSeries? series = pieChart.Elements<PieChartSeries>().FirstOrDefault();

        if (series is null)
            throw new NullReferenceException($"No pie chart series has been found");

        return series;
    }

    public static void ReplaceValuesWith(this PieChartSeries series, List<string> newValues)
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
