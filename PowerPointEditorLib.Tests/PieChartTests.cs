using DocumentFormat.OpenXml.Drawing.Charts;

namespace PowerPointEditorLib.Tests;

[Collection("UsesFile-Test")]
public class PieChartTests
{
    [Fact]
    public void FindPieChart_NoCharts_ThrowsError()
    {
        string examplePpt = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt, true);

        Assert.Throws<NullReferenceException>(() => ppt.ForSlide(1).FindPieChart(1));

        ppt.Close();
    }

    [Fact]
    public void FindPieChart_ChartNumberIncorrect_ThrowsError()
    {
        string examplePpt = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt, true);

        Assert.Throws<NullReferenceException>(() => ppt.ForSlide(4).FindPieChart(2));

        ppt.Close();
    }

    [Fact]
    public void FindPieChart_ChartReturned()
    {
        string examplePpt = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt, true);

        Assert.NotNull(ppt.ForSlide(4).FindPieChart(1));

        ppt.Close();
    }

    [Fact]
    public void FindSeries_SeriesReturned()
    {
        string examplePpt = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt, true);

        Assert.NotNull(ppt.ForSlide(4).FindPieChart(1).FindSeries());

        ppt.Close();
    }

    [Fact]
    public void ReplaceValuesWith_MakesChange()
    {
        string examplePpt = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt, true);

        PieChartSeries series = ppt.ForSlide(4).FindPieChart(1).FindSeries();
        string beforeXml = series.InnerXml;
        series.ReplaceValuesWith(new List<string> { "2", "4", "4", "2" });
        string afterXml = series.InnerXml;

        Assert.NotEqual(beforeXml, afterXml);

        ppt.Close();
    }

    [Fact]
    public void ReplaceValuesWith_WrongNumberOfValues_ThrowsErrors()
    {
        string examplePpt = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt, true);

        Assert.Throws<ArgumentException>(() => ppt.ForSlide(4).FindPieChart(1).FindSeries().ReplaceValuesWith(new List<string> { "2", "4", "4" }));

        ppt.Close();
    }
}
