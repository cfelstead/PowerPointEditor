using DocumentFormat.OpenXml.Drawing.Charts;

namespace PowerPointEditorLib.Tests;

[Collection("UsesFile-Test")]
public class BarChartTests
{
    [Fact]
    public void FindBarChart_NoCharts_ThrowsError()
    {
        string examplePpt = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt, true);

        Assert.Throws<NullReferenceException>(() => ppt.ForSlide(1).FindBarChart(1));

        ppt.Close();
    }

    [Fact]
    public void FindBarChart_ChartNumberIncorrect_ThrowsError()
    {
        string examplePpt = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt, true);

        Assert.Throws<NullReferenceException>(() => ppt.ForSlide(3).FindBarChart(2));

        ppt.Close();
    }

    [Fact]
    public void FindBarChart_ChartReturned()
    {
        string examplePpt = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt, true);

        Assert.NotNull(ppt.ForSlide(3).FindBarChart(1));

        ppt.Close();
    }

    [Fact]
    public void FindSeriesByName_ErrorsIfNameNotFound()
    {
        string examplePpt = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt, true);

        Assert.Throws<NullReferenceException>(() => ppt.ForSlide(3).FindBarChart(1).FindSeriesByName("Series Not Found"));

        ppt.Close();
    }

    [Fact]
    public void FindSeriesByName_SeriesReturned()
    {
        string examplePpt = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt, true);

        Assert.NotNull(ppt.ForSlide(3).FindBarChart(1).FindSeriesByName("Series 1"));

        ppt.Close();
    }

    [Fact]
    public void ReplaceValuesWith_MakesChange()
    {
        string examplePpt = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt, true);

        BarChartSeries series = ppt.ForSlide(3).FindBarChart(1).FindSeriesByName("Series 1");
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

        Assert.Throws<ArgumentException>(() => ppt.ForSlide(3).FindBarChart(1).FindSeriesByName("Series 1").ReplaceValuesWith(new List<string> { "2", "4", "4" }));

        ppt.Close();
    }
}
