namespace PowerPointEditorLib.Tests;

[Collection("UsesFile-Test")]
public class SlideSelectionTests
{
    [Fact]
    public void ForAllSlides_ReturnsCorrectCount()
    {
        string examplePpt = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt, true);

        Assert.Equal(4, ppt.ForAllSlides().Count);

        ppt.Close();
    }

    [Fact]
    public void ForSlide_ReturnsCorrectSlide()
    {
        string examplePpt = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt, true);

        Assert.Contains("Slide 1", ppt.ForSlide(1).Slide.InnerXml);

        ppt.Close();
    }
}
