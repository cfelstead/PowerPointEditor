namespace PowerPointEditorLib.Tests;

[Collection("UsesFile-Test")]
public class TextTests
{
    [Fact]
    public void SingleSlideEdit_TextNotFound_NoChangeMade()
    {
        string examplePpt = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt, true);

        SlidePart slidePart = ppt.ForSlide(1);
        string beforeXml = slidePart.Slide.InnerXml;
        slidePart.ReplaceText("NOT FOUND").With("SOMETHING ELSE");
        string afterXml = slidePart.Slide.InnerXml;

        Assert.Equal(beforeXml, afterXml);

        ppt.Close();
    }

    [Fact]
    public void SingleSlideEdit_TextFound_ChangeMade()
    {
        string examplePpt = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt, true);

        SlidePart slidePart = ppt.ForSlide(1);
        string beforeXml = slidePart.Slide.InnerXml;
        slidePart.ReplaceText("Slide 1").With("SLIDE 1");
        string afterXml = slidePart.Slide.InnerXml;

        Assert.NotEqual(beforeXml, afterXml);

        ppt.Close();
    }

    [Fact]
    public void SingleSlideEdit_TextFoundWithFormattingBreakingTheParagraph_ChangeMadeWithFormattingRemoved()
    {
        string examplePpt = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt, true);

        SlidePart slidePart = ppt.ForSlide(2);
        string beforeXml = slidePart.Slide.InnerXml;
        slidePart.ReplaceText("{with styling to break to the paragraph}").IgnoringStylingWith("looking good");
        string afterXml = slidePart.Slide.InnerXml;

        Assert.NotEqual(beforeXml, afterXml);

        ppt.Close();
    }

    [Fact]
    public void MultiSlideEdit_TextNotFound_NoChangeMade()
    {
        string examplePpt = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt, true);

        List<SlidePart> slideParts = ppt.ForAllSlides();
        List<string> beforeXmls = new();
        foreach (SlidePart sp in slideParts)
        {
            beforeXmls.Add(sp.Slide.InnerXml);
        }
        slideParts.ReplaceText("NOT FOUND").With("SOMETHING ELSE");
        List<string> afterXmls = new();
        foreach (SlidePart sp in slideParts)
        {
            afterXmls.Add(sp.Slide.InnerXml);
        }

        Assert.Equal(beforeXmls, afterXmls);

        ppt.Close();
    }
    
    [Fact]
    public void MultiSlideEdit_TextFound_ChangeMade()
    {
        string examplePpt = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt, true);

        List<SlidePart> slideParts = ppt.ForAllSlides();
        List<string> beforeXmls = new();
        foreach (SlidePart sp in slideParts)
        {
            beforeXmls.Add(sp.Slide.InnerXml);
        }
        slideParts.ReplaceText("This is slide ").With("Slide #");
        List<string> afterXmls = new();
        foreach (SlidePart sp in slideParts)
        {
            afterXmls.Add(sp.Slide.InnerXml);
        }

        Assert.NotEqual(beforeXmls, afterXmls);

        ppt.Close();
    }
}
