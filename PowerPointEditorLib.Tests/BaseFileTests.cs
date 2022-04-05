namespace PowerPointEditorLib.Tests;

[Collection("UsesFile-Test")]
public class BaseFileTests
{
    [Fact]
    public void LoadingFromDisk_BasicLoadWorks()
    {
        string examplePpt = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt, true);

        Assert.True(ppt.PresentationPartLoaded());

        ppt.Close();
    }

    [Fact]
    public void LoadingFromDisk_NotFoundErrorProduced()
    {
        string examplePpt = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "notfound.pptx");

        Assert.Throws<FileNotFoundException>(() => new PowerPointPresentation(examplePpt));
    }
}
