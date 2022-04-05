namespace PowerPointEditorLib.Tests;

[Collection("UsesFile-Test")]
public class NotesTests
{
    [Fact]
    public void ReadNote_Empty_ThrowsError()
    {
        string examplePpt = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt, true);

        Assert.Throws<NullReferenceException>(() => ppt.ForSlide(1).ReadNotes());

        ppt.Close();
    }

    [Fact]
    public void ReadNote_Populated()
    {
        string examplePpt = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt ,true);

        var notes = ppt.ForSlide(2).ReadNotes();
        Assert.NotEmpty(notes);
        Assert.Equal("Slide 2 has notes", notes.ElementAt(0));
        Assert.Equal("This is line 3", notes.ElementAt(2));

        ppt.Close();
    }
}
