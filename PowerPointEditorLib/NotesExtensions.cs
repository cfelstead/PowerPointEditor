namespace PowerPointEditorLib;

public static class NotesExtensions
{
    public static IEnumerable<string> ReadNotes(this SlidePart slidePart)
    {
        var notesPart = slidePart.NotesSlidePart;

        if (notesPart is null)
            throw new NullReferenceException("The slide requested has no notes part.");

        var notes = notesPart.NotesSlide;
        IEnumerable<string> noteLines = notes.Descendants<OpenXmlDrawing.Paragraph>().Select(n => n.InnerText);
        return noteLines;
    }
}
