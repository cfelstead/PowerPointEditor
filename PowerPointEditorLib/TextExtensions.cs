namespace PowerPointEditorLib;

public static class TextExtensions
{
    public static (List<SlidePart> slideParts, string searchText) ReplaceText(this List<SlidePart> slideParts, string searchText)
    {
        return (slideParts, searchText);
    }

    public static (SlidePart slidePart, string searchText) ReplaceText(this SlidePart slidePart, string searchText)
    {
        return (slidePart, searchText);
    }

    public static void With(this (List<SlidePart> slideParts, string searchText) input, string replacementText)
    {
        foreach (SlidePart slidepart in input.slideParts)
        {
            With((slidepart, input.searchText), replacementText);
        }
    }

    public static void With(this (SlidePart slidePart, string searchText) input, string replacementText)
    {
        IEnumerable<OpenXmlDrawing.Paragraph> paragraphs = input.slidePart.Slide.Descendants<OpenXmlDrawing.Paragraph>();
        foreach (OpenXmlDrawing.Paragraph para in paragraphs)
        {
            if (para.InnerText.Contains(input.searchText))
            {
                var runInParagraph = para.Descendants<DocumentFormat.OpenXml.Drawing.Run>();
                foreach (var run in runInParagraph)
                {
                    if (run is not null
                        && run.Text is not null)
                    {
                        run.Text.Text = run.Text.Text.Replace(input.searchText, replacementText);
                    }
                }
            }
        }
    }
}
