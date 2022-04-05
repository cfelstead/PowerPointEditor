namespace PowerPointEditorLib;

public static class SlideSelectionExtensions
{
    public static List<SlidePart> ForAllSlides(this PowerPointPresentation presentation)
    {
        presentation.EnsurePresenationPartIsValid();
        var presentationPart = presentation.GetPresentationPart();



        List<SlidePart> slidePartsList = new();

        var slideIds = presentationPart.Presentation.SlideIdList?.ChildElements ?? throw new NullReferenceException("The SlideIdList in your presentation is empty");
        foreach (var slideId in slideIds)
        {
            if (slideId is SlideId)
            {
                var relationshipId = ((SlideId)slideId).RelationshipId;
                slidePartsList.Add((SlidePart)presentationPart.GetPartById(relationshipId!));
            }
        }

        return slidePartsList;
    }

    public static SlidePart ForSlide(this PowerPointPresentation presentation, int slideNumber)
    {
        var allSlides = ForAllSlides(presentation);
        var slidePart = allSlides.ElementAt(slideNumber - 1);
        return slidePart;
    }
}
