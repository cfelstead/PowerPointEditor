namespace PowerPointEditorLib;

public static class PictureExtensions
{
    public static Picture FindPictureWithAltText(this SlidePart slidePart, string altText)
    {
        IEnumerable<Picture> pictures = slidePart.Slide.Descendants<Picture>();
        Picture? picture = pictures.Where(p => p.NonVisualPictureProperties?.GetFirstChild<NonVisualDrawingProperties>()?.Description == altText).FirstOrDefault();

        if (picture is null)
            throw new KeyNotFoundException($"Cannot find a picture in the slide with the alternate text specified ('{altText}')");

        return picture;
    }

    public static void ReplaceAltTextWith(this Picture picture, string newAltText)
    {
        var altTextHolder = picture.NonVisualPictureProperties?.GetFirstChild<NonVisualDrawingProperties>();

        if (altTextHolder is null)
            throw new NullReferenceException("The holder for the alternate text could not be found");

        altTextHolder.Description = newAltText;
    }

    public static void ReplaceHyperlinkWith(this Picture picture, string newUrl)
    {
        var nonVisualProps = picture.NonVisualPictureProperties?.GetFirstChild<NonVisualDrawingProperties>()
            ?? throw new ArgumentException("The picture has no visual properties.");

        var hyperlinkId = nonVisualProps.HyperlinkOnClick?.Id
            ?? throw new ArgumentException("The picture has no hyperlink setup.");

        SlidePart? slidePart = picture.FindParentSlidePart()
            ?? throw new ArgumentException("The parent slide part found for the image.");




        // Create new hyperlink relationship
        HyperlinkRelationship newHyperlink = slidePart.AddHyperlinkRelationship(new Uri(newUrl, UriKind.Absolute), true);
        var clickHandler = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.HyperlinkType>().First(l => l.Id == hyperlinkId);
        HyperlinkRelationship currentHyperlink = slidePart.HyperlinkRelationships.Where(r => r.Id.Equals(hyperlinkId)).First();

        // Delete the old reference and replace with the new one
        slidePart.DeleteReferenceRelationship(currentHyperlink);
        clickHandler.Id = newHyperlink.Id;
    }

    public static void ReplaceImageWith(this Picture picture, string filepath)
    {
        if (File.Exists(filepath) == false)
            throw new FileNotFoundException($"The image cannot be found ('{filepath}')", filepath);
        
        ImagePart imagePart = GetRelatedImagePart(picture);

        using (var stream = File.OpenRead(filepath))
        {
            imagePart.FeedData(stream);
        }
    }

    public static ImagePart GetRelatedImagePart(this Picture picture)
    {
        string? imagePartRelationshipId = picture.BlipFill?.Blip?.Embed?.Value;

        if (imagePartRelationshipId is null)
            throw new NullReferenceException("Cannot find a relationship from the picture to the image part.");



        SlidePart? slidePart = picture.FindParentSlidePart();

        if (slidePart is null)
            throw new NullReferenceException("Cannot find the slide part associated with the picture.");



        ImagePart? imagePart = (ImagePart)slidePart.GetPartById(imagePartRelationshipId);

        if (imagePart is null)
            throw new NullReferenceException("Cannot find the image part with the corresponding relationship id.");
        
        

        return imagePart;
    }

    public static byte[] GetImageData(this ImagePart imagePart)
    {
        using MemoryStream ms = new();
        Stream ipStream = imagePart.GetStream(FileMode.Open, FileAccess.Read);
        ipStream.CopyTo(ms);
        ipStream.Close();
        ipStream.Dispose();

        return ms.ToArray();
    }

    private static SlidePart? FindParentSlidePart(this Picture picture)
    {
        SlidePart? slidePart = null;
        OpenXmlElement? parent = picture;

        do
        {

            parent = parent.Parent;

            if (parent is not null
                && parent is Slide s)
            {
                slidePart = s.SlidePart;
            }

        } while (slidePart is null
                 && parent is not null);


        return slidePart;
    }
}
