using DocumentFormat.OpenXml.Presentation;

namespace PowerPointEditorLib.Tests;

[Collection("UsesFile-Test")]
public class PictureTests
{
    [Fact]
    public void Picture_WithAltText_Exists()
    {
        string examplePpt = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt);

        var picture = ppt.ForSlide(2).FindPictureWithAltText("Slide 2 Image");
        Assert.IsType<Picture>(picture);
        Assert.NotNull(picture);

        ppt.Close();
    }

    [Fact]
    public void Picture_WithAltText_NotExists()
    {
        string examplePpt = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt, true);

        Assert.Throws<KeyNotFoundException>(() => ppt.ForSlide(2).FindPictureWithAltText("Does not exist"));

        ppt.Close();
    }

    [Fact]
    public void Picture_WithHyperlinkChanged_Works()
    {
        string examplePpt = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");
        string originalAltText = "Slide 2 Image";
        string newUrl = "http://www.futuressport.com";


        string destructiveReplacementPpt = examplePpt.Replace("test.pptx", "test_replaceimage.pptx");
        if (File.Exists(destructiveReplacementPpt)) File.Delete(destructiveReplacementPpt);
        File.Copy(examplePpt, destructiveReplacementPpt);

        var changablePpt = new PowerPointPresentation(destructiveReplacementPpt);

        Picture picture = changablePpt.ForSlide(2).FindPictureWithAltText(originalAltText);
        string beforeXml = picture.InnerXml;

        picture.ReplaceHyperlinkWith(newUrl);

        Picture picture2 = changablePpt.ForSlide(2).FindPictureWithAltText(originalAltText);
        string afterXml = picture2.InnerXml;

        Assert.NotEqual(beforeXml, afterXml);
        
        changablePpt.Close();
        File.Delete(destructiveReplacementPpt);
    }

    [Fact]
    public void Picture_ReplaceAltText_Works()
    {
        string examplePpt = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");
        string originalAltText = "Slide 2 Image";
        string newAltText = "New Alt Text";

        var ppt = new PowerPointPresentation(examplePpt, true);

        ppt.ForSlide(2).FindPictureWithAltText(originalAltText).ReplaceAltTextWith(newAltText);
        
        Assert.Throws<KeyNotFoundException>(() => ppt.ForSlide(2).FindPictureWithAltText(originalAltText));
        Picture replacementPic = ppt.ForSlide(2).FindPictureWithAltText(newAltText);
        Assert.IsType<Picture>(replacementPic);
        Assert.NotNull(replacementPic);

        ppt.Close();
    }

    [Fact]
    public void Picture_ReplaceImage_Works()
    {
        string examplePpt = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");
        string originalAltText = "Slide 2 Image";
        string newImage = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "testing-assets",
                                         "multi-color-fabric-texture-samples_1373-435.jpg");


        string destructiveReplacementPpt = examplePpt.Replace("test.pptx", "test_replaceimage.pptx");
        if (File.Exists(destructiveReplacementPpt)) File.Delete(destructiveReplacementPpt);
        File.Copy(examplePpt, destructiveReplacementPpt);

        var changablePpt = new PowerPointPresentation(destructiveReplacementPpt);

        Picture picture = changablePpt.ForSlide(2).FindPictureWithAltText(originalAltText);
        string beforeXml = picture.InnerXml;
        ImagePart beforeImagePart = picture.GetRelatedImagePart();
        byte[] beforeImage = beforeImagePart.GetImageData();
        picture.ReplaceImageWith(newImage);
        
        Picture picture2 = changablePpt.ForSlide(2).FindPictureWithAltText(originalAltText);
        string afterXml = picture2.InnerXml;
        ImagePart afterImagePart = picture.GetRelatedImagePart();
        byte[] afterImage = afterImagePart.GetImageData();
                
        Assert.Equal(beforeXml, afterXml);
        Assert.Equal(beforeImagePart, afterImagePart);
        Assert.NotEqual(beforeImage, afterImage); // This is the child element that will show the change
        
        changablePpt.Close();
        File.Delete(destructiveReplacementPpt);
    }

    [Fact]
    public void Picture_ReplaceAllImages_Works()
    {
        string examplePpt = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");
        string originalAltText = "ReplaceAllImages";
        string newImage = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "testing-assets",
                                         "multi-color-fabric-texture-samples_1373-435.jpg");


        string destructiveReplacementPpt = examplePpt.Replace("test.pptx", "test_replaceimage.pptx");
        if (File.Exists(destructiveReplacementPpt)) File.Delete(destructiveReplacementPpt);
        File.Copy(examplePpt, destructiveReplacementPpt);

        var changablePpt = new PowerPointPresentation(destructiveReplacementPpt);

        List<Picture> pictures = changablePpt.ForAllSlides().FindPictureWithAltText(originalAltText);
        Picture picture = pictures[0];
        string beforeXml = picture.InnerXml;
        ImagePart beforeImagePart = picture.GetRelatedImagePart();
        byte[] beforeImage = beforeImagePart.GetImageData();
        
        picture.ReplaceImageWith(newImage);

        List<Picture> pictures2 = changablePpt.ForAllSlides().FindPictureWithAltText(originalAltText);
        Picture picture2 = pictures2[0];
        string afterXml = picture2.InnerXml;
        ImagePart afterImagePart = picture.GetRelatedImagePart();
        byte[] afterImage = afterImagePart.GetImageData();

        Assert.Equal(beforeXml, afterXml);
        Assert.Equal(beforeImagePart, afterImagePart);
        Assert.NotEqual(beforeImage, afterImage); // This is the child element that will show the change

        changablePpt.Close();
        File.Delete(destructiveReplacementPpt);
    }
}
