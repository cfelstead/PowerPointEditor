using PowerPointEditorLib;

string pptxFilePath = @"___filepath___";
var pptx = new PowerPointPresentation(pptxFilePath);

pptx.ForSlide(3)
            .ReplaceText("{BrandName}")
            .IgnoringStylingWith("FooBar");

pptx.SaveAs(pptxFilePath.Replace(".pptx", "MODIFIED.pptx"), true);