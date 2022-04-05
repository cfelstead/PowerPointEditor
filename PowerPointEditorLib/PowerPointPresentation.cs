namespace PowerPointEditorLib;

public class PowerPointPresentation
{
    private PresentationDocument? _presentationDocument;
    private PresentationPart? _presentationPart;
    private string? _originalFilepath;

    /// <summary>
    /// Loads a presentation file from the local disk
    /// </summary>
    /// <param name="filepath">The full filepath to the powerpoint presentation</param>
    public PowerPointPresentation(string filepath)
    {
        if (File.Exists(filepath) == false)
            throw new FileNotFoundException("Cannot find the presentation file.", filepath);

        _originalFilepath = filepath;
        _presentationDocument = PresentationDocument.Open(_originalFilepath, true);
        _presentationPart = _presentationDocument.PresentationPart;
    }

    public PowerPointPresentation(string filepath, bool openReadOnly)
    {
        if (File.Exists(filepath) == false)
            throw new FileNotFoundException("Cannot find the presentation file.", filepath);

        _originalFilepath = filepath;
        _presentationDocument = PresentationDocument.Open(_originalFilepath, !openReadOnly);
        _presentationPart = _presentationDocument.PresentationPart;
    }

    internal PresentationDocument GetDocument() => _presentationDocument!;

    internal PresentationPart GetPresentationPart() => _presentationPart!;

    internal void ClosePresentation() => _presentationDocument!.Close();

    public bool PresentationPartLoaded() => _presentationPart is not null;
}
