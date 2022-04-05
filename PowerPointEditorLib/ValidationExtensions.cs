namespace PowerPointEditorLib;

internal static class ValidationExtensions
{
    internal static void EnsurePresenationPartIsValid(this PowerPointPresentation presentation)
    {
        if (presentation.GetPresentationPart() is null)
            throw new NullReferenceException($"The presentation part of your document could not be loaded.");
    }
}
