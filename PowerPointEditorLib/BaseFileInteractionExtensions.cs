namespace PowerPointEditorLib;

public static class BaseFileInteractionExtensions
{
    public static void Save(this PowerPointPresentation presentation)
    {
        presentation.GetDocument().Save();
    }

    public static void SaveAs(this PowerPointPresentation presentation, string filepath, bool overwrite = false)
    {
        if (overwrite
            && File.Exists(filepath))
        {
            try
            {
                File.Delete(filepath);
            }
            catch (Exception ex)
            {
                throw new AccessViolationException($"The file {filepath} already exists and cannot be deleted in order to save the new copy.", ex);
            }
        }

        presentation.GetDocument().SaveAs(filepath);
    }

    public static void Close(this PowerPointPresentation presentation)
    {
        presentation.GetDocument().Close();
    }
}
