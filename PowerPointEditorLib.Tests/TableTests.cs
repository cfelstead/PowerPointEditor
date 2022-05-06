using DocumentFormat.OpenXml.Drawing;

namespace PowerPointEditorLib.Tests;

[Collection("UsesFile-Test")]
public class TableTests
{
    [Fact]
    public void FindTable_NoTable_ThrowsError()
    {
        string examplePpt = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt, true);

        Assert.Throws<NullReferenceException>(() => ppt.ForSlide(1).FindTable(1));

        ppt.Close();
    }

    [Fact]
    public void FindTable_TableNumberIncorrect_ThrowsError()
    {
        string examplePpt = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt, true);

        Assert.Throws<NullReferenceException>(() => ppt.ForSlide(5).FindTable(2));

        ppt.Close();
    }

    [Fact]
    public void FindTable_TableReturned()
    {
        string examplePpt = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt, true);

        Assert.NotNull(ppt.ForSlide(5).FindTable(1));

        ppt.Close();
    }

    [Fact]
    public void FindRow_RowNumberIncorrect_ThrowsError()
    {
        string examplePpt = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt, true);

        Table tbl = ppt.ForSlide(5).FindTable(1);

        Assert.Throws<NullReferenceException>(() => tbl.FindRow(99));

        ppt.Close();
    }

    [Fact]
    public void FindRow_RowReturned()
    {
        string examplePpt = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt, true);

        Assert.NotNull(ppt.ForSlide(5).FindTable(1).FindRow(2));

        ppt.Close();
    }

    [Fact]
    public void FindCell_CellNumberIncorrect_ThrowsError()
    {
        string examplePpt = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt, true);

        TableRow row = ppt.ForSlide(5).FindTable(1).FindRow(2);

        Assert.Throws<NullReferenceException>(() => row.FindCell(99));

        ppt.Close();
    }

    [Fact]
    public void FindCell_CellReturned()
    {
        string examplePpt = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt, true);

        Assert.NotNull(ppt.ForSlide(5).FindTable(1).FindRow(2).FindCell(2));

        ppt.Close();
    }

    [Fact]
    public void Cell_TextChanged()
    {
        string examplePpt = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                         "ppt-examples",
                                         "test.pptx");

        var ppt = new PowerPointPresentation(examplePpt, true);


        TableCell cell = ppt.ForSlide(5).FindTable(1).FindRow(2).FindCell(1);
        string beforeCellInnerText = cell.InnerText;

        cell.ReplaceCellText("S7 Puebla R8", "Race 01");
        string afterCellInnerText = cell.InnerText;

        Assert.NotEqual(beforeCellInnerText, afterCellInnerText);

        ppt.Close();
    }
}
