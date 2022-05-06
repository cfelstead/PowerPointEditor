using DocumentFormat.OpenXml.Drawing;

namespace PowerPointEditorLib;

public static class TableExtensions
{
    public static Table FindTable(this SlidePart slidePart, int tableNumber)
    {
        IEnumerable<Table> tables = slidePart.Slide.Descendants<Table>();
        
        if (tables.Any() == false)
            throw new NullReferenceException("Cannot find any tables on the slide");


        if (tableNumber > tables.Count())
            throw new NullReferenceException($"Cannot find a table at the given position ('{tableNumber}')");


        Table tbl = tables.ElementAt(tableNumber - 1);
        
        if (tbl is null)
            throw new NullReferenceException($"Cannot find a table at the given position ('{tableNumber}')");



        return tbl;
    }

    public static TableRow FindRow(this Table table, int rowNumber)
    {
        IEnumerable<TableRow> rows = table.Descendants<TableRow>();

        if (rows.Any() == false)
            throw new NullReferenceException("Cannot find any rows in this table");


        if (rowNumber > rows.Count())
            throw new NullReferenceException($"Cannot find a row at the given position ('{rowNumber}')");


        TableRow row = rows.ElementAt(rowNumber - 1);

        if (row is null)
            throw new NullReferenceException($"Cannot find a row at the given position ('{rowNumber}')");



        return row;
    }

    public static TableCell FindCell(this TableRow row, int cellNumber)
    {
        IEnumerable<TableCell> cells = row.Descendants<TableCell>();

        if (cells.Any() == false)
            throw new NullReferenceException("Cannot find any cells in this row");


        if (cellNumber > cells.Count())
            throw new NullReferenceException($"Cannot find a cell at the given position ('{cellNumber}')");


        TableCell cell = cells.ElementAt(cellNumber - 1);

        if (cell is null)
            throw new NullReferenceException($"Cannot find a cell at the given position ('{cellNumber}')");



        return cell;
    }

    public static void ReplaceCellText(this TableCell cell, string searchForText, string replaceWithText)
    {
        List<OpenXmlDrawing.Paragraph> paragraphs = new();
        foreach (var textBody in cell.Descendants<OpenXmlDrawing.TextBody>())
        {
            foreach (var paragraph in textBody.Descendants<OpenXmlDrawing.Paragraph>())
            {
                paragraphs.Add(paragraph);
            }
        }



        foreach (OpenXmlDrawing.Paragraph para in paragraphs)
        {
            if (para.InnerText.Contains(searchForText))
            {
                var runInParagraph = para.Descendants<DocumentFormat.OpenXml.Drawing.Run>();
                foreach (var run in runInParagraph)
                {
                    if (run is not null
                        && run.Text is not null)
                    {
                        run.Text.Text = run.Text.Text.Replace(searchForText, replaceWithText);
                    }
                }
            }
        }
    }
}
