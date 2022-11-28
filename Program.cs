using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        Example1();
    }

    static void Example1()
    {
        // In order to convert Excel to PDF, we just need to:
        //   1. Load XLS or XLSX file into ExcelFile object.
        //   2. Save ExcelFile object to PDF file.
        FontSettings.FontsBaseDirectory = ".";
        ExcelFile workbook = ExcelFile.Load("ComplexTemplate.xlsx");
        var cell = workbook.Worksheets[0].Cells[0, 0];
        var box1 = workbook.Worksheets[0].FormControls.AddCheckBox("Checked box 1", new AnchorCell(cell,true), cell.Column.GetWidth(LengthUnit.Point), cell.Row.GetHeight(LengthUnit.Point), LengthUnit.Point);
        box1.Checked = true;

        cell = workbook.Worksheets[0].Cells[1, 0];
        var box2 = workbook.Worksheets[0].FormControls.AddCheckBox("Checked box 2", new AnchorCell(cell, true), cell.Column.GetWidth(LengthUnit.Point), cell.Row.GetHeight(LengthUnit.Point), LengthUnit.Point);
        workbook.Save("Convert1.pdf", SaveOptions.PdfDefault);
    }
}