using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelRead.Models;

namespace ExcelRead;

static class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("Excel Read Started.");

        string path = "";

        var excelData = ReadExcelFile(path);

        var result = GetExcelDataList(excelData);

        Console.WriteLine("Excel Read Completed.");
    }

    private static ExcelData ReadExcelFile(string filePath, int headerRowIndex = 1)
    {
        var excelData = new ExcelData();
        var columnInfos = new List<ColumnInfo>();

        using (var document = SpreadsheetDocument.Open(filePath, false))
        {
            var workbookPart = document.WorkbookPart;
            var sheet = workbookPart?.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault();
            var worksheetPart = (WorksheetPart)workbookPart?.GetPartById(sheet?.Id);

            var sheetData = worksheetPart?.Worksheet.Elements<SheetData>().FirstOrDefault();
            SharedStringTablePart? sstpart = workbookPart?.GetPartsOfType<SharedStringTablePart>().First();
            SharedStringTable sst = sstpart.SharedStringTable;

            Row headerRow = sheetData?.Elements<Row>().FirstOrDefault(r => r.RowIndex == headerRowIndex);

            if (headerRow != null)
            {
                foreach (Cell cell in headerRow.Elements<Cell>())
                {
                    string? headerValue = GetCellValue(cell, sst);
                    string columnRef = GetColumnReference(cell.CellReference);

                    columnInfos.Add(new ColumnInfo
                    {
                        Header = headerValue ?? string.Empty,
                        ColumnReference = columnRef
                    });

                    excelData.Headers.Add(headerValue ?? string.Empty);
                }
            }

            // Veri satırlarını oku
            var dataRows = sheetData.Elements<Row>()
                .Where(r => r.RowIndex > headerRowIndex);

            foreach (Row row in dataRows)
            {
                var rowData = new Dictionary<string, string?>();

                // Önce tüm başlıklar için boş değer ata
                foreach (var header in excelData.Headers)
                {
                    rowData[header] = string.Empty;
                }

                // Her hücreyi doğru header ile eşleştir
                foreach (Cell cell in row.Elements<Cell>())
                {
                    string columnRef = GetColumnReference(cell.CellReference);
                    var columnInfo = columnInfos.FirstOrDefault(c => c.ColumnReference == columnRef);

                    if (columnInfo != null)
                    {
                        string? value = GetCellValue(cell, sst) ?? string.Empty;
                        rowData[columnInfo.Header] = value;
                    }
                }

                excelData.Rows.Add(rowData);
            }
        }

        return excelData;
    }

    private static string GetColumnReference(string cellReference)
    {
        // A1 -> A, B2 -> B gibi sütun referansını ayıkla
        return new string(cellReference.TakeWhile(c => !char.IsDigit(c)).ToArray());
    }

    private static string GetCellValue(Cell cell, SharedStringTable sst)
    {
        if (cell.CellValue == null)
            return string.Empty;

        string? value = cell.CellValue.Text;

        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            return sst.ElementAt(int.Parse(value)).InnerText;
        }

        return value;
    }


    private static List<DataModel> GetExcelDataList(ExcelData excelData)
    {
        var result = new List<DataModel>();
        foreach (var row in excelData.Rows)
            result.Add(new DataModel
            {
                Firstname = !string.IsNullOrEmpty(row["Firstname"]) ? row["Firstname"] : null,
                Lastname = !string.IsNullOrEmpty(row["Lastname"]) ? row["Lastname"] : null,
                Email = !string.IsNullOrEmpty(row["Email"]) ? row["Email"] : null
            });

        return result;
    }
    
}