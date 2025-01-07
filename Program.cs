using System.Globalization;
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
                foreach (var cell in headerRow.Elements<Cell>())
                {
                    var headerValue = GetCellValue(cell, sst, workbookPart);
                    var columnRef = GetColumnReference(cell.CellReference);

                    columnInfos.Add(new ColumnInfo
                    {
                        Header = headerValue ?? string.Empty,
                        ColumnReference = columnRef
                    });

                    excelData.Headers.Add(headerValue ?? string.Empty);
                }

            // Veri satırlarını oku
            var dataRows = sheetData.Elements<Row>()
                .Where(r => r.RowIndex > headerRowIndex);

            foreach (var row in dataRows)
            {
                var rowData = new Dictionary<string, string>();

                // Önce tüm başlıklar için boş değer ata
                foreach (var header in excelData.Headers) rowData[header] = string.Empty;

                // Her hücreyi doğru header ile eşleştir
                foreach (var cell in row.Elements<Cell>())
                {
                    var columnRef = GetColumnReference(cell.CellReference);
                    var columnInfo = columnInfos.FirstOrDefault(c => c.ColumnReference == columnRef);

                    if (columnInfo != null)
                    {
                        var value = GetCellValue(cell, sst, workbookPart) ?? string.Empty;
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

    private static string GetCellValue(Cell cell, SharedStringTable sst, WorkbookPart workbookPart)
    {
        if (cell.CellValue == null)
            return string.Empty;

        string value = cell.CellValue.Text;

        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            return sst.ElementAt(int.Parse(value.Trim())).InnerText.Trim();
        }

        if (cell.StyleIndex != null)
        {
            CellFormat cellFormat = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Elements<CellFormat>().ElementAt((int)cell.StyleIndex.Value);
                
            //14: Gün/Ay/Yıl biçiminde tarih (d/m/yyyy)
            //15: Ay/Gün/Yıl (m/d/yy)
            //16: Ay/Yıl (m/yy)
            //17, 18, 19, 20: Diğer özel tarih formatları
            if (cellFormat.NumberFormatId > 13 && cellFormat.NumberFormatId < 20) // Tarih formatı kontrolü
            {
                var excelDateValue = double.Parse(value, CultureInfo.InvariantCulture);
                    
                DateTime baseDate = new DateTime(1900, 1, 1); // Excel'in tarih başlangıcı

                // Excel'in başlangıç tarihine gün ekleyerek tarihi hesapla
                DateTime date = baseDate.AddDays(excelDateValue - 2); // -2 eklenmesinin sebebi: Excel'deki tarihlerde 1 Ocak 1900 = 1 gün sayılıyor ve ayrıca 29 Şubat 1900 hatası var

                return date.ToString(CultureInfo.InvariantCulture);
            }

            return value.Trim();
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
    
    public static DateTime? DateTimeParse(string dateStr)
    {
        if (string.IsNullOrEmpty(dateStr))
            return null;

        // Önce TryParseExact ile belirtilen formatları dene
        DateTime sonuc;
        if (DateTime.TryParseExact(
                dateStr,
                DateFormats,
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out sonuc))
        {
            return sonuc;
        }

        // Eğer başarısız olursa, genel Parse'ı dene
        if (DateTime.TryParse(
                dateStr,
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out sonuc))
        {
            return sonuc;
        }

        return null;
    }
   
    private static readonly string[] DateFormats = new[] {
        "M/d/yyyy",
        "M/d/yyyy HH:mm:ss",
        "MM/dd/yyyy",
        "MM/dd/yyyy HH:mm:ss",
        "d.M.yyyy",
        "dd.MM.yyyy",
        "d.M.yyyy HH:mm:ss",
        "dd.MM.yyyy HH:mm:ss"
    };
}