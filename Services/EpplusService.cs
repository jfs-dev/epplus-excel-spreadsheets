using epplus_excel_spreadsheets.Models;
using OfficeOpenXml;

namespace epplus_excel_spreadsheets.Services;

public static class EpplusService
{
    private static ExcelPackage CreateExcelPackage(string filePath)
    {
        var fileInfo = new FileInfo(filePath);

        return new ExcelPackage(fileInfo);
    }

    private static void WriteDataToWorksheet(ExcelWorksheet worksheet, List<Customer> customers)
    {
        int lastRow = worksheet.Dimension.End.Row;

        foreach (var currentCustomer in customers)
        {
            lastRow++;
            
            worksheet.Cells[lastRow, 1].Value = currentCustomer.Id;
            worksheet.Cells[lastRow, 2].Value = currentCustomer.Name;
            worksheet.Cells[lastRow, 3].Value = currentCustomer.Email;
        }
    }

    public static void WriteDataToExcel(string filePath, string worksheetName, List<Customer> customers)
    {
        using var package = CreateExcelPackage(filePath);

        if (File.Exists(filePath))
        {
            var existingWorksheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == worksheetName);
            
            if (existingWorksheet is not null)
            {
                WriteDataToWorksheet(existingWorksheet, customers);
                package.Save();

                return;
            }
        }

        var worksheet = package.Workbook.Worksheets.Add(worksheetName);

        worksheet.Cells["A1"].Value = "Id";
        worksheet.Cells["B1"].Value = "Name";
        worksheet.Cells["C1"].Value = "Email";

        WriteDataToWorksheet(worksheet, customers);
        package.Save();
    }

    public static List<Customer> ReadDataFromExcel(string filePath, string worksheetName)
    {
        using var package = CreateExcelPackage(filePath);

        var worksheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == worksheetName) ?? throw new Exception($"Planilha '{ worksheetName }' não encontrada.");
        int rowCount = worksheet.Dimension.End.Row - worksheet.Dimension.Start.Row + 1;

        var customers = new List<Customer>();

        for (int row = 2; row <= rowCount; row++)
        {
            customers.Add(new Customer
            {
                Id = Guid.Parse(worksheet.Cells[row, 1].Text),
                Name = worksheet.Cells[row, 2].Text,
                Email = worksheet.Cells[row, 3].Text
            });
        }

        return customers;
    }    
}