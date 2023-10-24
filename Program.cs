using epplus_excel_spreadsheets.Models;
using epplus_excel_spreadsheets.Services;
using OfficeOpenXml;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

var excelFileName = "Data/EPPlus.xlsx";
var worksheetName = "Customers";

List<Customer> writeCustomers =
[
    new() { Name = "Peter Parker", Email = "peter.parker@marvel.com" },
    new() { Name = "Ben Parker", Email = "ben.parker@marvel.com" },
    new() { Name = "Mary Jane", Email = "mary.jane@marvel.com" }
];

EpplusService.WriteDataToExcel(excelFileName, worksheetName, writeCustomers);

var readCustomers = EpplusService.ReadDataFromExcel(excelFileName, worksheetName);

Console.ForegroundColor = ConsoleColor.Magenta;
foreach (var currentCustomer in readCustomers)
{
    Console.WriteLine($"{ currentCustomer.Id } - { currentCustomer.Name } - { currentCustomer.Email }");
}