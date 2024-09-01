using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using ReportGenerationInExcel.Models;

namespace ReportGenerationInExcel.Controllers;

[Route("api/[controller]")]
[ApiController]
public class ReportsGenerationController : ControllerBase
{
    [HttpGet("generate-report")]
    public async Task<IActionResult> GenerateExcel()
    {
        var data = new List<ReportData>
        {
            new ReportData { Id = 1, Name = "Amit Kumar", Age = 32, Email = "amit.kumarabc@example.com", Address = "12/34, Sector 15, Noida, Uttar Pradesh", PhoneNumber = "97123-456780" },
            new ReportData { Id = 2, Name = "Sneha Patel", Age = 29, Email = "sneha.patelabc@example.com", Address = "56, Gali No. 7, Delhi", PhoneNumber = "97765-432100" },
            new ReportData { Id = 3, Name = "Rajesh Sharma", Age = 45, Email = "rajesh.sharmaabc@example.com", Address = "78, Green Park, Mumbai, Maharashtra", PhoneNumber = "97234-567890" },
            new ReportData { Id = 4, Name = "Mohammed Ali", Age = 40, Email = "mohammed.aliabc@example.com", Address = "22/45, Jamia Nagar, Delhi", PhoneNumber = "97123-876540" },
            new ReportData { Id = 5, Name = "Tahir Ansari", Age = 35, Email = "tahir.ansariabc@example.com", Address = "8-9, Park Road, Bangalore, Karnataka", PhoneNumber = "97001-234560" },
            new ReportData { Id = 6, Name = "Imran Sheikh", Age = 28, Email = "imran.sheikhabc@example.com", Address = "10/12, Saki Naka, Mumbai, Maharashtra", PhoneNumber = "97234-567890" }
        };


        ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // LicenseContext.NonCommercial


        // Generate Excel file
        using var package = new ExcelPackage();
        var worksheet = package.Workbook.Worksheets.Add("Sheet1");

        // Set column headers
        worksheet.Cells[1, 1].Value = "ID";
        worksheet.Cells[1, 2].Value = "Name";
        worksheet.Cells[1, 3].Value = "Age";
        worksheet.Cells[1, 4].Value = "Email";
        worksheet.Cells[1, 5].Value = "Address";
        worksheet.Cells[1, 6].Value = "PhoneNumber";

        // Populate data
        for (int i = 0; i < data.Count; i++)
        {
            worksheet.Cells[i + 2, 1].Value = data[i].Id;
            worksheet.Cells[i + 2, 2].Value = data[i].Name;
            worksheet.Cells[i + 2, 3].Value = data[i].Age;
            worksheet.Cells[i + 2, 4].Value = data[i].Email;
            worksheet.Cells[i + 2, 5].Value = data[i].Address;
            worksheet.Cells[i + 2, 6].Value = data[i].PhoneNumber;
        }

        // Set the content type and file name
        var stream = new MemoryStream();
        await package.SaveAsAsync(stream);
        stream.Position = 0;
        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Report.xlsx");
    }
}
