using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;

namespace MieProject.Controllers.FMV
{
    [Route("api/[controller]")]
    [ApiController]
    public class FMVController : ControllerBase
    {

        private readonly string accessToken;
        private readonly IConfiguration configuration;

        public FMVController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

        }

        [HttpGet("GetfmvColumnValue")]
        public IActionResult GetfmvColumnValue(string specialty, string columnTitle)
        {
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            string sheetId = configuration.GetSection("SmartsheetSettings:fmv").Value;
            long.TryParse(sheetId, out long parsedSheetId);
            Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);

            Column SpecialityColumn = sheet.Columns.FirstOrDefault(column =>
           string.Equals(column.Title, "Speciality", StringComparison.OrdinalIgnoreCase));
            Column targetColumn = sheet.Columns.FirstOrDefault(column =>
           string.Equals(column.Title, columnTitle, StringComparison.OrdinalIgnoreCase));
            if (targetColumn != null && SpecialityColumn != null)
            {
                // Find the row with the specified speciality
                Row targetRow = sheet.Rows.FirstOrDefault(row =>
                    row.Cells.Any(cell => cell.ColumnId == SpecialityColumn.Id && cell.Value.ToString() == specialty));

                if (targetRow != null)
                {
                    // Retrieve the value of the specified column for the given speciality
                    var columnValue = targetRow.Cells.FirstOrDefault(cell => cell.ColumnId == targetColumn.Id)?.Value;
                    if (columnValue != null)
                    {
                        return Ok(columnValue);
                    }
                    else
                    {
                        return NotFound($"Value not found for {specialty} in {columnTitle} column.");
                    }
                }
                else
                {
                    return NotFound($"Speciality '{specialty}' not found.");
                }
            }
            else
            {
                return NotFound($"Column '{columnTitle}' not found.");
            }
        }
    }
}
