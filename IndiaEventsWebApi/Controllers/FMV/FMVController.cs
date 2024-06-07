using IndiaEventsWebApi.Helper;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;

namespace IndiaEventsWebApi.Controllers.FMV
{
    [Route("api/[controller]")]
    [ApiController]
    public class FMVController : ControllerBase
    {

        private readonly string accessToken;
        private readonly IConfiguration configuration;
        private readonly SmartsheetClient smartsheet;
        public FMVController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
            smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
        }

        [HttpGet("GetfmvColumnValue")]
        public IActionResult GetfmvColumnValue(string specialty, string columnTitle)
        {
            try
            {


                int defaultval = 0;
                //SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:fmv").Value;
                //  long.TryParse(sheetId, out long parsedSheetId);
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);

                Column SpecialityColumn = sheet.Columns.FirstOrDefault(column => string.Equals(column.Title, "Speciality", StringComparison.OrdinalIgnoreCase));
                Column targetColumn = sheet.Columns.FirstOrDefault(column => string.Equals(column.Title, columnTitle, StringComparison.OrdinalIgnoreCase));
                if (targetColumn != null && SpecialityColumn != null)
                {
                    Row targetRow = sheet.Rows.FirstOrDefault(row => row.Cells.Any(cell => cell.ColumnId == SpecialityColumn.Id && cell.Value.ToString() == specialty));

                    if (targetRow != null)
                    {
                        var columnValue = targetRow.Cells.FirstOrDefault(cell => cell.ColumnId == targetColumn.Id)?.Value;
                        if (columnValue != null)
                        {
                            return Ok(columnValue);
                        }
                        else
                        {
                            return Ok(defaultval);
                        }
                    }
                    else
                    {
                        return Ok(defaultval);
                    }
                }
                else
                {
                    return Ok(defaultval);
                }
            }
            catch (Exception ex)
            {

                return BadRequest(new
                {
                    Message = ex.Message + "------" + ex.StackTrace
                });
            }
        }


        [HttpGet("GetFMVData")]
        public IActionResult GetFMVData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:fmv").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
    }
}
