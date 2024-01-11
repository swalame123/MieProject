using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;

namespace MieProject.Controllers.Testing
{
    [Route("api/[controller]")]
    [ApiController]
    public class TestController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;

        public TestController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

        }
        [HttpGet("GetApprovedSpeakersData")]
        public IActionResult GetApprovedSpeakersData()
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:TestApprovedSpeakers").Value;
                long.TryParse(sheetId, out long parsedSheetId);
                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
                List<Dictionary<string, object>> sheetData = new List<Dictionary<string, object>>();
                List<string> columnNames = new List<string>();
                foreach (Column column in sheet.Columns)
                {
                    columnNames.Add(column.Title);
                }
                foreach (Row row in sheet.Rows)
                {
                    Dictionary<string, object> rowData = new Dictionary<string, object>();
                    for (int i = 0; i < row.Cells.Count && i < columnNames.Count; i++)
                    {
                        rowData[columnNames[i]] = row.Cells[i].Value;

                    }
                    sheetData.Add(rowData);
                }
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
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
            if (targetColumn != null && SpecialityColumn!= null)
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
        //[HttpGet("GetTyreData")]
        //public IActionResult GetTyreData(string val, string Tyre)
        //{
        //    SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
        //    string sheetId = configuration.GetSection("SmartsheetSettings:fmv").Value;
        //    long.TryParse(sheetId, out long parsedSheetId);
        //    Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);

        //    Column targetColumn = sheet.Columns.FirstOrDefault(column =>
        //    string.Equals(column.Title, "Speciality", StringComparison.OrdinalIgnoreCase));
        //    if (targetColumn != null)
        //    {
        //        // Find the row with the specified specialty
        //        Row targetRow = sheet.Rows.FirstOrDefault(row =>
        //            row.Cells.Any(cell => cell.ColumnId == targetColumn.Id && cell.Value.ToString() == val));
        //        if (targetRow != null)
        //        {
        //            // Retrieve the value of the specified tier column
        //            var tierValue = targetRow.Cells.FirstOrDefault(cell => cell.ColumnId == targetColumn.Id)?.Value;
        //            if (tierValue != null)
        //            {
        //                return Ok($"Tier value for {val} in {Tyre} column is: {tierValue}");
        //            }
        //            else
        //            {
        //                return NotFound($"Tier value not found for {val} in {Tyre} column.");
        //            }
        //        }
        //        else
        //        {
        //            return NotFound($"Specialty '{val}' not found.");
        //        }

        //    }
        //    else
        //    {
        //        return NotFound($"Column '{Tyre}' not found.");
        //    }




        //[HttpGet("GetDataBySheetId")]
        //public IActionResult GetDataBySheetId(string Value)
        //{
        //    try
        //    {
        //        SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
        //        PaginatedResult<Sheet> sheets = smartsheet.SheetResources.ListSheets(null, new PaginationParameters(true, null, null), null);
        //        List<long> sheetIds = new List<long>();
        //        foreach (Sheet sheet in sheets.Data)
        //        {
        //            sheetIds.Add((long)sheet.Id);
        //        }
        //        //foreach (var sheet in sheetIds)
        //        //{
        //        //    if(sheetId == sheet)
        //        //    {
        //        //        return Ok("Done");
        //        //    }

        //        //}
        //        if (sheetIds.Contains(sheetId))
        //        {
        //            Sheet sheet = smartsheet.SheetResources.GetSheet(sheetId, null, null, null, null, null, null, null);

        //            // Extract data from the sheet as key-value pairs
        //            List<Dictionary<string, object>> sheetData = new List<Dictionary<string, object>>();

        //            // Get column names
        //            List<string> columnNames = new List<string>();
        //            foreach (Column column in sheet.Columns)
        //            {
        //                columnNames.Add(column.Title);
        //            }

        //            foreach (Row row in sheet.Rows)
        //            {
        //                Dictionary<string, object> rowData = new Dictionary<string, object>();

        //                for (int i = 0; i < row.Cells.Count && i < columnNames.Count; i++)
        //                {
        //                    // Use column name as key
        //                    rowData[columnNames[i]] = row.Cells[i].Value;
        //                }

        //                sheetData.Add(rowData);
        //            }

        //            return Ok(sheetData);
        //            //return Ok("Sheet Found");
        //        }
        //        else
        //        {
        //            return BadRequest("Give Proper sheet Id");
        //        }



        //        return Ok(sheetId);
        //    }
        //    catch (Exception ex)
        //    {
        //        return BadRequest(ex.Message);
        //    }
        //}

        //}
    }
}
