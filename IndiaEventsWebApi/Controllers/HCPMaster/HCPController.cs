using IndiaEventsWebApi.Models;
using IndiaEventsWebApi.Models.MasterSheets;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using Smartsheet.Api.OAuth;

namespace IndiaEventsWebApi.Controllers.HCPMaster
{
    [Route("api/[controller]")]
    [ApiController]
    public class HCPController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;

        public HCPController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

        }
        [HttpGet("GetHCPDataUsingNameAndMISCode")]
        public IActionResult GetHCPDataUsingNameAndMISCode(string Name, string misCode)
        {
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            string[] sheetIds = {
                configuration.GetSection("SmartsheetSettings:HcpMaster").Value,
                configuration.GetSection("SmartsheetSettings:HcpMaster1").Value,
                configuration.GetSection("SmartsheetSettings:HcpMaster2").Value,
                configuration.GetSection("SmartsheetSettings:HcpMaster3").Value,
                configuration.GetSection("SmartsheetSettings:HcpMaster4").Value
            };
            foreach (string i in sheetIds)
            {
                long.TryParse(i, out long p);
                Sheet sheeti = smartsheet.SheetResources.GetSheet(p, null, null, null, null, null, null, null);
                Column hcpNameColumn = sheeti.Columns.FirstOrDefault(column => column.Title == "HCPName");
                Column misCodeColumn = sheeti.Columns.FirstOrDefault(column => column.Title == "MISCode");

                if (hcpNameColumn != null && misCodeColumn != null)
                {
                    Row existingRow = sheeti.Rows.FirstOrDefault(row =>
                        row.Cells != null &&
                        row.Cells.Any(cell =>
                            cell.ColumnId == hcpNameColumn.Id && cell.Value != null && cell.Value.ToString() == Name
                        ) &&
                        row.Cells.Any(cell =>
                            cell.ColumnId == misCodeColumn.Id && cell.Value != null && cell.Value.ToString() == misCode
                        )
                    );
                    if (existingRow != null)
                    {
                        // Both Name and MISCode are present in the same row, return success
                        return Ok("True");
                    }

                }
            }
            return Ok("False");
        }




        [HttpPost("PostHcpData1")]
        public IActionResult PostHcpData1(HCPMaster1 formDataList)
        {
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            string[] sheetIds = {
                configuration.GetSection("SmartsheetSettings:HcpMaster").Value,
                configuration.GetSection("SmartsheetSettings:HcpMaster1").Value,
                configuration.GetSection("SmartsheetSettings:HcpMaster2").Value,
                configuration.GetSection("SmartsheetSettings:HcpMaster3").Value,
                configuration.GetSection("SmartsheetSettings:HcpMaster4").Value
            };
            foreach (string i in sheetIds)
            {
                long.TryParse(i, out long p);
                Sheet sheeti = smartsheet.SheetResources.GetSheet(p, null, null, null, null, null, null, null);

                // Check if any row contains the same MISCode
                //Row existingRow = sheeti.Rows.FirstOrDefault(row => row.Cells.Any(cell => cell.Value.ToString() == formDataList.MISCode));
                Row existingRow = sheeti.Rows.FirstOrDefault(row =>
                    row.Cells != null &&
                    row.Cells.Any(cell => cell.Value != null && cell.Value.ToString() == formDataList.MISCode));

                if (existingRow != null)
                {
                    // Data with the same MISCode already exists, return a response
                    return BadRequest("Data with the same MISCode already exists.");
                }
            }
            string sheetId = configuration.GetSection("SmartsheetSettings:HcpMaster1").Value;
            long.TryParse(sheetId, out long parsedSheetId);
            Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
            var newRow = new Row();
            newRow.Cells = new List<Cell>();
            newRow.Cells.Add(new Cell
            {
                ColumnId = GetColumnIdByName(sheet, "FirstName"),
                Value = formDataList.FirstName
            });
            newRow.Cells.Add(new Cell
            {
                ColumnId = GetColumnIdByName(sheet, "LastName"),
                Value = formDataList.LastName
            });
            newRow.Cells.Add(new Cell
            {
                ColumnId = GetColumnIdByName(sheet, "HCPName"),
                Value = formDataList.HCPName
            });
            newRow.Cells.Add(new Cell
            {
                ColumnId = GetColumnIdByName(sheet, "GO/Non-GO"),
                Value = formDataList.GOorNGO
            });
            newRow.Cells.Add(new Cell
            {
                ColumnId = GetColumnIdByName(sheet, "MISCode"),
                Value = formDataList.MISCode
            });

            smartsheet.SheetResources.RowResources.AddRows(parsedSheetId, new Row[] { newRow });




            return Ok(new
            { Message = " Success!" });

        }
        private long GetColumnIdByName(Sheet sheet, string columnname)
        {
            foreach (var column in sheet.Columns)
            {
                if (column.Title == columnname)
                {
                    return column.Id.Value;
                }
            }
            return 0;
        }
    }
    
}




