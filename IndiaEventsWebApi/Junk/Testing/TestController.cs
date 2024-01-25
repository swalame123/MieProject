using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using IndiaEventsWebApi.Models;
using IndiaEventsWebApi.Models.RequestSheets;

using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Text;

namespace IndiaEventsWebApi.Junk.Testing
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
        [HttpGet("GetEventRequestsHcpRoleById/{eventIdorEventRequestId}")]
        public IActionResult GetEventRequestsHcpRoleById(string eventIdorEventRequestId)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
                long.TryParse(sheetId, out long parsedSheetId);
                Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
                List<Dictionary<string, object>> hcpRoleData = new List<Dictionary<string, object>>();
                List<string> columnNames = new List<string>();

                foreach (Column column in sheet1.Columns)
                {
                    columnNames.Add(column.Title);
                }
                foreach (Row row in sheet1.Rows)
                {

                    // Check if the row has the specified EventIdorEventRequestId
                    var eventIdorEventRequestIdCell = row.Cells.FirstOrDefault(cell => cell.ColumnId == GetColumnIdByName(sheet1, "EventId/EventRequestId"));
                    var x = eventIdorEventRequestIdCell.Value.ToString();
                    // Console.WriteLine($"Row {row.RowNumber}: EventId/EventRequestId in cell: {eventIdorEventRequestIdCell.Value}");
                    if (eventIdorEventRequestIdCell != null && x == eventIdorEventRequestId)
                    {
                        Dictionary<string, object> hcpRoleRowData = new Dictionary<string, object>();

                        for (int i = 0; i < row.Cells.Count && i < columnNames.Count; i++)
                        {
                            hcpRoleRowData[columnNames[i]] = row.Cells[i].Value;
                        }

                        hcpRoleData.Add(hcpRoleRowData);
                    }

                }

                return Ok(hcpRoleData);

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
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
