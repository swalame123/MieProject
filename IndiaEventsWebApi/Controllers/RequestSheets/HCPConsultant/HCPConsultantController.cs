using IndiaEventsWebApi.Helper;
using IndiaEventsWebApi.Models;
using IndiaEventsWebApi.Models.EventTypeSheets;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Globalization;
using System.Text;

namespace IndiaEventsWebApi.Controllers.RequestSheets.HCPConsultant
{
    [Route("api/[controller]")]
    [ApiController]
    //[Authorize]
    public class HCPConsultantController : ControllerBase
    {

        private readonly string accessToken;
        private readonly IConfiguration configuration;

        public HCPConsultantController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

        }


        [HttpGet("HcpFollowUpData")]
        public IActionResult GetEventRequestWebData()
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:HcpFollowup").Value;
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


        [HttpPost("HcpFollowup"), DisableRequestSizeLimit]
        public IActionResult HcpFollowup(List<HCPfollow_upsheet> formDataList)
        {
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            string sheetId1 = configuration.GetSection("SmartsheetSettings:HcpFollowup").Value;
            long.TryParse(sheetId1, out long parsedSheetId1);
            Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId1, null, null, null, null, null, null, null);

            foreach (var formdata in formDataList)
            {
                try
                {

                    var newRow = new Row();
                    newRow.Cells = new List<Cell>();
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "HCP Name"), Value = formdata.HCPName });

                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "MIS Code"), Value = SheetHelper.MisCodeCheck(formdata.MisCode )});
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "GO/N-GO"), Value = formdata.GO_NGO });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Country"), Value = formdata.Country });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "How many days since the parent event Completes"), Value = formdata.How_many_days_since_the_parent_event_completes });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Follow-up Event Date"), Value = formdata.Follow_up_Event_Date });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Date"), Value = formdata.Event_Date });

                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Follow-up Event Code"), Value = formdata.Follow_up_Event });

                    var addedRows = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId1, new Row[] { newRow });
                    var columnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                    var Cell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == columnId);
                    var value = Cell.DisplayValue;
                    if (formdata.AgreementFile != "")
                    {

                        var filename = "AgreementFile";
                        var filePath = SheetHelper.testingFile(formdata.AgreementFile, filename);




                        var addedRow = addedRows[0];


                        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                parsedSheetId1, addedRow.Id.Value, filePath, "application/msword");

                    }



                }
                catch (Exception ex)
                {
                    return BadRequest(ex.Message);
                }
            }
            return Ok(new { Message = " success!" });
        }



       

    }
}
