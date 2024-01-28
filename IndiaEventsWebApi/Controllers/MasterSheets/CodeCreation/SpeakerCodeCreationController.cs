using IndiaEventsWebApi.Models.MasterSheets.CodeCreation;
using IndiaEventsWebApi.Models.RequestSheets;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;

namespace IndiaEventsWebApi.Controllers.MasterSheets.CodeCreation
{
    [Route("api/[controller]")]
    [ApiController]
    public class SpeakerCodeCreationController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;

        public SpeakerCodeCreationController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

        }

        [HttpPost("AddSpeakersData")]
        public IActionResult AddSpeakersData(SpeakerCodeGeneration[] formData)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

                string sheetId = configuration.GetSection("SmartsheetSettings:SpeakerCodeCreation").Value;


                long.TryParse(sheetId, out long parsedSheetId);

                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);

                foreach (var i in formData)
                {
                    //var newRow = new Row();
                    //newRow.Cells = new List<Cell>();
                    //newRow.Cells.Add(new Cell
                    //{
                    //    ColumnId = GetColumnIdByName(sheet, "HCP Name"),
                    //    Value = i.EventId
                    //});

                    //newRow.Cells.Add(new Cell
                    //{
                    //    ColumnId = GetColumnIdByName(sheet, "EventId/EventRequestId"),
                    //    Value = i.Expense
                    //});
                    //newRow.Cells.Add(new Cell
                    //{
                    //    ColumnId = GetColumnIdByName(sheet, "EventType"),
                    //    Value = i.Amount
                    //});
                    //newRow.Cells.Add(new Cell
                    //{
                    //    ColumnId = GetColumnIdByName(sheet, "HCPRole"),
                    //    Value = i.AmountExcludingTax
                    //});
                    //newRow.Cells.Add(new Cell
                    //{
                    //    ColumnId = GetColumnIdByName(sheet, "MISCODE"),
                    //    Value = i.BtcorBte
                    //});
                    //newRow.Cells.Add(new Cell
                    //{
                    //    ColumnId = GetColumnIdByName(sheet, "GO/Non-GO"),
                    //    Value = i.BtcAmount
                    //});
                    //newRow.Cells.Add(new Cell
                    //{
                    //    ColumnId = GetColumnIdByName(sheet, "IsItincludingGST?"),
                    //    Value = i.BteAmount
                    //});
                    //newRow.Cells.Add(new Cell
                    //{
                    //    ColumnId = GetColumnIdByName(sheet, "AgreementAmount"),
                    //    Value = i.BudgetAmount
                    //});



                    //var addedRows = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId, new Row[] { newRow });
                    //var eventId = i.EventId;






                }



                return Ok(new
                { Message = "Data added successfully." });

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
