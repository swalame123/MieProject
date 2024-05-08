using IndiaEventsWebApi.Helper;
using IndiaEventsWebApi.Models.RequestSheets;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using static Org.BouncyCastle.Bcpg.Attr.ImageAttrib;

namespace IndiaEventsWebApi.Controllers.RequestSheets
{
    [Route("api/[controller]")]
    [ApiController]
    public class AddorDeleteInviteesController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;

        public AddorDeleteInviteesController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

        }
        [HttpPost("AddNewInvitees")]
        public IActionResult AddNewInvitees(AddNewInvitee[] formDataList)
        {
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

            string sheetId3 = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
            long.TryParse(sheetId3, out long parsedSheetId3);

            Sheet sheet3 = smartsheet.SheetResources.GetSheet(parsedSheetId3, null, null, null, null, null, null, null);

            foreach (var formdata in formDataList)
            {
                var newRow3 = new Row();
                newRow3.Cells = new List<Cell>();
                newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "HCPName"), Value = formdata.InviteeName });
                newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Designation"), Value = formdata.Designation });
                //newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Addedinpostevent"), Value = formdata.Addedinpostevent });
                newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Employee Code"), Value = formdata.EmployeeCode });
                newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "LocalConveyance"), Value = formdata.LocalConveyance });
                newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "BTC/BTE"), Value = formdata.BtcorBte });
                newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "LcAmount"), Value = formdata.LcAmount });
                newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Lc Amount Excluding Tax"), Value = formdata.LcAmountExcludingTax });
                newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "EventId/EventRequestId"), Value = formdata.EventIdOrEventRequestId });
                newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Invitee Source"), Value = formdata.InviteedFrom });
                newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "HCP Type"), Value = formdata.HCPType });
                if (formdata.InviteedFrom == "Others")
                {
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Speciality"), Value = formdata.Speciality });
                }
                else
                {
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "MISCode"), Value = SheetHelper.MisCodeCheck(formdata.MISCode )});
                }
                newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Topic"), Value = formdata.EventTopic });
                newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Type"), Value = formdata.EventType });
                newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Venue name"), Value = formdata.VenueName });
                newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Date Start"), Value = formdata.EventStartDate });
                newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event End Date"), Value = formdata.EventEndDate });


                smartsheet.SheetResources.RowResources.AddRows(parsedSheetId3, new Row[] { newRow3 });


            }
            return Ok(new
            { Message = "Data added successfully." });
        }




        [HttpDelete("DeleteData/{RowInvId}")]
        public IActionResult DeleteData(string RowInvId)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
                long.TryParse(sheetId, out long parsedSheetId);

                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
                //Row existingRow = GetRowById(smartsheet, parsedSheetId, RowInvId);
                Row existingRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == RowInvId));


                if (existingRow == null)
                {
                    return NotFound($"Row with id {RowInvId} not found.");
                }
                var Id = (long)existingRow.Id;
                smartsheet.SheetResources.RowResources.DeleteRows(parsedSheetId, new long[] { Id }, true);

                return Ok(new { Message = "Data Deleted successfully." });
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }















        private Row GetRowById(SmartsheetClient smartsheet, long sheetId, string email)
        {
            Sheet sheet = smartsheet.SheetResources.GetSheet(sheetId, null, null, null, null, null, null, null);

            // Assuming you have a column named "Id"

            Column idColumn = sheet.Columns.FirstOrDefault(col => col.Title == "INV");

            if (idColumn != null)
            {
                foreach (var row in sheet.Rows)
                {
                    var cell = row.Cells.FirstOrDefault(c => c.ColumnId == idColumn.Id && c.Value.ToString() == email);

                    if (cell != null)
                    {
                        return row;
                    }
                }
            }

            return null;
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
