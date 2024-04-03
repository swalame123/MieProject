using Aspose.Pdf.Plugins;
using IndiaEvents.Models.Models.EventTypeSheets;
using IndiaEventsWebApi.Helper;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;

namespace IndiaEventsWebApi.Controllers.RequestSheets
{
    [Route("api/[controller]")]
    [ApiController]
    public class UpdateAttendeesController : ControllerBase
    {

        private readonly string accessToken;
        private readonly IConfiguration configuration;
        private readonly SmartsheetClient smartsheet;

        private readonly Sheet PanelSheet;
        private readonly Sheet InviteesSheet;

        public UpdateAttendeesController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

            string sheetId_SpeakerCode = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
            string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;

            smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

            InviteesSheet = SheetHelper.GetSheetById(smartsheet, sheetId);
            PanelSheet = SheetHelper.GetSheetById(smartsheet, sheetId_SpeakerCode);
        }
        [HttpPut("UpdateAttendees")]
        public IActionResult UpdateAttendees(UpdateAttendees formData)
        {
            try
            {
                if (formData.InviteesAttendees.Count > 0)
                {
                    foreach (var id in formData.InviteesAttendees)
                    {
                        var targetRow = InviteesSheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == id));
                        if (targetRow != null)
                        {
                            long AttendedColumnId = SheetHelper.GetColumnIdByName(InviteesSheet, "Attended?");
                            var cellToUpdateB = new Cell { ColumnId = AttendedColumnId, Value = "Yes" };
                            Row updateRow = new Row { Id = targetRow.Id, Cells = new Cell[] { cellToUpdateB } };
                            var cellToUpdate = targetRow.Cells.FirstOrDefault(c => c.ColumnId == AttendedColumnId);
                            if (cellToUpdate != null) { cellToUpdate.Value = "Yes"; }

                            smartsheet.SheetResources.RowResources.UpdateRows(InviteesSheet.Id.Value, new Row[] { updateRow });
                        }
                    }
                }
                if (formData.PanelAttendees.Count > 0)
                {
                    foreach (var id in formData.PanelAttendees)
                    {
                        var targetRow = PanelSheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == id));
                        if (targetRow != null)
                        {
                            long AttendedColumnId = SheetHelper.GetColumnIdByName(PanelSheet, "Attended?");
                            var cellToUpdateB = new Cell { ColumnId = AttendedColumnId, Value = "Yes" };
                            Row updateRow = new Row { Id = targetRow.Id, Cells = new Cell[] { cellToUpdateB } };
                            var cellToUpdate = targetRow.Cells.FirstOrDefault(c => c.ColumnId == AttendedColumnId);
                            if (cellToUpdate != null) { cellToUpdate.Value = "Yes"; }

                            smartsheet.SheetResources.RowResources.UpdateRows(PanelSheet.Id.Value, new Row[] { updateRow });
                        }
                    }
                }
                return Ok(new { Message = "Attendees Updated Successfully" });
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }

        }
    }
}

