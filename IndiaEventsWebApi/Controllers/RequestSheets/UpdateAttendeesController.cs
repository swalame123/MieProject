//using Aspose.Pdf.Plugins;
//using IndiaEvents.Models.Models.EventTypeSheets;
//using IndiaEventsWebApi.Helper;
//using IndiaEventsWebApi.Models.EventTypeSheets;
//using Microsoft.AspNetCore.Http;
//using Microsoft.AspNetCore.Mvc;
//using Smartsheet.Api;
//using Smartsheet.Api.Models;

//namespace IndiaEventsWebApi.Controllers.RequestSheets
//{
//    [Route("api/[controller]")]
//    [ApiController]
//    public class UpdateAttendeesController : ControllerBase
//    {

//        private readonly string accessToken;
//        private readonly IConfiguration configuration;
//        private readonly SmartsheetClient smartsheet;

//        private readonly Sheet sheet1;
//        private readonly Sheet sheet2;
//        private readonly Sheet sheet3;
//        private readonly Sheet sheet4;

//        public UpdateAttendeesController(IConfiguration configuration)
//        {
//            this.configuration = configuration;
//            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

//            string panelSheet = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
//            string InviteeSheet = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
//            string ExpenseSheet = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
//            string SlideKitSheet = configuration.GetSection("SmartsheetSettings:EventRequestsHcpSlideKit").Value;

//            smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

//            sheet1 = SheetHelper.GetSheetById(smartsheet, panelSheet);
//            sheet2 = SheetHelper.GetSheetById(smartsheet, InviteeSheet);
//            sheet3 = SheetHelper.GetSheetById(smartsheet, ExpenseSheet);
//            sheet4 = SheetHelper.GetSheetById(smartsheet, SlideKitSheet);
//        }
//        [HttpPut("UpdateAttendees")]
//        public IActionResult UpdateAttendees(UpdateAttendees formData)
//        {
//            try
//            {
//                if (formData.InviteesData.Count > 0)
//                {
//                    foreach (var formdata in formData.InviteesData)
//                    {
//                        var targetRow = sheet2.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == formdata.InviteeId));
//                        if (targetRow != null)
//                        {
//                            Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };

//                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Attended?"), Value = formdata.IsAttended });
//                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Actual Local Conveyance Amount"), Value = formdata.ActualAmount });
//                            IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet2.Id.Value, new Row[] { updateRow });
//                            if (formdata.IsUploadDocument == "Yes")
//                            {
//                                foreach (var p in formdata.UploadDocument)
//                                {
//                                    string[] words = p.Split(':');
//                                    var r = words[0];
//                                    var q = words[1];
//                                    var name = r.Split(".")[0];
//                                    var filePath = SheetHelper.testingFile(q, formData.EventId, name);
//                                    var addedRow = updatedRow[0];
//                                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
//                                            sheet2.Id.Value, addedRow.Id.Value, filePath, "application/msword");
//                                    if (System.IO.File.Exists(filePath))
//                                    {
//                                        SheetHelper.DeleteFile(filePath);
//                                    }
//                                }

//                            }
//                        }
//                    }
//                }
//                if (formData.PanelData.Count > 0)
//                {
//                    foreach (var formdata in formData.PanelData)
//                    {
//                        var targetRow = sheet1.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == formdata.PanelId));
//                        if (targetRow != null)
//                        {

//                            Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
//                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Attended?"), Value = formdata.IsAttended });
//                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Actual Accomodation Amount"), Value = formdata.ActualAccomodationAmount });
//                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Actual Travel Amount"), Value = formdata.ActualTravelAmount });
//                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Actual Local Conveyance Amount"), Value = formdata.ActualLCAmount });

//                            IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRow });

//                        }
//                    }
//                }
//                if (formData.ExpenseData.Count > 0)
//                {
//                    foreach (var formdata in formData.ExpenseData)
//                    {
//                        var targetRow = sheet3.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == formdata.ExpenseId));
//                        if (targetRow != null)
//                        {
//                            Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
//                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Actual Amount"), Value = formdata.ActualAmount });

//                            IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet3.Id.Value, new Row[] { updateRow });
//                            if (formdata.IsUploadDocument == "Yes")
//                            {
//                                foreach (var p in formdata.UploadDocument)
//                                {

//                                    string[] words = p.Split(':');
//                                    var r = words[0];
//                                    var q = words[1];
//                                    var name = r.Split(".")[0];
//                                    var filePath = SheetHelper.testingFile(q, formData.EventId, name);
//                                    var addedRow = updatedRow[0];
//                                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
//                                            sheet3.Id.Value, addedRow.Id.Value, filePath, "application/msword");
//                                    if (System.IO.File.Exists(filePath))
//                                    {
//                                        SheetHelper.DeleteFile(filePath);
//                                    }
//                                }

//                            }
//                        }
//                    }
//                }
//                if (formData.SlideKitData.Count > 0)
//                {
//                    foreach (var formdata in formData.SlideKitData)
//                    {
//                        var targetRow = sheet4.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == formdata.SlideKitId));
//                        if (targetRow != null)
//                        {
//                            Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
//                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Actual Amount"), Value = formdata.ProductName });
//                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Actual Amount"), Value = formdata.IndicationsDone });
//                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Actual Amount"), Value = formdata.BatchNumber });
//                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Actual Amount"), Value = formdata.SubjectNameandSurName });

//                            IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet4.Id.Value, new Row[] { updateRow });
//                            if (formdata.IsUploadDocument == "Yes")
//                            {
//                                foreach (var p in formdata.UploadDocument)
//                                {

//                                    string[] words = p.Split(':');
//                                    var r = words[0];
//                                    var q = words[1];
//                                    var name = r.Split(".")[0];
//                                    var filePath = SheetHelper.testingFile(q, formData.EventId, name);
//                                    var addedRow = updatedRow[0];
//                                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
//                                            sheet4.Id.Value, addedRow.Id.Value, filePath, "application/msword");
//                                    if (System.IO.File.Exists(filePath))
//                                    {
//                                        SheetHelper.DeleteFile(filePath);
//                                    }
//                                }

//                            }
//                        }
//                    }
//                }

//                return Ok(new { Message = "Attendees Updated Successfully" });
//            }
//            catch (Exception ex)
//            {
//                return BadRequest(ex.Message);
//            }

//        }
//    }
//}

