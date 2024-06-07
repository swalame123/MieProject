using IndiaEvents.Models.Models.EventTypeSheets;
using IndiaEventsWebApi.Helper;
using IndiaEventsWebApi.Models.EventTypeSheets;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Serilog;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Globalization;
using System.Text;

namespace IndiaEventsWebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class HonorariumController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;
        //private readonly SmartsheetClient smartsheet;
        //private readonly Sheet sheet1;
        //private readonly Sheet sheet2;
        //private readonly Sheet sheet3;
        //private readonly Sheet sheet4;
        //private readonly Sheet sheet5;
        //private readonly Sheet sheet6;
        //private readonly Sheet sheet7;
        private readonly SemaphoreSlim _externalApiSemaphore;
        public HonorariumController(IConfiguration configuration, SemaphoreSlim externalApiSemaphore)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
            _externalApiSemaphore = externalApiSemaphore;
            //smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
        }

        [HttpPost("AddHonorariumData")]
        public async Task<IActionResult> AddHonorariumData(HonorariumPaymentListPh2 formData)
        {
            try
            {
                SmartsheetClient smartsheet = await Task.Run(() => SmartSheetBuilder.AccessClient(accessToken, _externalApiSemaphore));

                string sheetId = configuration.GetSection("SmartsheetSettings:HonorariumPayment").Value;
                string sheetId1 = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;
                string sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;
                string UI_URL = configuration.GetSection("SmartsheetSettings:UI_URL").Value;

                Sheet UrlData = SheetHelper.GetSheetById(smartsheet, UI_URL);

                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
                Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);
                

                StringBuilder addedHcpData = new();
                int addedHcpDataNo = 1;

                foreach (var i in formData.HcpRoles)
                {
                    string rowData = $"{addedHcpDataNo}. Name:{i.HcpName} | Role:{i.HcpRole} |Code:{i.MisCode} | HCP Type:{i.GOorNGO}| Including GST:{i.IsInclidingGst}| Agreement Amount:{i.AgreementAmount} |Annual Trainer Agreement Valid: {i.IsAnnualTrainerAgreementValid} ";
                    addedHcpData.AppendLine(rowData);
                    addedHcpDataNo++;
                }
                string panalist = addedHcpData.ToString();
                Dictionary<string, long> Sheetcolumns = new();
                foreach (var column in sheet.Columns)
                {
                    Sheetcolumns.Add(column.Title, (long)column.Id);
                }
                Row newRow = new()
                {
                    Cells = new List<Cell>()
                };

                Row? targetRow1 = UrlData.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == "Finance Accounts URL"));
                Row? targetRow2 = UrlData.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == "Finance Treasury URL"));
                Row? targetRow3 = UrlData.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == "Approver Honorarium URL"));
                Row? targetRow4 = UrlData.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == "Initiator URL"));


                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Finance Accounts URL"], Value = targetRow1?.Cells[1].Value ?? "no url" });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Finance Treasury URL"], Value = targetRow2?.Cells[1].Value ?? "no url" });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Approver Honorarium URL"], Value = targetRow3?.Cells[1].Value ?? "no url" });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Initiator URL"], Value = targetRow4?.Cells[1].Value ?? "no url" });

                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["SlideKits"], Value = formData.RequestHonorariumList.SlideKits });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["EventId/EventRequestId"], Value = formData.RequestHonorariumList.EventId });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Event Type"], Value = formData.RequestHonorariumList.EventType });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Event Date"], Value = formData.RequestHonorariumList.EventDate });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Event Topic"], Value = formData.RequestHonorariumList.EventTopic });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["City"], Value = formData.RequestHonorariumList.City });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["State"], Value = formData.RequestHonorariumList.State });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Start Time"], Value = formData.RequestHonorariumList.StartTime });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["End Time"], Value = formData.RequestHonorariumList.EndTime });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Venue Name"], Value = formData.RequestHonorariumList.VenueName });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Total Travel & Accommodation Amount"], Value = formData.RequestHonorariumList.TotalTravelAndAccomodationSpend });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Total Honorarium Amount"], Value = formData.RequestHonorariumList.TotalHonorariumSpend });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Budget Amount"], Value = formData.RequestHonorariumList.TotalSpend });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Expenses"], Value = formData.RequestHonorariumList.Expenses });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Total Travel Amount"], Value = formData.RequestHonorariumList.TotalTravelSpend });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Total Accommodation Amount"], Value = formData.RequestHonorariumList.TotalAccomodationSpend });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Total Expense"], Value = formData.RequestHonorariumList.TotalExpenses });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Total Local Conveyance"], Value = formData.RequestHonorariumList.TotalLocalConveyance });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Brands"], Value = formData.RequestHonorariumList.Brands });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Invitees"], Value = formData.RequestHonorariumList.Invitees });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Panelists"], Value = formData.RequestHonorariumList.Panelists });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Initiator Name"], Value = formData.RequestHonorariumList.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Panelists & Agreements"], Value = panalist });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Honorarium Submitted?"], Value = formData.RequestHonorariumList.HonarariumSubmitted });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Initiator Email"], Value = formData.RequestHonorariumList.InitiatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["RBM/BM"], Value = formData.RequestHonorariumList.RBMorBM });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Sales Head"], Value = formData.RequestHonorariumList.SalesHeadEmail });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Sales Coordinator"], Value = formData.RequestHonorariumList.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Marketing Coordinator"], Value = formData.RequestHonorariumList.MarketingCoordinatorEmail }); newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Marketing Head"), Value = formData.RequestHonorariumList.MarketingHeadEmail });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Compliance"], Value = formData.RequestHonorariumList.Compliance });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Finance Accounts"], Value = formData.RequestHonorariumList.FinanceAccounts });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Finance Treasury"], Value = formData.RequestHonorariumList.FinanceTreasury });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Reporting Manager"], Value = formData.RequestHonorariumList.ReportingManagerEmail });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["1 Up Manager"], Value = formData.RequestHonorariumList.FirstLevelEmail });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Medical Affairs Head"], Value = formData.RequestHonorariumList.MedicalAffairsEmail });
                newRow.Cells.Add(new Cell { ColumnId = Sheetcolumns["Role"], Value = formData.RequestHonorariumList.Role });

                //IList<Row> addedRows = smartsheet.SheetResources.RowResources.AddRows(sheet.Id.Value, new Row[] { newRow });


                IList<Row> addedRows = ApiCalls.HonorariumDetails(smartsheet, sheet, newRow);



                string? eventId = formData.RequestHonorariumList.EventId;

                foreach (string p in formData.RequestHonorariumList.Files)
                {
                    string[] words = p.Split(':');
                    string r = words[0];
                    string q = words[1];
                    string name = r.Split(".")[0];
                    string filePath = SheetHelper.testingFile(q, name);
                    Row addedRow = addedRows[0];
                    //Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                    //       sheet.Id.Value, addedRow.Id.Value, filePath, "application/msword");

                    Attachment attachment = await ApiCalls.AddAttachmentsToSheet(smartsheet, sheet, addedRow, filePath);




                    if (System.IO.File.Exists(filePath))
                    {
                        SheetHelper.DeleteFile(filePath);
                    }
                }

                if (formData.RequestHonorariumList.IsDeviationUpload == "Yes")
                {
                    Dictionary<string, long> Sheet7columns = new();
                    foreach (var column in sheet7.Columns)
                    {
                        Sheet7columns.Add(column.Title, (long)column.Id);
                    }
                    try
                    {
                        Row newRow7 = new()
                        {
                            Cells = new List<Cell>()
                        };
                        newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["EventId/EventRequestId"], Value = eventId });
                        newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Event Topic"], Value = formData.RequestHonorariumList.EventTopic });
                        newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Event Type"], Value = formData.RequestHonorariumList.EventType });
                        newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Event Date"], Value = formData.RequestHonorariumList.EventDate });
                        newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Start Time"], Value = formData.RequestHonorariumList.StartTime });
                        newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["End Time"], Value = formData.RequestHonorariumList.EndTime });
                        newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Venue Name"], Value = formData.RequestHonorariumList.VenueName });
                        newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["City"], Value = formData.RequestHonorariumList.City });
                        newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["State"], Value = formData.RequestHonorariumList.State });
                        newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Deviation Type"], Value = configuration.GetSection("DeviationNamesInHonorarium:5WorkingdaysDeviationDateTrigger").Value });
                        newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["HON-5Workingdays Deviation Date Trigger"], Value = formData.RequestHonorariumList.IsDeviationUpload });
                        newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Sales Head"], Value = formData.RequestHonorariumList.SalesHeadEmail });
                        newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Finance Head"], Value = formData.RequestHonorariumList.FinanceHead });
                        newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Initiator Name"], Value = formData.RequestHonorariumList.InitiatorName });
                        newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Initiator Email"], Value = formData.RequestHonorariumList.InitiatorEmail });
                        newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Sales Coordinator"], Value = formData.RequestHonorariumList.SalesCoordinatorEmail });

                        // IList<Row> addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });
                        IList<Row> addeddeviationrow = ApiCalls.DeviationData(smartsheet, sheet7, newRow7);

                        foreach (string p in formData.RequestHonorariumList.DeviationFiles)
                        {
                            string[] words = p.Split(':');
                            string r = words[0];
                            string q = words[1];
                            string name = r.Split("*")[0];
                            string filePath = SheetHelper.testingFile(q, name);
                            Row addedRow = addeddeviationrow[0];
                            Row addedRowInHonr = addedRows[0];

                            Attachment attachment = await ApiCalls.AddAttachmentsToSheet(smartsheet, sheet7, addedRow, filePath);
                            Attachment Deviationattachment = await ApiCalls.AddAttachmentsToSheet(smartsheet, sheet, addedRows[0], filePath);


                            //Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                            //Attachment Deviationattachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet.Id.Value, addedRowInHonr.Id.Value, filePath, "application/msword");


                            if (System.IO.File.Exists(filePath))
                            {
                                SheetHelper.DeleteFile(filePath);
                            }
                        }
                        //}
                        //catch (Exception ex)
                        //{
                        //    Log.Error($"Error occured on HonorariumController method {ex.Message} at {DateTime.Now}");
                        //    Log.Error(ex.StackTrace);
                        //    return BadRequest(ex.Message);
                        //}
                    }
                    catch (Exception ex)
                    {

                        return BadRequest(new
                        {
                            Message = ex.Message + "------" + ex.StackTrace
                        });
                    }
                }

                Row? targetRow = sheet1.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == eventId));
                if (targetRow != null)
                {
                    long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(sheet1, "Honorarium Submitted?");
                    Cell cellToUpdateB = new Cell
                    { ColumnId = honorariumSubmittedColumnId, Value = "Yes" };
                    Row updateRow = new Row { Id = targetRow.Id, Cells = new Cell[] { cellToUpdateB } };
                    Cell? cellToUpdate = targetRow.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                    if (cellToUpdate != null) { cellToUpdate.Value = "Yes"; }

                    await Task.Run(() => ApiCalls.UpdateRole(smartsheet, sheet1, updateRow));

                    //smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRow });
                }
                return Ok(new
                { Message = "Data added successfully." });
                //}
                //catch (Exception ex)
                //{
                //    Log.Error($"Error occured on HonorariumController method {ex.Message} at {DateTime.Now}");
                //    Log.Error(ex.StackTrace);
                //    return BadRequest(ex.Message);
                //}
            }
            catch (Exception ex)
            {

                return BadRequest(new
                {
                    Message = ex.Message + "------" + ex.StackTrace
                });
            }
        }


        [HttpPut("HonorariumUpdate")]
        public async Task<IActionResult> HonorariumUpdate(List<HonorariumUpdate> formDataArray)
        {
            try
            {
                SmartsheetClient smartsheet = await Task.Run(() => SmartSheetBuilder.AccessClient(accessToken, _externalApiSemaphore));

                string sheetId4 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;

                Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
                foreach (var formdata in formDataArray)
                {
                    var targetRow = sheet4.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == formdata.PanelId));
                    if (targetRow != null)
                    {
                        Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Annual Trainer Agreement Valid?"), Value = formdata.IsAnnualTrainerAgreementValid });
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Including GST?"), Value = formdata.IsInclidingGst });

                        IList<Row> updatedRow = await Task.Run(() => ApiCalls.UpdateRole(smartsheet, sheet4, updateRow));
                        //smartsheet.SheetResources.RowResources.UpdateRows(sheet4.Id.Value, new Row[] { updateRow });

                        if (formdata.FilesToUpload.Count > 0)
                        {
                            // PaginatedResult<Attachment> attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet4.Id.Value, targetRow.Id.Value, null);

                            PaginatedResult<Attachment> attachments = await Task.Run(() => ApiCalls.GetAttachmantsFromSheet(smartsheet, sheet4, targetRow));


                            if (attachments.Data != null || attachments.Data.Count > 0)
                            {

                                foreach (var attachment in attachments.Data)
                                {
                                    var name = attachment.Name;
                                    if (name.ToLower().Contains("speakeragreement") || name.ToLower().Contains("honorariumInvoice"))
                                    {
                                        long Id = attachment.Id.Value;
                                        smartsheet.SheetResources.AttachmentResources.DeleteAttachment(sheet4.Id.Value, Id);
                                    }
                                }

                            }
                        }
                        foreach (var p in formdata.FilesToUpload)
                        {
                            string[] words = p.Split(':');
                            string r = words[0];
                            string q = words[1];
                            string val = formdata.EventId;
                            string name = r.Split(".")[0];
                            string filePath = SheetHelper.testingFile(q, name);
                            Row addedRow = targetRow;
                            //Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet4.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                            Attachment attachment = await ApiCalls.AddAttachmentsToSheet(smartsheet, sheet4, addedRow, filePath);

                            if (System.IO.File.Exists(filePath))
                            {
                                SheetHelper.DeleteFile(filePath);
                            }

                        }
                    }
                }
                return Ok(new { Message = " Updated Successfully" });
                //}
                //catch (Exception ex)
                //{
                //    Log.Error($"Error occured on HonorariumController method {ex.Message} at {DateTime.Now}");
                //    Log.Error(ex.StackTrace);
                //    return BadRequest(ex.Message);
                //}
            }
            catch (Exception ex)
            {

                return BadRequest(new
                {
                    Message = ex.Message + "------" + ex.StackTrace
                });
            }

        }

    }
}
