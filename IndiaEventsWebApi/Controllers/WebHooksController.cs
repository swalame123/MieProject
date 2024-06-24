using IndiaEventsWebApi.Helper;
using IndiaEventsWebApi.Models.RequestSheets;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Serilog;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Data;
using System.Runtime.CompilerServices;
using System.Text;
using static IndiaEvents.Models.Models.Webhook.WebhookModels;
using IndiaEvents.Models.Models.Webhook;
using iTextSharp.text.pdf;
using iTextSharp.text;
using static iTextSharp.text.pdf.AcroFields;
using Microsoft.Extensions.Logging;
using IndiaEventsWebApi.Models.EventTypeSheets;
using Microsoft.Extensions.Hosting;
using System.Globalization;
using Aspose.Cells.Rendering;
using IndiaEvents.Models.Models.Draft;
using Microsoft.CodeAnalysis.Elfie.Serialization;

namespace IndiaEventsWebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class WebHooksController : ControllerBase
    {

        private readonly string accessToken;
        private readonly IConfiguration configuration;
        private readonly SmartsheetClient smartsheet;
        private readonly string processSheet;
        private readonly string sheetId_SpeakerCode;
        private readonly string sheetId;
        private readonly string sheetId1;
        public WebHooksController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessTokenForWebHooks").Value;
            smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

            processSheet = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;
            sheetId_SpeakerCode = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
            sheetId = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
            sheetId1 = configuration.GetSection("SmartsheetSettings:Class1").Value;


        }

        [HttpPost("WebHookData")]
        public async Task<IActionResult> WebHookData()
        {
            try
            {
                Dictionary<string, string> requestHeaders = new Dictionary<string, string>();
                string rawContent = string.Empty;
                using (var reader = new StreamReader(Request.Body, encoding: Encoding.UTF8, detectEncodingFromByteOrderMarks: false))
                {
                    rawContent = await reader.ReadToEndAsync();
                }
                requestHeaders.Add("Body", rawContent);
                Log.Information(string.Join(";", requestHeaders.Select(x => x.Key + "=" + x.Value).ToArray()));


                Root? RequestWebhook = JsonConvert.DeserializeObject<Root>(rawContent);
                Attachementfile(RequestWebhook);

                string? challenge = requestHeaders.Where(x => x.Key == "challenge").Select(x => x.Value).FirstOrDefault();

                return Ok(new Webhook { smartsheetHookResponse = RequestWebhook.challenge });
                //return Ok();
            }
            catch (Exception ex)
            {
                Log.Error($"Error occured on Webhook apicontroller PostData method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(ex.StackTrace);
            }
        }

        [HttpPost("EmailWebHook")]
        public async Task<IActionResult> EmailWebHook()
        {
            try
            {
                Dictionary<string, string> requestHeaders = new Dictionary<string, string>();
                string rawContent = string.Empty;
                using (var reader = new StreamReader(Request.Body, encoding: Encoding.UTF8, detectEncodingFromByteOrderMarks: false))
                {
                    rawContent = await reader.ReadToEndAsync();
                }
                requestHeaders.Add("Body", rawContent);
                Log.Information(string.Join(";", requestHeaders.Select(x => x.Key + "=" + x.Value).ToArray()));


                Root? RequestWebhook = JsonConvert.DeserializeObject<Root>(rawContent);
                MailChange(RequestWebhook);

                string? challenge = requestHeaders.Where(x => x.Key == "challenge").Select(x => x.Value).FirstOrDefault();

                return Ok(new Webhook { smartsheetHookResponse = RequestWebhook.challenge });
                //return Ok();
            }
            catch (Exception ex)
            {
                Log.Error($"Error occured on Webhook mailchange  method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(ex.StackTrace);
            }
        }

        [HttpPost("EmailWebHookForEmployeeMaster")]
        public async Task<IActionResult> EmailWebHookForEmployeeMaster()
        {
            try
            {
                Dictionary<string, string> requestHeaders = new Dictionary<string, string>();
                string rawContent = string.Empty;
                using (var reader = new StreamReader(Request.Body, encoding: Encoding.UTF8, detectEncodingFromByteOrderMarks: false))
                {
                    rawContent = await reader.ReadToEndAsync();
                }
                requestHeaders.Add("Body", rawContent);
                Log.Information(string.Join(";", requestHeaders.Select(x => x.Key + "=" + x.Value).ToArray()));


                Root? RequestWebhook = JsonConvert.DeserializeObject<Root>(rawContent);
                MailChangeInMasters(RequestWebhook);

                string? challenge = requestHeaders.Where(x => x.Key == "challenge").Select(x => x.Value).FirstOrDefault();

                return Ok(new Webhook { smartsheetHookResponse = RequestWebhook.challenge });
                //return Ok();
            }
            catch (Exception ex)
            {
                Log.Error($"Error occured on Webhook mailchange  method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(ex.StackTrace);
            }
        }

        [HttpPost("WebHookForAgreements")]
        public async Task<IActionResult> WebHookPostmethod()
        {
            try
            {
                Dictionary<string, string> requestHeaders = new Dictionary<string, string>();
                string rawContent = string.Empty;
                using (var reader = new StreamReader(Request.Body, encoding: Encoding.UTF8, detectEncodingFromByteOrderMarks: false))
                {
                    rawContent = await reader.ReadToEndAsync();
                }
                requestHeaders.Add("Body", rawContent);



                Root? RequestWebhook = JsonConvert.DeserializeObject<Root>(rawContent);
                AgreementsTrigger(RequestWebhook);

                var challenge = requestHeaders.Where(x => x.Key == "challenge").Select(x => x.Value).FirstOrDefault();

                return Ok(new Webhook { smartsheetHookResponse = RequestWebhook.challenge });
                //return Ok();
            }
            catch (Exception ex)
            {
                Log.Error($"Error occured on Webhook apicontroller PostData method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(ex.StackTrace);
            }
        }

        [HttpPost("WebHookForHonorariumApprovals")]
        public async Task<IActionResult> WebHookForHonorariumApprovals()
        {
            try
            {
                Dictionary<string, string> requestHeaders = new Dictionary<string, string>();
                string rawContent = string.Empty;
                using (var reader = new StreamReader(Request.Body, encoding: Encoding.UTF8, detectEncodingFromByteOrderMarks: false))
                {
                    rawContent = await reader.ReadToEndAsync();
                }
                requestHeaders.Add("Body", rawContent);


                Root? RequestWebhook = JsonConvert.DeserializeObject<Root>(rawContent);
                ApprovalCheckBox(RequestWebhook);

                string? challenge = requestHeaders.Where(x => x.Key == "challenge").Select(x => x.Value).FirstOrDefault();

                return Ok(new Webhook { smartsheetHookResponse = RequestWebhook.challenge });
                //return Ok();
            }
            catch (Exception ex)
            {
                Log.Error($"Error occured on Webhook apicontroller PostData method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(ex.StackTrace);
            }
        }

        [HttpPost("WebHookForEventSettlement")]
        public async Task<IActionResult> WebHookForEventSettlement()
        {
            try
            {
                Dictionary<string, string> requestHeaders = new Dictionary<string, string>();
                string rawContent = string.Empty;
                using (StreamReader reader = new StreamReader(Request.Body, encoding: Encoding.UTF8, detectEncodingFromByteOrderMarks: false))
                {
                    rawContent = await reader.ReadToEndAsync();
                }
                requestHeaders.Add("Body", rawContent);


                Root? RequestWebhook = JsonConvert.DeserializeObject<Root>(rawContent);
                //EventSettlementApproval(RequestWebhook);
                EventSettlementDeviationApproval(RequestWebhook);
                string? challenge = requestHeaders.Where(x => x.Key == "challenge").Select(x => x.Value).FirstOrDefault();

                return Ok(new Webhook { smartsheetHookResponse = RequestWebhook.challenge });
                //return Ok();
            }
            catch (Exception ex)
            {
                Log.Error($"Error occured on Webhook apicontroller PostData method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(ex.StackTrace);
            }
        }

        //[HttpPost("WebHookForEventSettlementDeviationApprovalCheck")]
        //public async Task<IActionResult> WebHookForEventSettlementDeviationApprovalCheck()
        //{
        //    try
        //    {
        //        Dictionary<string, string> requestHeaders = new Dictionary<string, string>();
        //        string rawContent = string.Empty;
        //        using (var reader = new StreamReader(Request.Body, encoding: Encoding.UTF8, detectEncodingFromByteOrderMarks: false))
        //        {
        //            rawContent = await reader.ReadToEndAsync();
        //        }
        //        requestHeaders.Add("Body", rawContent);


        //        var RequestWebhook = JsonConvert.DeserializeObject<Root>(rawContent);
        //        EventSettlementDeviationApproval(RequestWebhook);

        //        var challenge = requestHeaders.Where(x => x.Key == "challenge").Select(x => x.Value).FirstOrDefault();

        //        return Ok(new Webhook { smartsheetHookResponse = RequestWebhook.challenge });
        //        //return Ok();
        //    }
        //    catch (Exception ex)
        //    {
        //        Log.Error($"Error occured on Webhook apicontroller PostData method {ex.Message} at {DateTime.Now}");
        //        Log.Error(ex.StackTrace);
        //        return BadRequest(ex.StackTrace);
        //    }
        //}

        [HttpPost("WebHookForPreEventApproval")]
        public async Task<IActionResult> WebHookForPreEventApproval()
        {
            try
            {
                Dictionary<string, string> requestHeaders = [];
                string rawContent = string.Empty;
                using (var reader = new StreamReader(Request.Body, encoding: Encoding.UTF8, detectEncodingFromByteOrderMarks: false))
                {
                    rawContent = await reader.ReadToEndAsync();
                }
                requestHeaders.Add("Body", rawContent);


                Root? RequestWebhook = JsonConvert.DeserializeObject<Root>(rawContent);
                PreEventApproval(RequestWebhook);

                string? challenge = requestHeaders.Where(x => x.Key == "challenge").Select(x => x.Value).FirstOrDefault();

                return Ok(new Webhook { smartsheetHookResponse = RequestWebhook.challenge });
                //return Ok();
            }
            catch (Exception ex)
            {
                Log.Error($"Error occured on Webhook apicontroller PostData method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(ex.StackTrace);
            }
        }
        private async void PreEventApproval(Root RequestWebhook)
        {

            try
            {
                string processSheet = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;

                Sheet TestingSheetData = SheetHelper.GetSheetById(smartsheet, processSheet);

                //Column? Trigger = TestingSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "Deviation Status", StringComparison.OrdinalIgnoreCase));
                //Column Updatecolumn = TestingSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "Is All Deviations Approved?", StringComparison.OrdinalIgnoreCase));
                Column? Column1 = TestingSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "EventOpenSalesHeadApproval", StringComparison.OrdinalIgnoreCase));
                Column? Column2 = TestingSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "7daysSalesHeadApproval", StringComparison.OrdinalIgnoreCase));
                Column? Column3 = TestingSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "PRE-F&B Expense Excluding Tax Approval", StringComparison.OrdinalIgnoreCase));
                Column? Column4 = TestingSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "HCP exceeds 1,00,000 FH Approval", StringComparison.OrdinalIgnoreCase));
                Column? Column5 = TestingSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "HCP exceeds 5,00,000 Trigger FH Approval", StringComparison.OrdinalIgnoreCase));
                Column? Column6 = TestingSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "HCP Honorarium 6,00,000 Exceeded Approval", StringComparison.OrdinalIgnoreCase));
                Column? Column7 = TestingSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "Trainer Honorarium 12,00,000 Exceeded Approval", StringComparison.OrdinalIgnoreCase));
                Column? Column8 = TestingSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "Travel/Accomodation 3,00,000 Exceeded Approval", StringComparison.OrdinalIgnoreCase));
                Column? Column9 = TestingSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "Vendor Status", StringComparison.OrdinalIgnoreCase));

                if (RequestWebhook != null && RequestWebhook.events != null)
                {
                    foreach (var WebHookEvent in RequestWebhook.events)
                    {
                        if (WebHookEvent.eventType.ToLower() == "updated" || WebHookEvent.eventType.ToLower() == "created")
                        {
                            Row targetRowId = TestingSheetData.Rows.FirstOrDefault(row => row.Id == WebHookEvent.rowId);


                            if (targetRowId != null)
                            {
                                //string? TriggerStatus = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == Trigger.Id)?.Value.ToString();
                                string? status1 = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == Column1.Id)?.Value?.ToString() ?? "Null";
                                string? status2 = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == Column2.Id)?.Value?.ToString() ?? "Null";
                                string? status3 = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == Column3.Id)?.Value?.ToString() ?? "Null";
                                string? status4 = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == Column4.Id)?.Value?.ToString() ?? "Null";
                                string? status5 = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == Column5.Id)?.Value?.ToString() ?? "Null";
                                string? status6 = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == Column6.Id)?.Value?.ToString() ?? "Null";
                                string? status7 = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == Column7.Id)?.Value?.ToString() ?? "Null";
                                string? status8 = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == Column8.Id)?.Value?.ToString() ?? "Null";
                                string? status9 = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == Column9.Id)?.Value?.ToString() ?? "Null";

                                if (status1.ToLower() == "approved")
                                {
                                    long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(TestingSheetData, "pre-45 days Approval");
                                    Cell cellToUpdateB = new() { ColumnId = honorariumSubmittedColumnId, Value = "Yes" };
                                    Row updateRow = new() { Id = targetRowId.Id, Cells = new Cell[] { cellToUpdateB } };
                                    Cell? cellToUpdate = targetRowId.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                                    if (cellToUpdate != null) { cellToUpdate.Value = "Yes"; }

                                    smartsheet.SheetResources.RowResources.UpdateRows(TestingSheetData.Id.Value, new Row[] { updateRow });
                                }
                                if (status2.ToLower() == "approved")
                                {
                                    long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(TestingSheetData, "pre 5 days approval");
                                    Cell cellToUpdateB = new() { ColumnId = honorariumSubmittedColumnId, Value = "Yes" };
                                    Row updateRow = new() { Id = targetRowId.Id, Cells = new Cell[] { cellToUpdateB } };
                                    Cell? cellToUpdate = targetRowId.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                                    if (cellToUpdate != null) { cellToUpdate.Value = "Yes"; }

                                    smartsheet.SheetResources.RowResources.UpdateRows(TestingSheetData.Id.Value, new Row[] { updateRow });
                                }
                                if (status3.ToLower() == "approved")
                                {
                                    long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(TestingSheetData, "f&b approved");
                                    Cell cellToUpdateB = new() { ColumnId = honorariumSubmittedColumnId, Value = "Yes" };
                                    Row updateRow = new() { Id = targetRowId.Id, Cells = new Cell[] { cellToUpdateB } };
                                    Cell? cellToUpdate = targetRowId.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                                    if (cellToUpdate != null) { cellToUpdate.Value = "Yes"; }

                                    smartsheet.SheetResources.RowResources.UpdateRows(TestingSheetData.Id.Value, new Row[] { updateRow });
                                }
                                if (status4.ToLower() == "approved")
                                {
                                    long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(TestingSheetData, "HCP 1L Approved");
                                    Cell cellToUpdateB = new() { ColumnId = honorariumSubmittedColumnId, Value = "Yes" };
                                    Row updateRow = new() { Id = targetRowId.Id, Cells = new Cell[] { cellToUpdateB } };
                                    Cell? cellToUpdate = targetRowId.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                                    if (cellToUpdate != null) { cellToUpdate.Value = "Yes"; }

                                    smartsheet.SheetResources.RowResources.UpdateRows(TestingSheetData.Id.Value, new Row[] { updateRow });
                                }
                                if (status5.ToLower() == "approved")
                                {
                                    long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(TestingSheetData, "HCP 5L Approval");
                                    Cell cellToUpdateB = new() { ColumnId = honorariumSubmittedColumnId, Value = "Yes" };
                                    Row updateRow = new() { Id = targetRowId.Id, Cells = new Cell[] { cellToUpdateB } };
                                    Cell? cellToUpdate = targetRowId.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                                    if (cellToUpdate != null) { cellToUpdate.Value = "Yes"; }

                                    smartsheet.SheetResources.RowResources.UpdateRows(TestingSheetData.Id.Value, new Row[] { updateRow });
                                }
                                if (status6.ToLower() == "approved")
                                {
                                    long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(TestingSheetData, "HCP 6L Approval");
                                    Cell cellToUpdateB = new() { ColumnId = honorariumSubmittedColumnId, Value = "Yes" };
                                    Row updateRow = new() { Id = targetRowId.Id, Cells = new Cell[] { cellToUpdateB } };
                                    Cell? cellToUpdate = targetRowId.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                                    if (cellToUpdate != null) { cellToUpdate.Value = "Yes"; }

                                    smartsheet.SheetResources.RowResources.UpdateRows(TestingSheetData.Id.Value, new Row[] { updateRow });
                                }
                                if (status7.ToLower() == "approved")
                                {
                                    long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(TestingSheetData, "Trainer 12L Approval");
                                    Cell cellToUpdateB = new() { ColumnId = honorariumSubmittedColumnId, Value = "Yes" };
                                    Row updateRow = new() { Id = targetRowId.Id, Cells = new Cell[] { cellToUpdateB } };
                                    Cell? cellToUpdate = targetRowId.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                                    if (cellToUpdate != null) { cellToUpdate.Value = "Yes"; }

                                    smartsheet.SheetResources.RowResources.UpdateRows(TestingSheetData.Id.Value, new Row[] { updateRow });
                                }
                                if (status8.ToLower() == "approved")
                                {
                                    long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(TestingSheetData, "T/A Approval");
                                    Cell cellToUpdateB = new() { ColumnId = honorariumSubmittedColumnId, Value = "Yes" };
                                    Row updateRow = new() { Id = targetRowId.Id, Cells = new Cell[] { cellToUpdateB } };
                                    Cell? cellToUpdate = targetRowId.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                                    if (cellToUpdate != null) { cellToUpdate.Value = "Yes"; }

                                    smartsheet.SheetResources.RowResources.UpdateRows(TestingSheetData.Id.Value, new Row[] { updateRow });
                                }
                                if (status9.ToLower() == "approved")
                                {
                                    long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(TestingSheetData, "Vendor Update Approved?");
                                    Cell cellToUpdateB = new() { ColumnId = honorariumSubmittedColumnId, Value = "Yes" };
                                    Row updateRow = new() { Id = targetRowId.Id, Cells = new Cell[] { cellToUpdateB } };
                                    Cell? cellToUpdate = targetRowId.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                                    if (cellToUpdate != null) { cellToUpdate.Value = "Yes"; }

                                    smartsheet.SheetResources.RowResources.UpdateRows(TestingSheetData.Id.Value, new Row[] { updateRow });
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error($"Error occured on Webhook apicontroller Attachementfile method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
            }
        }

        private async void EventSettlementApproval(Root RequestWebhook)
        {
            try
            {
                string processSheet = configuration.GetSection("SmartsheetSettings:EventSettlement").Value;

                //var TestingId = "6831673324818308";
                Sheet TestingSheetData = SheetHelper.GetSheetById(smartsheet, processSheet);

                if (RequestWebhook != null && RequestWebhook.events != null)
                {
                    foreach (var WebHookEvent in RequestWebhook.events)
                    {
                        if (WebHookEvent.eventType.ToLower() == "updated" || WebHookEvent.eventType.ToLower() == "created")
                        {
                            Row targetRowId = TestingSheetData.Rows.FirstOrDefault(row => row.Id == WebHookEvent.rowId);
                            if (targetRowId != null)
                            {
                                string? status = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == 5946709627588484)?.Value?.ToString();
                                if (status.ToLower() == "approved")
                                {
                                    long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(TestingSheetData, "Is All Deviations Approved?");
                                    Cell cellToUpdateB = new() { ColumnId = honorariumSubmittedColumnId, Value = "Yes" };
                                    Row updateRow = new() { Id = targetRowId.Id, Cells = new Cell[] { cellToUpdateB } };
                                    Cell? cellToUpdate = targetRowId.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                                    if (cellToUpdate != null) { cellToUpdate.Value = "Yes"; }

                                    smartsheet.SheetResources.RowResources.UpdateRows(TestingSheetData.Id.Value, new Row[] { updateRow });
                                }
                            }
                        }
                    }
                }
            }

            catch (Exception ex)
            {
                Log.Error($"Error occured on Webhook apicontroller Attachementfile method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
            }
        }

        private async void EventSettlementDeviationApproval(Root RequestWebhook)
        {
            try
            {
                string processSheet = configuration.GetSection("SmartsheetSettings:EventSettlement").Value;


                Sheet TestingSheetData = SheetHelper.GetSheetById(smartsheet, processSheet);

                Column? Trigger = TestingSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "Deviation Status", StringComparison.OrdinalIgnoreCase));
                //Column Updatecolumn = TestingSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "Is All Deviations Approved?", StringComparison.OrdinalIgnoreCase));
                Column? Column1 = TestingSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "POST- Beyond30Days Deviation Approval", StringComparison.OrdinalIgnoreCase));
                Column? Column2 = TestingSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "POST-LessThan5Invitees Deviation Approval", StringComparison.OrdinalIgnoreCase));
                Column? Column3 = TestingSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "POST-Deviation Costperpaxabove1500 Approval ", StringComparison.OrdinalIgnoreCase));
                Column? Column4 = TestingSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "POST-Deviation Change in venue Approval", StringComparison.OrdinalIgnoreCase));
                Column? Column5 = TestingSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "POST-Deviation Change in topic Approval", StringComparison.OrdinalIgnoreCase));
                Column? Column6 = TestingSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "POST-Deviation Change in speaker Approval", StringComparison.OrdinalIgnoreCase));
                Column? Column7 = TestingSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "POST-Deviation Attendees not captured Approval", StringComparison.OrdinalIgnoreCase));
                Column? Column8 = TestingSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "POST-Deviation Speaker not captured  Approval", StringComparison.OrdinalIgnoreCase));
                Column? Column9 = TestingSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "POST-Deviation Other Deviation Approval", StringComparison.OrdinalIgnoreCase));

                if (RequestWebhook != null && RequestWebhook.events != null)
                {
                    foreach (var WebHookEvent in RequestWebhook.events)
                    {
                        if (WebHookEvent.eventType.ToLower() == "updated" || WebHookEvent.eventType.ToLower() == "created")
                        {
                            Row targetRowId = TestingSheetData.Rows.FirstOrDefault(row => row.Id == WebHookEvent.rowId);
                            //var columnValue = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == processIdColumn1.Id)?.Value.ToString();


                            if (targetRowId != null)
                            {
                                string? TriggerStatus = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == Trigger.Id)?.Value.ToString();
                                string? status1 = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == Column1.Id)?.Value?.ToString() ?? "Null";
                                string? status2 = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == Column2.Id)?.Value?.ToString() ?? "Null";
                                string? status3 = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == Column3.Id)?.Value?.ToString() ?? "Null";
                                string? status4 = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == Column4.Id)?.Value?.ToString() ?? "Null";
                                string? status5 = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == Column5.Id)?.Value?.ToString() ?? "Null";
                                string? status6 = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == Column6.Id)?.Value?.ToString() ?? "Null";
                                string? status7 = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == Column7.Id)?.Value?.ToString() ?? "Null";
                                string? status8 = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == Column8.Id)?.Value?.ToString() ?? "Null";
                                string? status9 = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == Column9.Id)?.Value?.ToString() ?? "Null";

                                if (status1.ToLower() == "approved")
                                {
                                    long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(TestingSheetData, "post 45 days approved");
                                    Cell cellToUpdateB = new() { ColumnId = honorariumSubmittedColumnId, Value = "Yes" };
                                    Row updateRow = new() { Id = targetRowId.Id, Cells = new Cell[] { cellToUpdateB } };
                                    Cell? cellToUpdate = targetRowId.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                                    if (cellToUpdate != null) { cellToUpdate.Value = "Yes"; }

                                    smartsheet.SheetResources.RowResources.UpdateRows(TestingSheetData.Id.Value, new Row[] { updateRow });
                                }
                                if (status2.ToLower() == "approved")
                                {
                                    long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(TestingSheetData, "Post <5 Invitees Approved");
                                    Cell cellToUpdateB = new() { ColumnId = honorariumSubmittedColumnId, Value = "Yes" };
                                    Row updateRow = new() { Id = targetRowId.Id, Cells = new Cell[] { cellToUpdateB } };
                                    Cell? cellToUpdate = targetRowId.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                                    if (cellToUpdate != null) { cellToUpdate.Value = "Yes"; }

                                    smartsheet.SheetResources.RowResources.UpdateRows(TestingSheetData.Id.Value, new Row[] { updateRow });
                                }
                                if (status3.ToLower() == "approved")
                                {
                                    long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(TestingSheetData, "Post CostperPax Approved");
                                    Cell cellToUpdateB = new() { ColumnId = honorariumSubmittedColumnId, Value = "Yes" };
                                    Row updateRow = new() { Id = targetRowId.Id, Cells = new Cell[] { cellToUpdateB } };
                                    Cell? cellToUpdate = targetRowId.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                                    if (cellToUpdate != null) { cellToUpdate.Value = "Yes"; }

                                    smartsheet.SheetResources.RowResources.UpdateRows(TestingSheetData.Id.Value, new Row[] { updateRow });
                                }
                                if (status4.ToLower() == "approved")
                                {
                                    long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(TestingSheetData, "Post ChangeInVenue Approved");
                                    Cell cellToUpdateB = new() { ColumnId = honorariumSubmittedColumnId, Value = "Yes" };
                                    Row updateRow = new() { Id = targetRowId.Id, Cells = new Cell[] { cellToUpdateB } };
                                    Cell? cellToUpdate = targetRowId.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                                    if (cellToUpdate != null) { cellToUpdate.Value = "Yes"; }

                                    smartsheet.SheetResources.RowResources.UpdateRows(TestingSheetData.Id.Value, new Row[] { updateRow });
                                }
                                if (status5.ToLower() == "approved")
                                {
                                    long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(TestingSheetData, "Post ChangeInTopic Approved");
                                    Cell cellToUpdateB = new() { ColumnId = honorariumSubmittedColumnId, Value = "Yes" };
                                    Row updateRow = new() { Id = targetRowId.Id, Cells = new Cell[] { cellToUpdateB } };
                                    Cell? cellToUpdate = targetRowId.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                                    if (cellToUpdate != null) { cellToUpdate.Value = "Yes"; }

                                    smartsheet.SheetResources.RowResources.UpdateRows(TestingSheetData.Id.Value, new Row[] { updateRow });
                                }
                                if (status6.ToLower() == "approved")
                                {
                                    long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(TestingSheetData, "Post ChangeInSpeaker Approved");
                                    Cell cellToUpdateB = new() { ColumnId = honorariumSubmittedColumnId, Value = "Yes" };
                                    Row updateRow = new() { Id = targetRowId.Id, Cells = new Cell[] { cellToUpdateB } };
                                    Cell? cellToUpdate = targetRowId.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                                    if (cellToUpdate != null) { cellToUpdate.Value = "Yes"; }

                                    smartsheet.SheetResources.RowResources.UpdateRows(TestingSheetData.Id.Value, new Row[] { updateRow });
                                }
                                if (status7.ToLower() == "approved")
                                {
                                    long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(TestingSheetData, "Post AttendeesNotCaptured Approved");
                                    Cell cellToUpdateB = new() { ColumnId = honorariumSubmittedColumnId, Value = "Yes" };
                                    Row updateRow = new() { Id = targetRowId.Id, Cells = new Cell[] { cellToUpdateB } };
                                    Cell? cellToUpdate = targetRowId.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                                    if (cellToUpdate != null) { cellToUpdate.Value = "Yes"; }

                                    smartsheet.SheetResources.RowResources.UpdateRows(TestingSheetData.Id.Value, new Row[] { updateRow });
                                }
                                if (status8.ToLower() == "approved")
                                {
                                    long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(TestingSheetData, "Post SpeakerNotCaptured Approved");
                                    Cell cellToUpdateB = new() { ColumnId = honorariumSubmittedColumnId, Value = "Yes" };
                                    Row updateRow = new() { Id = targetRowId.Id, Cells = new Cell[] { cellToUpdateB } };
                                    Cell? cellToUpdate = targetRowId.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                                    if (cellToUpdate != null) { cellToUpdate.Value = "Yes"; }

                                    smartsheet.SheetResources.RowResources.UpdateRows(TestingSheetData.Id.Value, new Row[] { updateRow });
                                }
                                if (status9.ToLower() == "approved")
                                {
                                    long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(TestingSheetData, "Post OtherDeviation Approved");
                                    Cell cellToUpdateB = new() { ColumnId = honorariumSubmittedColumnId, Value = "Yes" };
                                    Row updateRow = new() { Id = targetRowId.Id, Cells = new Cell[] { cellToUpdateB } };
                                    Cell? cellToUpdate = targetRowId.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                                    if (cellToUpdate != null) { cellToUpdate.Value = "Yes"; }

                                    smartsheet.SheetResources.RowResources.UpdateRows(TestingSheetData.Id.Value, new Row[] { updateRow });
                                }




                            }
                        }
                    }
                }
            }

            catch (Exception ex)
            {
                Log.Error($"Error occured on Webhook apicontroller Attachementfile method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
            }
        }

        private async void ApprovalCheckBox(Root RequestWebhook)
        {
            try
            {
                string processSheet = configuration.GetSection("SmartsheetSettings:HonorariumPayment").Value;

                //var TestingId = "6831673324818308";
                Sheet TestingSheetData = SheetHelper.GetSheetById(smartsheet, processSheet);

                if (RequestWebhook != null && RequestWebhook.events != null)
                {
                    foreach (var WebHookEvent in RequestWebhook.events)
                    {

                        if (WebHookEvent.eventType.ToLower() == "updated" || WebHookEvent.eventType.ToLower() == "created")
                        {
                            Row targetRowId = TestingSheetData.Rows.FirstOrDefault(row => row.Id == WebHookEvent.rowId);
                            if (targetRowId != null)
                            {
                                string? status = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == 7933735728009092)?.Value?.ToString();
                                if (status.ToLower() == "approved")
                                {
                                    long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(TestingSheetData, "5working days");
                                    Cell cellToUpdateB = new() { ColumnId = honorariumSubmittedColumnId, Value = "Yes" };
                                    Row updateRow = new() { Id = targetRowId.Id, Cells = new Cell[] { cellToUpdateB } };
                                    Cell? cellToUpdate = targetRowId.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                                    if (cellToUpdate != null) { cellToUpdate.Value = "Yes"; }

                                    smartsheet.SheetResources.RowResources.UpdateRows(TestingSheetData.Id.Value, new Row[] { updateRow });
                                }
                            }
                        }
                    }
                }
            }

            catch (Exception ex)
            {
                Log.Error($"Error occured on Webhook apicontroller Attachementfile method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
            }

        }

        private async void AgreementsTrigger(Root RequestWebhook)
        {
            try
            {
                //Sheet processSheetData = SheetHelper.GetSheetById(smartsheet, processSheet);
                //Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
                //Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                Sheet sheet_SpeakerCode = SheetHelper.GetSheetById(smartsheet, sheetId_SpeakerCode);

                if (RequestWebhook != null && RequestWebhook.events != null)
                {
                    foreach (var WebHookEvent in RequestWebhook.events)
                    {

                        if (WebHookEvent.eventType.ToLower() == "updated" /*|| WebHookEvent.eventType.ToLower() == "created"*/)
                        {
                            //var DataInSheet = smartsheet.SheetResources.GetSheet(sheet_SpeakerCode.Id.Value, null, null, new List<long> { WebHookEvent.rowId }, null, null, null, null, null, null).Rows;



                            Row targetRowId = sheet_SpeakerCode.Rows.FirstOrDefault(row => row.Id == WebHookEvent.rowId);


                            long ColumnId = SheetHelper.GetColumnIdByName(sheet_SpeakerCode, "Agreement Trigger");
                            Cell updatedCell = new()
                            {
                                ColumnId = ColumnId,
                                Value = "Yes"
                            };
                            Row updatedRow = new() { Id = targetRowId?.Id, Cells = new List<Cell> { updatedCell } };



                            smartsheet.SheetResources.RowResources.UpdateRows(sheet_SpeakerCode.Id.Value, new Row[] { updatedRow });





                        }
                    }
                }
            }

            catch (Exception ex)
            {
                Log.Error($"Error occured on Webhook apicontroller Attachementfile method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
            }

        }

        private async void Attachementfile(Root RequestWebhook)
        {
            try
            {
                Sheet processSheetData = SheetHelper.GetSheetById(smartsheet, processSheet);
                //Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
                //Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                //Sheet sheet_SpeakerCode = SheetHelper.GetSheetById(smartsheet, sheetId_SpeakerCode);
                if (RequestWebhook != null && RequestWebhook.events != null)
                {
                    foreach (var WebHookEvent in RequestWebhook.events)
                    {


                        if (WebHookEvent.eventType.ToLower() == "updated" || WebHookEvent.eventType.ToLower() == "created")
                        {
                            //var DataInSheet = smartsheet.SheetResources.GetSheet(parsedProcessSheet, null, null, new List<long> { WebHookEvent.rowId }, null, null, null, null, null, null).Rows;


                            Column processIdColumn1 = processSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "EventId/EventRequestId", StringComparison.OrdinalIgnoreCase));
                            Column processIdColumn2 = processSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "Event Request Status", StringComparison.OrdinalIgnoreCase));
                            Column processIdColumn3 = processSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "Meeting Type", StringComparison.OrdinalIgnoreCase));
                            Column processIdColumn4 = processSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "Event Type", StringComparison.OrdinalIgnoreCase));


                            Row targetRowId = processSheetData.Rows.FirstOrDefault(row => row.Id == WebHookEvent.rowId);



                            if (processIdColumn1 != null && processIdColumn2 != null)
                            {
                                string? columnValue = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == processIdColumn1.Id)?.Value.ToString();
                                string? status = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == processIdColumn2.Id)?.Value.ToString();
                                object? meetingType = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == processIdColumn3.Id)?.Value;
                                string? EventType = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == processIdColumn4.Id)?.Value.ToString();
                                if (EventType == "Class I" || EventType == "Webinar")
                                {
                                    if (status != null && (status == "Marketing Head Approved" || status == "Waiting for Finance Treasury Approval"))

                                    {
                                        int timeInterval = 60000;
                                        await Task.Delay(timeInterval);
                                        if (meetingType != null)
                                        {
                                            if (meetingType.ToString().ToLower().Contains("other"))
                                            {

                                                moveAttachments(columnValue, WebHookEvent.rowId);
                                            }
                                            else
                                            {
                                                GenerateSummaryPDF(columnValue, WebHookEvent.rowId);
                                                moveAttachments(columnValue, WebHookEvent.rowId);
                                            }
                                        }

                                        else
                                        {
                                            GenerateSummaryPDF(columnValue, WebHookEvent.rowId);
                                            moveAttachments(columnValue, WebHookEvent.rowId);
                                        }


                                    }

                                }

                                else if (status != null && status == "Marketing Head Approved" || status == "Waiting for Finance Treasury Approval")
                                {
                                    int timeInterval = 60000;
                                    await Task.Delay(timeInterval);
                                    moveAttachments(columnValue, WebHookEvent.rowId);
                                }

                            }


                        }
                    }
                }
            }

            catch (Exception ex)
            {
                Log.Error($"Error occured on Webhook apicontroller Attachementfile method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
            }

        }

        private async void moveAttachments(string EventID, long rowId)
        {
            List<Attachment> attachments = new List<Attachment>();
            Sheet processSheetData = SheetHelper.GetSheetById(smartsheet, processSheet);
            //Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
            //Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
            Sheet sheet_SpeakerCode = SheetHelper.GetSheetById(smartsheet, sheetId_SpeakerCode);

            foreach (Row row in sheet_SpeakerCode.Rows)
            {
                Cell matchingCell = row.Cells.FirstOrDefault(cell => cell.DisplayValue == EventID);

                if (matchingCell != null && matchingCell.Value != null)
                {
                    long Id = (long)row.Id;
                    string eventId = matchingCell.Value.ToString();
                    if (!string.IsNullOrEmpty(eventId) && eventId.Equals(EventID, StringComparison.OrdinalIgnoreCase))
                    {
                        PaginatedResult<Attachment> a = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments((long)sheet_SpeakerCode.Id, Id, null);
                        string url = "";
                        string name = "";
                        foreach (var x in a.Data)
                        {
                            if (x != null)
                            {
                                long AID = (long)x.Id;
                                Attachment file = smartsheet.SheetResources.AttachmentResources.GetAttachment((long)sheet_SpeakerCode.Id, AID);
                                //string filename = file.Name.Split(".")[0].Split("-")[1];
                                string FullName = file.Name;
                                string substring = "agreement";
                                bool test = FullName.Contains(substring);
                                if (test)
                                {
                                    url = file.Url;
                                    name = file.Name;
                                }

                            }
                            if (url != "")
                            {
                                using (HttpClient client = new HttpClient())
                                {
                                    byte[] data = client.GetByteArrayAsync(url).Result;
                                    string base64 = Convert.ToBase64String(data);
                                    byte[] xy = Convert.FromBase64String(base64);
                                    string f = Path.Combine("Resources", "Images");
                                    string ps = Path.Combine(Directory.GetCurrentDirectory(), f);
                                    if (!Directory.Exists(ps))
                                    {
                                        Directory.CreateDirectory(ps);
                                    }
                                    string ft = SheetHelper.GetFileType(xy);
                                    string fileName = name;
                                    string fp = Path.Combine(ps, fileName);

                                    System.IO.File.WriteAllBytes(fp, xy);
                                    string type = SheetHelper.GetContentType(ft);
                                    Attachment z = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile((long)processSheetData.Id, rowId, fp, "application/msword");

                                }
                                url = "";
                                string bs64 = "";
                            }
                        }
                    }
                }
            }
        }

        private async void GenerateSummaryPDF(string EventID, long rowId)
        {
            try
            {
                Sheet processSheetData = SheetHelper.GetSheetById(smartsheet, processSheet);
                Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                Sheet sheet_SpeakerCode = SheetHelper.GetSheetById(smartsheet, sheetId_SpeakerCode);

                string? EventCode = "";
                string? EventName = "";
                string EventDate = "";
                string? EventVenue = "";
                DateTime parsedDate;
                List<string> Speakers = new List<string>();




                Column SpecialityColumn = sheet1.Columns.FirstOrDefault(column => string.Equals(column.Title, "EventId/EventRequestId", StringComparison.OrdinalIgnoreCase));
                Column targetColumn1 = sheet1.Columns.FirstOrDefault(column => string.Equals(column.Title, "Event Topic", StringComparison.OrdinalIgnoreCase));
                Column targetColumn2 = sheet1.Columns.FirstOrDefault(column => string.Equals(column.Title, "Event Date", StringComparison.OrdinalIgnoreCase));
                Column targetColumn3 = sheet1.Columns.FirstOrDefault(column => string.Equals(column.Title, "Venue Name", StringComparison.OrdinalIgnoreCase));

                if (SpecialityColumn != null)
                {
                    Row targetRow = sheet1.Rows
                     .FirstOrDefault(row => row.Cells?.Any(cell => cell.ColumnId == SpecialityColumn.Id && cell.Value?.ToString() == EventID) == true);

                    if (targetRow != null)
                    {
                        string originalDate = targetRow.Cells.FirstOrDefault(cell => cell.ColumnId == targetColumn2.Id)?.Value?.ToString();
                        if (!string.IsNullOrEmpty(originalDate))
                        {

                            if (DateTime.TryParseExact(originalDate, new string[] { "yyyy-MM-dd", "dd/MM/yyyy", "MM/dd/yyyy" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate))
                            {
                                string formattedDate = parsedDate.ToString("dd/MM/yyyy");
                                EventDate = formattedDate;
                            }
                        }

                        EventCode = targetRow.Cells.FirstOrDefault(cell => cell.ColumnId == SpecialityColumn.Id)?.Value?.ToString();
                        EventName = targetRow.Cells.FirstOrDefault(cell => cell.ColumnId == targetColumn1.Id)?.Value?.ToString();
                        //EventDate = targetRow.Cells.FirstOrDefault(cell => cell.ColumnId == targetColumn2.Id)?.Value?.ToString();
                        //string formattedDate = EventDate.ToString("dd/MM/yyyy");
                        EventVenue = targetRow.Cells.FirstOrDefault(cell => cell.ColumnId == targetColumn3.Id)?.Value?.ToString();





                    }
                }

                List<string> requiredColumns = new List<string> { "HCPName", "MISCode", "Speciality", "HCP Type" };
                List<string> MenariniColumns = new List<string> { "HCPName", "Employee Code", "Designation" };

                List<Column> selectedColumns = sheet_SpeakerCode.Columns.Where(column => requiredColumns.Contains(column.Title, StringComparer.OrdinalIgnoreCase)).ToList();
                List<Column> selectedMenariniColumns = sheet.Columns.Where(column => MenariniColumns.Contains(column.Title, StringComparer.OrdinalIgnoreCase)).ToList();

                DataTable dtMai = new DataTable();
                dtMai.Columns.Add("S.No", typeof(int));
                foreach (Column column in selectedColumns)
                {
                    dtMai.Columns.Add(column.Title);
                }
                dtMai.Columns.Add("Sign");

                DataTable MenariniTable = new DataTable();
                MenariniTable.Columns.Add("S.No", typeof(int));
                foreach (Column column in selectedMenariniColumns)
                {
                    MenariniTable.Columns.Add(column.Title);
                }
                MenariniTable.Columns.Add("Sign");

                int Sr_No = 1;
                int m_no = 1;

                foreach (Row row in sheet_SpeakerCode.Rows)
                {
                    string eventId = row.Cells.FirstOrDefault(cell => sheet_SpeakerCode.Columns.FirstOrDefault(c => c.Id == cell.ColumnId)?.Title == "EventId/EventRequestId")?.DisplayValue;
                    if (!string.IsNullOrEmpty(eventId) && eventId.Equals(EventID, StringComparison.OrdinalIgnoreCase))
                    {
                        DataRow newRow = dtMai.NewRow();
                        newRow["S.No"] = Sr_No;
                        foreach (Cell cell in row.Cells)
                        {
                            string columnName = sheet_SpeakerCode.Columns.FirstOrDefault(c => c.Id == cell.ColumnId)?.Title;
                            if (requiredColumns.Contains(columnName, StringComparer.OrdinalIgnoreCase))
                            {
                                if (columnName == "HCPName")
                                {
                                    var val = cell.DisplayValue;
                                    Speakers.Add(val);
                                }
                                newRow[columnName] = cell.DisplayValue;
                            }
                        }
                        dtMai.Rows.Add(newRow);
                        Sr_No++;
                    }
                }
                foreach (Row row in sheet.Rows)
                {
                    string eventId = row.Cells.FirstOrDefault(cell => sheet.Columns.FirstOrDefault(c => c.Id == cell.ColumnId)?.Title == "EventId/EventRequestId")?.DisplayValue;
                    string InviteeSource = row.Cells.FirstOrDefault(cell => sheet.Columns.FirstOrDefault(c => c.Id == cell.ColumnId)?.Title == "Invitee Source")?.DisplayValue;
                    if (!string.IsNullOrEmpty(eventId) && eventId.Equals(EventID, StringComparison.OrdinalIgnoreCase) && !InviteeSource.Equals("Menarini Employees", StringComparison.OrdinalIgnoreCase))
                    {
                        DataRow newRow = dtMai.NewRow();
                        newRow["S.No"] = Sr_No;
                        foreach (Cell cell in row.Cells)
                        {
                            string columnName = sheet.Columns.FirstOrDefault(c => c.Id == cell.ColumnId)?.Title;
                            if (requiredColumns.Contains(columnName, StringComparer.OrdinalIgnoreCase))
                            {
                                newRow[columnName] = cell.DisplayValue;
                            }
                        }
                        dtMai.Rows.Add(newRow);
                        Sr_No++;
                    }
                    else if (!string.IsNullOrEmpty(eventId) && eventId.Equals(EventID, StringComparison.OrdinalIgnoreCase) && InviteeSource.Equals("Menarini Employees", StringComparison.OrdinalIgnoreCase))
                    {
                        DataRow newRow = MenariniTable.NewRow();
                        newRow["S.No"] = m_no;
                        foreach (Cell cell in row.Cells)
                        {
                            string columnName = sheet.Columns
                                .FirstOrDefault(c => c.Id == cell.ColumnId)?.Title;
                            if (MenariniColumns.Contains(columnName, StringComparer.OrdinalIgnoreCase))
                            {
                                newRow[columnName] = cell.DisplayValue;
                            }
                        }
                        MenariniTable.Rows.Add(newRow);
                        m_no++;
                    }






                }

                string resultString = string.Join(", ", Speakers);

                byte[] fileBytes = SheetHelper.exportAttendencepdfnew(dtMai, MenariniTable, EventCode, EventName, EventDate, EventVenue, resultString);
                string filename = "Attendance Sheet_" + EventID + ".pdf";
                string folderName = Path.Combine("Resources", "Images");
                string pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                if (!Directory.Exists(pathToSave))
                {
                    Directory.CreateDirectory(pathToSave);
                }
                string fileType = SheetHelper.GetFileType(fileBytes);
                string filePath = Path.Combine(pathToSave, filename);
                System.IO.File.WriteAllBytes(filePath, fileBytes);

                var a = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments((long)processSheetData.Id, rowId, null);
                foreach (var x in a.Data)
                {
                    long Id = (long)x.Id;
                    var Fullname = x.Name.Split("-");
                    string? splitName = Fullname[0];

                    if (splitName.ToLower().Contains("attendance sheet"))
                    {

                        //var ExistingFile = smartsheet.SheetResources.AttachmentResources.GetAttachment((long)processSheetData.Id, Id);
                        //url = ExistingFile.Url;

                        smartsheet.SheetResources.AttachmentResources.DeleteAttachment(
                          (long)processSheetData.Id,           // sheetId
                          Id            // attachmentId
                        );

                    }
                }
                Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile((long)processSheetData.Id, rowId, filePath, "application/msword");
                if (System.IO.File.Exists(filePath))
                {
                    System.IO.File.Delete(filePath);
                }

            }
            catch (Exception ex)
            {
                //return BadRequest(ex.Message);
            }
        }


        private async void MailChange(Root RequestWebhook)
        {

            string eventSettlement = configuration.GetSection("SmartsheetSettings:EventSettlement").Value;
            string honorarium = configuration.GetSection("SmartsheetSettings:HonorariumPayment").Value;
            string Deviation = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;
            string SpeakerCodeCreation = configuration.GetSection("SmartsheetSettings:SpeakerCodeCreation").Value;
            string ApprovedSpeakers = configuration.GetSection("SmartsheetSettings:ApprovedSpeakers").Value;
            string TrainerCodeCreation = configuration.GetSection("SmartsheetSettings:TrainerCodeCreation").Value;
            string ApprovedTrainers = configuration.GetSection("SmartsheetSettings:ApprovedTrainers").Value;
            string VendorMasterSheet = configuration.GetSection("SmartsheetSettings:VendorMasterSheet").Value;
            string VendorCodeCreation = configuration.GetSection("SmartsheetSettings:VendorCodeCreation").Value;

            Sheet sheet = SheetHelper.GetSheetById(smartsheet, processSheet);
            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, honorarium);
            Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, eventSettlement);
            Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, Deviation);
            Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, SpeakerCodeCreation);
            Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, ApprovedSpeakers);
            Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, TrainerCodeCreation);
            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, ApprovedTrainers);
            Sheet sheet8 = SheetHelper.GetSheetById(smartsheet, VendorMasterSheet);
            Sheet sheet9 = SheetHelper.GetSheetById(smartsheet, VendorCodeCreation);



            string wehookSheetId = "" + RequestWebhook.scopeObjectId;
            Sheet WebHookSheet = SheetHelper.GetSheetById(smartsheet, wehookSheetId);



            var Sheetcolumns = sheet.Columns.ToDictionary(column => column.Title, column => (long)column.Id);
            var Sheetcolumns1 = sheet1.Columns.ToDictionary(column => column.Title, column => (long)column.Id);
            var Sheetcolumns2 = sheet2.Columns.ToDictionary(column => column.Title, column => (long)column.Id);
            var Sheetcolumns3 = sheet3.Columns.ToDictionary(column => column.Title, column => (long)column.Id);
            var Sheetcolumns4 = sheet4.Columns.ToDictionary(column => column.Title, column => (long)column.Id);
            var Sheetcolumns5 = sheet5.Columns.ToDictionary(column => column.Title, column => (long)column.Id);
            var Sheetcolumns6 = sheet6.Columns.ToDictionary(column => column.Title, column => (long)column.Id);
            var Sheetcolumns7 = sheet7.Columns.ToDictionary(column => column.Title, column => (long)column.Id);
            var Sheetcolumns8 = sheet8.Columns.ToDictionary(column => column.Title, column => (long)column.Id);
            var Sheetcolumns9 = sheet9.Columns.ToDictionary(column => column.Title, column => (long)column.Id);

            //Dictionary<string, long> Sheetcolumns9 = new();
            //foreach (Column? column in sheet9.Columns)
            //{
            //    Sheetcolumns9.Add(column.Title, (long)column.Id);
            //}
            if (RequestWebhook != null && RequestWebhook.events != null)
            {

                foreach (var WebHookEvent in RequestWebhook.events)
                {
                    if (WebHookEvent.eventType.ToLower() == "updated" || WebHookEvent.eventType.ToLower() == "created")
                    {
                        long RowId = WebHookEvent.rowId;
                        string processSheet = await Task.Run(() => processSheetEmailChange(RowId, WebHookSheet, sheet, Sheetcolumns, smartsheet));
                        string HonorariumSheet = await Task.Run(() => HonorariumEmailChange(RowId, WebHookSheet, sheet1, Sheetcolumns1, smartsheet));
                        string EventSettlementSheet = await Task.Run(() => EventSettlementEmailChange(RowId, WebHookSheet, sheet2, Sheetcolumns2, smartsheet));
                        string DeviationSheet = await Task.Run(() => DeviatonEmailChange(RowId, WebHookSheet, sheet3, Sheetcolumns3, smartsheet));
                        string SpeakerCodeCreationSheet = await Task.Run(() => ApprovedSpeakersEmailChange(RowId, WebHookSheet, sheet4, Sheetcolumns4, smartsheet));
                        string ApprovedSpeakersSheet = await Task.Run(() => ApprovedSpeakersEmailChange(RowId, WebHookSheet, sheet5, Sheetcolumns5, smartsheet));
                        string TrainerCodeCreationSheet = await Task.Run(() => ApprovedSpeakersEmailChange(RowId, WebHookSheet, sheet6, Sheetcolumns6, smartsheet));
                        string ApprovedTrainersSheet = await Task.Run(() => ApprovedSpeakersEmailChange(RowId, WebHookSheet, sheet7, Sheetcolumns7, smartsheet));
                        string VendorMasterSheetSheet = await Task.Run(() => VendorEmailChange(RowId, WebHookSheet, sheet8, Sheetcolumns8, smartsheet));
                        string VendorCodeCreationSheet = await Task.Run(() => VendorCodeEmailChange(RowId, WebHookSheet, sheet9, Sheetcolumns9, smartsheet));
                        #region
                        //string DesignationValue = "";
                        //string EmailValue = "";
                        //Row row = GetRowById(WebHookSheet, RowId);
                        //Column? designationColumn = WebHookSheet.Columns.FirstOrDefault(c => c.Title == "Designation");
                        //Column? EmailColumn = WebHookSheet.Columns.FirstOrDefault(c => c.Title == "Email");

                        ////string processSheet = await Task.Run(() => processSheetEmailChange(RowId, WebHookSheet, sheet, Sheetcolumns, smartsheet));

                        //if (designationColumn != null && EmailColumn != null)
                        //{
                        //    int? designationColumnIndex = designationColumn.Index;
                        //    int? EmailColumnIndex = EmailColumn.Index;
                        //    DesignationValue = row.Cells[(int)designationColumnIndex].Value.ToString();
                        //    EmailValue = row.Cells[(int)EmailColumnIndex].Value.ToString();
                        //    string columnName = "";
                        //    if (DesignationValue == "Sales Head")
                        //        columnName = "PRE-Sales Head Approval";
                        //    else if (DesignationValue == "Marketing Head")
                        //        columnName = "PRE-Marketing Head Approval";
                        //    else if (DesignationValue == "Finance Treasury")
                        //        columnName = "PRE-Finance Treasury Approval";
                        //    else if (DesignationValue == "Medical Affairs Head")
                        //        columnName = "PRE-Medical Affairs Head Approval";

                        //    IEnumerable<Row> DataInSheet1 = [];
                        //    int? statusColumnIndex = sheet.Columns.Where(y => y.Title == "Event Request Status").Select(z => z.Index).FirstOrDefault();
                        //    int? designationStatusIndex = sheet.Columns.Where(y => y.Title == columnName).Select(z => z.Index).FirstOrDefault();

                        //    DataInSheet1 = sheet.Rows.Where(x =>
                        //    {
                        //        string cellValue = Convert.ToString(x.Cells[(int)statusColumnIndex].Value).ToLower();
                        //        string designationValue = Convert.ToString(x.Cells[(int)statusColumnIndex].Value).ToLower();
                        //        return ((cellValue != "approved" || cellValue != "advance approved") && designationValue != "approved");
                        //    });
                        //    List<Row> liRowsToUpdate = new();
                        //    foreach (Row rowData in DataInSheet1)
                        //    {
                        //        Cell[] cellsToUpdate = new Cell[]
                        //        { new Cell {  ColumnId = Sheetcolumns[DesignationValue], Value = EmailValue  }};

                        //        row = new Row
                        //        {
                        //            Id = rowData.Id,
                        //            Cells = cellsToUpdate
                        //        };
                        //        liRowsToUpdate.Add(row);
                        //    }
                        //    if (liRowsToUpdate.Count > 0)
                        //    {
                        //        smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, liRowsToUpdate);
                        //        Log.Information("updated " + liRowsToUpdate.Count + " rows");
                        //    }
                        //}
                        #endregion
                    }
                }

            }
        }
        private async void MailChangeInMasters(Root RequestWebhook)
        {

            string eventSettlement = configuration.GetSection("SmartsheetSettings:EventSettlement").Value;
            string honorarium = configuration.GetSection("SmartsheetSettings:HonorariumPayment").Value;
            string Deviation = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, processSheet);
            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, honorarium);
            Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, eventSettlement);
            string wehookSheetId = "" + RequestWebhook.scopeObjectId;
            Sheet WebHookSheet = SheetHelper.GetSheetById(smartsheet, wehookSheetId);

            Dictionary<string, long> Sheetcolumns = sheet.Columns.ToDictionary(column => column.Title, column => (long)column.Id);
            Dictionary<string, long> Sheetcolumns1 = sheet1.Columns.ToDictionary(column => column.Title, column => (long)column.Id);
            Dictionary<string, long> Sheetcolumns2 = sheet2.Columns.ToDictionary(column => column.Title, column => (long)column.Id);
            if (RequestWebhook != null && RequestWebhook.events != null)
            {
                foreach (var WebHookEvent in RequestWebhook.events)
                {
                    if (WebHookEvent.eventType.ToLower() == "updated" || WebHookEvent.eventType.ToLower() == "created")
                    {
                        long RowId = WebHookEvent.rowId;
                        string processSheet = await Task.Run(() => TestprocessSheetEmailChange(RowId, WebHookSheet, sheet, Sheetcolumns, smartsheet));
                        string HonorariumSheet = await Task.Run(() => TestHonorariumEmailChange(RowId, WebHookSheet, sheet1, Sheetcolumns1, smartsheet));
                        string EventSettlementSheet = await Task.Run(() => TestEventSettlementEmailChange(RowId, WebHookSheet, sheet2, Sheetcolumns2, smartsheet));


                    }
                }
            }
        }

        public static string processSheetEmailChange(long rowId, Sheet WebHookSheet, Sheet sheet,
        Dictionary<string, long> Sheetcolumns, SmartsheetClient smartsheet)
        {
            long RowId = rowId;
            string DesignationValue = "";
            string EmailValue = "";

            Row row = GetRowById(WebHookSheet, RowId);
            Column? designationColumn = WebHookSheet.Columns.FirstOrDefault(c => c.Title == "Designation"); //Approval master Role and should be the column name in respective sheet
            Column? EmailColumn = WebHookSheet.Columns.FirstOrDefault(c => c.Title == "Email"); // Approval master Email

            if (designationColumn != null && EmailColumn != null)
            {
                int? designationColumnIndex = designationColumn.Index;
                int? EmailColumnIndex = EmailColumn.Index;
                DesignationValue = row.Cells[(int)designationColumnIndex].Value.ToString();
                EmailValue = row.Cells[(int)EmailColumnIndex].Value.ToString();

                string columnName = "";
                if (DesignationValue == "Sales Head")
                    columnName = "PRE-Sales Head Approval";
                else if (DesignationValue == "Marketing Head")
                    columnName = "PRE-Marketing Head Approval";
                else if (DesignationValue == "Finance Treasury")
                    columnName = "PRE-Finance Treasury Approval";
                else if (DesignationValue == "Medical Affairs Head")
                    columnName = "PRE-Medical Affairs Head Approval";
                else if (DesignationValue == "Compliance")
                    columnName = "PRE-Compliance Approval";
                if (columnName != "")
                {


                    IEnumerable<Row> DataInSheet1 = [];
                    int? statusColumnIndex = sheet.Columns.Where(y => y.Title == "Event Request Status").Select(z => z.Index).FirstOrDefault();
                    int? designationStatusIndex = sheet.Columns.Where(y => y.Title == columnName).Select(z => z.Index).FirstOrDefault();

                    DataInSheet1 = sheet.Rows.Where(x =>
                    {
                        string cellValue = Convert.ToString(x.Cells[(int)statusColumnIndex].Value).ToLower();
                        string designationValue = Convert.ToString(x.Cells[(int)designationStatusIndex].Value).ToLower();
                        return ((cellValue != "approved" || cellValue != "advance approved") && designationValue != "approved");
                    });

                    List<Row> liRowsToUpdate = new();
                    foreach (Row rowData in DataInSheet1)
                    {
                        Cell[] cellsToUpdate = new Cell[]
                        { new Cell {  ColumnId = Sheetcolumns[DesignationValue], Value = EmailValue  }};

                        row = new Row
                        {
                            Id = rowData.Id,
                            Cells = cellsToUpdate
                        };
                        liRowsToUpdate.Add(row);
                    }
                    if (liRowsToUpdate.Count > 0)
                    {
                        ApiCalls.BulkUpdateRows(smartsheet, sheet, liRowsToUpdate);
                        //smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, liRowsToUpdate);
                        Log.Information("updated " + liRowsToUpdate.Count + " rows");
                    }
                }
                else
                {
                    return "designation not found";
                }
            }


            return "process sheet updated";
        }

        public static string HonorariumEmailChange(long rowId, Sheet WebHookSheet, Sheet sheet,
            Dictionary<string, long> Sheetcolumns, SmartsheetClient smartsheet)
        {
            long RowId = rowId;
            string DesignationValue = "";
            string EmailValue = "";

            Row row = GetRowById(WebHookSheet, RowId);
            Column? designationColumn = WebHookSheet.Columns.FirstOrDefault(c => c.Title == "Designation");
            Column? EmailColumn = WebHookSheet.Columns.FirstOrDefault(c => c.Title == "Email");

            if (designationColumn != null && EmailColumn != null)
            {
                int? designationColumnIndex = designationColumn.Index;
                int? EmailColumnIndex = EmailColumn.Index;
                DesignationValue = row.Cells[(int)designationColumnIndex].Value.ToString();
                EmailValue = row.Cells[(int)EmailColumnIndex].Value.ToString();
                string columnName = "";
                if (DesignationValue == "Sales Head")
                    columnName = "HON-Sales Head Approval";
                else if (DesignationValue == "Marketing Head")
                    columnName = "HON-Marketing Head Approval";
                else if (DesignationValue == "Finance Treasury")
                    columnName = "HON-Finance Treasury Approval";
                else if (DesignationValue == "Medical Affairs Head")
                    columnName = "HON-Medical Affairs Head Approval";
                else if (DesignationValue == "Compliance")
                    columnName = "HON-Compliance Approval";
                else if (DesignationValue == "Finance Accounts")
                    columnName = "HON-Finance Accounts Approval";
                if (columnName != "")
                {
                    IEnumerable<Row> DataInSheet1 = [];
                    int? statusColumnIndex = sheet.Columns.Where(y => y.Title == "Honorarium Request Status").Select(z => z.Index).FirstOrDefault();
                    int? designationStatusIndex = sheet.Columns.Where(y => y.Title == columnName).Select(z => z.Index).FirstOrDefault();

                    DataInSheet1 = sheet.Rows.Where(x =>
                    {
                        string cellValue = Convert.ToString(x.Cells[(int)statusColumnIndex].Value).ToLower();
                        string designationValue = Convert.ToString(x.Cells[(int)designationStatusIndex].Value).ToLower();
                        return ((cellValue != "honorarium approved") && designationValue != "approved");
                    });
                    List<Row> liRowsToUpdate = new();
                    foreach (Row rowData in DataInSheet1)
                    {
                        Cell[] cellsToUpdate = new Cell[]
                        { new Cell {  ColumnId = Sheetcolumns[DesignationValue], Value = EmailValue  }};

                        row = new Row
                        {
                            Id = rowData.Id,
                            Cells = cellsToUpdate
                        };
                        liRowsToUpdate.Add(row);
                    }
                    if (liRowsToUpdate.Count > 0)
                    {
                        ApiCalls.BulkUpdateRows(smartsheet, sheet, liRowsToUpdate);
                        //smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, liRowsToUpdate);
                        Log.Information("updated " + liRowsToUpdate.Count + " rows");
                    }
                }
                else
                {
                    return "designation not found";
                }
            }


            return "Honorarium sheet updated";
        }

        public static string EventSettlementEmailChange(long rowId, Sheet WebHookSheet, Sheet sheet,
            Dictionary<string, long> Sheetcolumns, SmartsheetClient smartsheet)
        {
            long RowId = rowId;
            string DesignationValue = "";
            string EmailValue = "";

            Row row = GetRowById(WebHookSheet, RowId);
            Column? designationColumn = WebHookSheet.Columns.FirstOrDefault(c => c.Title == "Designation");
            Column? EmailColumn = WebHookSheet.Columns.FirstOrDefault(c => c.Title == "Email");

            if (designationColumn != null && EmailColumn != null)
            {
                int? designationColumnIndex = designationColumn.Index;
                int? EmailColumnIndex = EmailColumn.Index;
                DesignationValue = row.Cells[(int)designationColumnIndex].Value.ToString();
                EmailValue = row.Cells[(int)EmailColumnIndex].Value.ToString();
                string columnName = "";
                if (DesignationValue == "Sales Head")
                    columnName = "EventSettlement-SalesHead Approval";
                else if (DesignationValue == "Marketing Head")
                    columnName = "EventSettlement-Marketing Head Approval";
                else if (DesignationValue == "Finance Treasury")
                    columnName = "EventSettlement-Finance Treasury Approval";
                else if (DesignationValue == "Medical Affairs Head")
                    columnName = "EventSettlement-Medical Affairs Head Approval";
                else if (DesignationValue == "Compliance")
                    columnName = "EventSettlement-Compliance Approval";
                else if (DesignationValue == "Finance Accounts")
                    columnName = "EventSettlement-Finance Account Approval";
                if (columnName != "")
                {
                    IEnumerable<Row> DataInSheet1 = [];
                    int? statusColumnIndex = sheet.Columns.Where(y => y.Title == "Post Event Request status").Select(z => z.Index).FirstOrDefault();
                    int? designationStatusIndex = sheet.Columns.Where(y => y.Title == columnName).Select(z => z.Index).FirstOrDefault();

                    DataInSheet1 = sheet.Rows.Where(x =>
                    {
                        string cellValue = Convert.ToString(x.Cells[(int)statusColumnIndex].Value).ToLower();
                        string designationValue = Convert.ToString(x.Cells[(int)designationStatusIndex].Value).ToLower();
                        return ((cellValue != "closure approved") && designationValue != "approved");
                    });
                    List<Row> liRowsToUpdate = new();
                    foreach (Row rowData in DataInSheet1)
                    {
                        Cell[] cellsToUpdate = new Cell[]
                        { new Cell {  ColumnId = Sheetcolumns[DesignationValue], Value = EmailValue  }};

                        row = new Row
                        {
                            Id = rowData.Id,
                            Cells = cellsToUpdate
                        };
                        liRowsToUpdate.Add(row);
                    }
                    if (liRowsToUpdate.Count > 0)
                    {
                        //smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, liRowsToUpdate);
                        ApiCalls.BulkUpdateRows(smartsheet, sheet, liRowsToUpdate);
                        Log.Information("updated " + liRowsToUpdate.Count + " rows");
                    }
                }
                else
                {
                    return "designation not found";
                }
            }


            return "Event Settlement sheet updated";
        }

        public static string DeviatonEmailChange(long rowId, Sheet WebHookSheet, Sheet sheet,
           Dictionary<string, long> Sheetcolumns, SmartsheetClient smartsheet)
        {
            long RowId = rowId;
            string DesignationValue = "";
            string EmailValue = "";

            Row row = GetRowById(WebHookSheet, RowId);
            Column? designationColumn = WebHookSheet.Columns.FirstOrDefault(c => c.Title == "Designation");
            Column? EmailColumn = WebHookSheet.Columns.FirstOrDefault(c => c.Title == "Email");

            if (designationColumn != null && EmailColumn != null)
            {
                int? designationColumnIndex = designationColumn.Index;
                int? EmailColumnIndex = EmailColumn.Index;
                DesignationValue = row.Cells[(int)designationColumnIndex].Value.ToString();
                EmailValue = row.Cells[(int)EmailColumnIndex].Value.ToString();
                string columnName = "";
                if (DesignationValue == "Sales Head")
                    columnName = "Sales Head approval";
                else if (DesignationValue == "Finance Head")
                    columnName = "Finance Head Approval";

                if (columnName != "")
                {
                    IEnumerable<Row> DataInSheet1 = [];
                    // int? statusColumnIndex = sheet.Columns.Where(y => y.Title == "Post Event Request status").Select(z => z.Index).FirstOrDefault();
                    int? designationStatusIndex = sheet.Columns.Where(y => y.Title == columnName).Select(z => z.Index).FirstOrDefault();

                    DataInSheet1 = sheet.Rows.Where(x =>
                    {
                        //string cellValue = Convert.ToString(x.Cells[(int)statusColumnIndex].Value).ToLower();
                        string designationValue = Convert.ToString(x.Cells[(int)designationStatusIndex].Value).ToLower();
                        return (designationValue != "approved");
                    });
                    List<Row> liRowsToUpdate = new();
                    foreach (Row rowData in DataInSheet1)
                    {
                        Cell[] cellsToUpdate = new Cell[]
                        { new Cell {  ColumnId = Sheetcolumns[DesignationValue], Value = EmailValue  }};

                        row = new Row
                        {
                            Id = rowData.Id,
                            Cells = cellsToUpdate
                        };
                        liRowsToUpdate.Add(row);
                    }
                    if (liRowsToUpdate.Count > 0)
                    {
                        //smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, liRowsToUpdate);
                        ApiCalls.BulkUpdateRows(smartsheet, sheet, liRowsToUpdate);
                        Log.Information("updated " + liRowsToUpdate.Count + " rows");
                    }
                }
                else
                {
                    return "designation not found";
                }
            }


            return "Event Settlement sheet updated";
        }

        public static string ApprovedSpeakersEmailChange(long rowId, Sheet WebHookSheet, Sheet sheet,
           Dictionary<string, long> Sheetcolumns, SmartsheetClient smartsheet)
        {
            long RowId = rowId;
            string DesignationValue = "";
            string EmailValue = "";

            Row row = GetRowById(WebHookSheet, RowId);
            Column? designationColumn = WebHookSheet.Columns.FirstOrDefault(c => c.Title == "Designation");
            Column? EmailColumn = WebHookSheet.Columns.FirstOrDefault(c => c.Title == "Email");

            if (designationColumn != null && EmailColumn != null)
            {
                int? designationColumnIndex = designationColumn.Index;
                int? EmailColumnIndex = EmailColumn.Index;
                DesignationValue = row.Cells[(int)designationColumnIndex].Value.ToString();
                EmailValue = row.Cells[(int)EmailColumnIndex].Value.ToString();
                string columnName = "";
                if (DesignationValue == "Sales Head")
                    columnName = "Sales Head Approval";
                else if (DesignationValue == "Medical Affairs Head")
                    columnName = "Medical Affairs Head Approval";

                if (columnName != "")
                {
                    IEnumerable<Row> DataInSheet1 = [];
                    //int? statusColumnIndex = sheet.Columns.Where(y => y.Title == "Post Event Request status").Select(z => z.Index).FirstOrDefault();
                    int? designationStatusIndex = sheet.Columns.Where(y => y.Title == columnName).Select(z => z.Index).FirstOrDefault();

                    DataInSheet1 = sheet.Rows.Where(x =>
                    {
                        //string cellValue = Convert.ToString(x.Cells[(int)statusColumnIndex].Value).ToLower();
                        string designationValue = Convert.ToString(x.Cells[(int)designationStatusIndex].Value).ToLower();
                        return (designationValue != "approved");
                    });
                    List<Row> liRowsToUpdate = new();
                    foreach (Row rowData in DataInSheet1)
                    {
                        Cell[] cellsToUpdate = new Cell[]
                        { new Cell {  ColumnId = Sheetcolumns[DesignationValue], Value = EmailValue  }};

                        row = new Row
                        {
                            Id = rowData.Id,
                            Cells = cellsToUpdate
                        };
                        liRowsToUpdate.Add(row);
                    }
                    if (liRowsToUpdate.Count > 0)
                    {
                        //smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, liRowsToUpdate);
                        ApiCalls.BulkUpdateRows(smartsheet, sheet, liRowsToUpdate);
                        Log.Information("updated " + liRowsToUpdate.Count + " rows");
                    }
                }
                else
                {
                    return "designation not found";
                }
            }


            return "Master sheet updated";
        }

        public static string VendorEmailChange(long rowId, Sheet WebHookSheet, Sheet sheet,
           Dictionary<string, long> Sheetcolumns, SmartsheetClient smartsheet)
        {
            long RowId = rowId;
            string DesignationValue = "";
            string EmailValue = "";

            Row row = GetRowById(WebHookSheet, RowId);
            Column? designationColumn = WebHookSheet.Columns.FirstOrDefault(c => c.Title == "Designation");
            Column? EmailColumn = WebHookSheet.Columns.FirstOrDefault(c => c.Title == "Email");

            if (designationColumn != null && EmailColumn != null)
            {
                int? designationColumnIndex = designationColumn.Index;
                int? EmailColumnIndex = EmailColumn.Index;
                DesignationValue = row.Cells[(int)designationColumnIndex].Value.ToString();
                EmailValue = row.Cells[(int)EmailColumnIndex].Value.ToString();
                string columnName = "";
                if (DesignationValue == "Finance Checker")
                    columnName = "Finance Checker Approval";


                if (columnName != "")
                {
                    IEnumerable<Row> DataInSheet1 = [];
                    // int? statusColumnIndex = sheet.Columns.Where(y => y.Title == "Post Event Request status").Select(z => z.Index).FirstOrDefault();
                    int? designationStatusIndex = sheet.Columns.Where(y => y.Title == columnName).Select(z => z.Index).FirstOrDefault();

                    DataInSheet1 = sheet.Rows.Where(x =>
                    {
                        //string cellValue = Convert.ToString(x.Cells[(int)statusColumnIndex].Value).ToLower();
                        string designationValue = Convert.ToString(x.Cells[(int)designationStatusIndex].Value).ToLower();
                        return (/*(cellValue != "closure approved") &&*/ designationValue != "approved");
                    });
                    List<Row> liRowsToUpdate = new();
                    foreach (Row rowData in DataInSheet1)
                    {
                        Cell[] cellsToUpdate = new Cell[]
                        { new Cell {  ColumnId = Sheetcolumns[DesignationValue], Value = EmailValue  }};

                        row = new Row
                        {
                            Id = rowData.Id,
                            Cells = cellsToUpdate
                        };
                        liRowsToUpdate.Add(row);
                    }
                    if (liRowsToUpdate.Count > 0)
                    {
                        //smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, liRowsToUpdate);
                        ApiCalls.BulkUpdateRows(smartsheet, sheet, liRowsToUpdate);
                        Log.Information("updated " + liRowsToUpdate.Count + " rows");
                    }
                }
                else
                {
                    return "designation not found";
                }
            }


            return "Vendor  sheet updated";
        }

        public static string VendorCodeEmailChange(long rowId, Sheet WebHookSheet, Sheet sheet,
           Dictionary<string, long> Sheetcolumns, SmartsheetClient smartsheet)
        {
            long RowId = rowId;
            string DesignationValue = "";
            string EmailValue = "";

            Row row = GetRowById(WebHookSheet, RowId);
            Column? designationColumn = WebHookSheet.Columns.FirstOrDefault(c => c.Title == "Designation");
            Column? EmailColumn = WebHookSheet.Columns.FirstOrDefault(c => c.Title == "Email");

            if (designationColumn != null && EmailColumn != null)
            {
                int? designationColumnIndex = designationColumn.Index;
                int? EmailColumnIndex = EmailColumn.Index;
                DesignationValue = row.Cells[(int)designationColumnIndex].Value.ToString();
                EmailValue = row.Cells[(int)EmailColumnIndex].Value.ToString();
                string columnName = "";
                if (DesignationValue == "Finance Checker")
                    columnName = "Finance Checker  approval";


                if (columnName != "")
                {
                    IEnumerable<Row> DataInSheet1 = [];
                    //int? statusColumnIndex = sheet.Columns.Where(y => y.Title == "Post Event Request status").Select(z => z.Index).FirstOrDefault();
                    int? designationStatusIndex = sheet.Columns.Where(y => y.Title == columnName).Select(z => z.Index).FirstOrDefault();

                    DataInSheet1 = sheet.Rows.Where(x =>
                    {
                        //string cellValue = Convert.ToString(x.Cells[(int)statusColumnIndex].Value).ToLower();
                        string designationValue = Convert.ToString(x.Cells[(int)designationStatusIndex].Value).ToLower();
                        return (/*(cellValue != "closure approved") &&*/ designationValue != "approved");
                    });
                    List<Row> liRowsToUpdate = new();
                    foreach (Row rowData in DataInSheet1)
                    {
                        Cell[] cellsToUpdate = new Cell[]
                        { new Cell {  ColumnId = Sheetcolumns[DesignationValue], Value = EmailValue  }};

                        row = new Row
                        {
                            Id = rowData.Id,
                            Cells = cellsToUpdate
                        };
                        liRowsToUpdate.Add(row);
                    }
                    if (liRowsToUpdate.Count > 0)
                    {
                        //smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, liRowsToUpdate);
                        ApiCalls.BulkUpdateRows(smartsheet, sheet, liRowsToUpdate);
                        Log.Information("updated " + liRowsToUpdate.Count + " rows");
                    }
                }
                else
                {
                    return "designation not found";
                }
            }


            return "Vendor sheet updated";
        }


        public class Webhook
        {
            public string? smartsheetHookResponse { get; set; }
            public string challenge { get; set; }
            public string webhookId { get; set; }

        }


        public static Row GetRowById(Sheet sheet, long id)
        {
            return sheet.Rows.FirstOrDefault(r => r.Id == id);
        }






        public static string TestprocessSheetEmailChange(long rowId, Sheet WebHookSheet, Sheet sheet,
        Dictionary<string, long> Sheetcolumns, SmartsheetClient smartsheet)
        {
            long RowId = rowId;
            string DesignationValue = "";
            string EmailValue = "";

            Row row = GetRowById(WebHookSheet, RowId);

            Column? designationColumn = WebHookSheet.Columns.FirstOrDefault(c => c.Title == "RBM/BM");
            Column? EmailColumn = WebHookSheet.Columns.FirstOrDefault(c => c.Title == "EmailId");

            if (designationColumn != null && EmailColumn != null)
            {
                int? designationColumnIndex = designationColumn.Index;
                int? EmailColumnIndex = EmailColumn.Index;
                DesignationValue = row.Cells[(int)designationColumnIndex].Value.ToString();//edited email
                EmailValue = row.Cells[(int)EmailColumnIndex].Value.ToString();//initiator email




                IEnumerable<Row> DataInSheet1 = [];
                int? statusColumnIndex = sheet.Columns.Where(y => y.Title == "Event Request Status").Select(z => z.Index).FirstOrDefault();
                int? designationStatusIndex = sheet.Columns.Where(y => y.Title == "RBM/BM").Select(z => z.Index).FirstOrDefault();
                int? InitiatorEmailIndex = sheet.Columns.Where(y => y.Title == "Initiator Email").Select(z => z.Index).FirstOrDefault();

                DataInSheet1 = sheet.Rows.Where(x =>
                {
                    string cellValue = Convert.ToString(x.Cells[(int)statusColumnIndex].Value).ToLower();
                    string designationValue = Convert.ToString(x.Cells[(int)designationStatusIndex].Value).ToLower();
                    string InitiatorValue = Convert.ToString(x.Cells[(int)InitiatorEmailIndex].Value);
                    return (cellValue != "approved" || cellValue != "advance approved") && designationValue != "approved" && InitiatorValue == EmailValue;
                });
                if (DataInSheet1 != null)
                {


                    List<Row> liRowsToUpdate = new();
                    foreach (Row rowData in DataInSheet1)
                    {
                        Cell[] cellsToUpdate = new Cell[]
                        { new Cell {  ColumnId = Sheetcolumns["RBM/BM"], Value = DesignationValue  }};

                        row = new Row
                        {
                            Id = rowData.Id,
                            Cells = cellsToUpdate
                        };
                        liRowsToUpdate.Add(row);
                    }
                    if (liRowsToUpdate.Count > 0)
                    {
                        ApiCalls.BulkUpdateRows(smartsheet, sheet, liRowsToUpdate);

                        Log.Information("updated " + liRowsToUpdate.Count + " rows");
                    }
                }

            }


            return "process sheet updated";
        }

        public static string TestHonorariumEmailChange(long rowId, Sheet WebHookSheet, Sheet sheet,
            Dictionary<string, long> Sheetcolumns, SmartsheetClient smartsheet)
        {
            long RowId = rowId;
            string DesignationValue = "";
            string EmailValue = "";

            Row row = GetRowById(WebHookSheet, RowId);

            Column? designationColumn = WebHookSheet.Columns.FirstOrDefault(c => c.Title == "RBM/BM");
            Column? EmailColumn = WebHookSheet.Columns.FirstOrDefault(c => c.Title == "EmailId");

            if (designationColumn != null && EmailColumn != null)
            {
                int? designationColumnIndex = designationColumn.Index;
                int? EmailColumnIndex = EmailColumn.Index;
                DesignationValue = row.Cells[(int)designationColumnIndex].Value.ToString();//edited email
                EmailValue = row.Cells[(int)EmailColumnIndex].Value.ToString();//initiator email




                IEnumerable<Row> DataInSheet1 = [];
                int? statusColumnIndex = sheet.Columns.Where(y => y.Title == "Honorarium Request Status").Select(z => z.Index).FirstOrDefault();
                int? designationStatusIndex = sheet.Columns.Where(y => y.Title == "RBM/BM").Select(z => z.Index).FirstOrDefault();
                int? InitiatorEmailIndex = sheet.Columns.Where(y => y.Title == "Initiator Email").Select(z => z.Index).FirstOrDefault();

                DataInSheet1 = sheet.Rows.Where(x =>
                {
                    string cellValue = Convert.ToString(x.Cells[(int)statusColumnIndex].Value).ToLower();
                    string designationValue = Convert.ToString(x.Cells[(int)designationStatusIndex].Value).ToLower();
                    string InitiatorValue = Convert.ToString(x.Cells[(int)InitiatorEmailIndex].Value);
                    return (cellValue != "honorarium approved") && designationValue != "approved" && InitiatorValue == EmailValue;
                });
                if (DataInSheet1 != null)
                {


                    List<Row> liRowsToUpdate = new();
                    foreach (Row rowData in DataInSheet1)
                    {
                        Cell[] cellsToUpdate = new Cell[]
                        { new Cell {  ColumnId = Sheetcolumns["RBM/BM"], Value = DesignationValue  }};

                        row = new Row
                        {
                            Id = rowData.Id,
                            Cells = cellsToUpdate
                        };
                        liRowsToUpdate.Add(row);
                    }
                    if (liRowsToUpdate.Count > 0)
                    {
                        ApiCalls.BulkUpdateRows(smartsheet, sheet, liRowsToUpdate);

                        Log.Information("updated " + liRowsToUpdate.Count + " rows");
                    }
                }

            }


            return "Honorarium sheet updated";
        }

        public static string TestEventSettlementEmailChange(long rowId, Sheet WebHookSheet, Sheet sheet,
            Dictionary<string, long> Sheetcolumns, SmartsheetClient smartsheet)
        {
            long RowId = rowId;
            string DesignationValue = "";
            string EmailValue = "";

            Row row = GetRowById(WebHookSheet, RowId);

            Column? designationColumn = WebHookSheet.Columns.FirstOrDefault(c => c.Title == "RBM/BM");
            Column? EmailColumn = WebHookSheet.Columns.FirstOrDefault(c => c.Title == "EmailId");

            if (designationColumn != null && EmailColumn != null)
            {
                int? designationColumnIndex = designationColumn.Index;
                int? EmailColumnIndex = EmailColumn.Index;
                DesignationValue = row.Cells[(int)designationColumnIndex].Value.ToString();//edited email
                EmailValue = row.Cells[(int)EmailColumnIndex].Value.ToString();//initiator email




                IEnumerable<Row> DataInSheet1 = [];
                int? statusColumnIndex = sheet.Columns.Where(y => y.Title == "Post Event Request status").Select(z => z.Index).FirstOrDefault();
                int? designationStatusIndex = sheet.Columns.Where(y => y.Title == "RBM/BM").Select(z => z.Index).FirstOrDefault();
                int? InitiatorEmailIndex = sheet.Columns.Where(y => y.Title == "Initiator Email").Select(z => z.Index).FirstOrDefault();

                DataInSheet1 = sheet.Rows.Where(x =>
                {
                    string cellValue = Convert.ToString(x.Cells[(int)statusColumnIndex].Value).ToLower();
                    string designationValue = Convert.ToString(x.Cells[(int)designationStatusIndex].Value).ToLower();
                    string InitiatorValue = Convert.ToString(x.Cells[(int)InitiatorEmailIndex].Value);
                    return (cellValue != "closure approved") && designationValue != "approved" && InitiatorValue == EmailValue;
                });
                if (DataInSheet1 != null)
                {


                    List<Row> liRowsToUpdate = new();
                    foreach (Row rowData in DataInSheet1)
                    {
                        Cell[] cellsToUpdate = new Cell[]
                        { new Cell {  ColumnId = Sheetcolumns["RBM/BM"], Value = DesignationValue  }};

                        row = new Row
                        {
                            Id = rowData.Id,
                            Cells = cellsToUpdate
                        };
                        liRowsToUpdate.Add(row);
                    }
                    if (liRowsToUpdate.Count > 0)
                    {
                        ApiCalls.BulkUpdateRows(smartsheet, sheet, liRowsToUpdate);

                        Log.Information("updated " + liRowsToUpdate.Count + " rows");
                    }
                }

            }


            return "Event Settlement sheet updated";
        }

    }

}
