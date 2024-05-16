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
using static Org.BouncyCastle.Bcpg.Attr.ImageAttrib;

namespace IndiaEventsWebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class WebHooksController : ControllerBase
    {

        private readonly string accessToken;
        private readonly IConfiguration configuration;
        private readonly SmartsheetClient smartsheet;
        private readonly Sheet processSheetData;
        private readonly Sheet sheet1;
        private readonly Sheet sheet;
        private readonly Sheet sheet_SpeakerCode;
        public WebHooksController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
            smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            string processSheet = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;
            string sheetId_SpeakerCode = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
            string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
            string sheetId1 = configuration.GetSection("SmartsheetSettings:Class1").Value;
            processSheetData = SheetHelper.GetSheetById(smartsheet, processSheet);
            sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
            sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
            sheet_SpeakerCode = SheetHelper.GetSheetById(smartsheet, sheetId_SpeakerCode);

        }

        [HttpPost(Name = "WebHook")]
        public async Task<IActionResult> PostData()
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
                Serilog.Log.Information(string.Join(";", requestHeaders.Select(x => x.Key + "=" + x.Value).ToArray()));


                var RequestWebhook = JsonConvert.DeserializeObject<Root>(rawContent);
                Attachementfile(RequestWebhook);

                var challenge = requestHeaders.Where(x => x.Key == "challenge").Select(x => x.Value).FirstOrDefault();

                return Ok(new Webhook { smartsheetHookResponse = RequestWebhook.challenge });
                //return Ok();
            }
            catch (Exception ex)
            {
                Serilog.Log.Error($"Error occured on Webhook apicontroller PostData method {ex.Message} at {DateTime.Now}");
                Serilog.Log.Error(ex.StackTrace);
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



                var RequestWebhook = JsonConvert.DeserializeObject<Root>(rawContent);
                AgreementsTrigger(RequestWebhook);

                var challenge = requestHeaders.Where(x => x.Key == "challenge").Select(x => x.Value).FirstOrDefault();

                return Ok(new Webhook { smartsheetHookResponse = RequestWebhook.challenge });
                //return Ok();
            }
            catch (Exception ex)
            {
                Serilog.Log.Error($"Error occured on Webhook apicontroller PostData method {ex.Message} at {DateTime.Now}");
                Serilog.Log.Error(ex.StackTrace);
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


                var RequestWebhook = JsonConvert.DeserializeObject<Root>(rawContent);
                ApprovalCheckBox(RequestWebhook);

                var challenge = requestHeaders.Where(x => x.Key == "challenge").Select(x => x.Value).FirstOrDefault();

                return Ok(new Webhook { smartsheetHookResponse = RequestWebhook.challenge });
                //return Ok();
            }
            catch (Exception ex)
            {
                Serilog.Log.Error($"Error occured on Webhook apicontroller PostData method {ex.Message} at {DateTime.Now}");
                Serilog.Log.Error(ex.StackTrace);
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
                using (var reader = new StreamReader(Request.Body, encoding: Encoding.UTF8, detectEncodingFromByteOrderMarks: false))
                {
                    rawContent = await reader.ReadToEndAsync();
                }
                requestHeaders.Add("Body", rawContent);


                var RequestWebhook = JsonConvert.DeserializeObject<Root>(rawContent);
                //EventSettlementApproval(RequestWebhook);
                EventSettlementDeviationApproval(RequestWebhook);
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
                Dictionary<string, string> requestHeaders = new Dictionary<string, string>();
                string rawContent = string.Empty;
                using (var reader = new StreamReader(Request.Body, encoding: Encoding.UTF8, detectEncodingFromByteOrderMarks: false))
                {
                    rawContent = await reader.ReadToEndAsync();
                }
                requestHeaders.Add("Body", rawContent);


                var RequestWebhook = JsonConvert.DeserializeObject<Root>(rawContent);
                PreEventApproval(RequestWebhook);

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
        private async void PreEventApproval(Root RequestWebhook)
        {
            //try
            //{
            //    string processSheet = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;

            //    //var TestingId = "6831673324818308";
            //    Sheet TestingSheetData = SheetHelper.GetSheetById(smartsheet, processSheet);

            //    if (RequestWebhook != null && RequestWebhook.events != null)
            //    {
            //        foreach (var WebHookEvent in RequestWebhook.events)
            //        {
            //            if (WebHookEvent.eventType.ToLower() == "updated" || WebHookEvent.eventType.ToLower() == "created")
            //            {
            //                Row targetRowId = TestingSheetData.Rows.FirstOrDefault(row => row.Id == WebHookEvent.rowId);
            //                if (targetRowId != null)
            //                {
            //                    string? status = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == 6770461806382980)?.Value?.ToString();
            //                    if (status.ToLower() == "approved")
            //                    {
            //                        long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(TestingSheetData, "Is All Deviations Approved?");
            //                        Cell cellToUpdateB = new() { ColumnId = honorariumSubmittedColumnId, Value = "Yes" };
            //                        Row updateRow = new() { Id = targetRowId.Id, Cells = new Cell[] { cellToUpdateB } };
            //                        Cell? cellToUpdate = targetRowId.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
            //                        if (cellToUpdate != null) { cellToUpdate.Value = "Yes"; }

            //                        smartsheet.SheetResources.RowResources.UpdateRows(TestingSheetData.Id.Value, new Row[] { updateRow });
            //                    }
            //                }
            //            }
            //        }
            //    }
            //}
            try
            {
                string processSheet = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;

                //var TestingId = "6831673324818308";
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
                            //var columnValue = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == processIdColumn1.Id)?.Value.ToString();


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

                //var TestingId = "6831673324818308";
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










                                //if ((status1.ToLower() == "approved" || status1.ToLower() == "null") &&
                                // (status2.ToLower() == "approved" || status2.ToLower() == "null") &&
                                // (status3.ToLower() == "submitted" || status3.ToLower() == "null") &&
                                // (status4.ToLower() == "approved" || status4.ToLower() == "null") &&
                                // (status5.ToLower() == "approved" || status5.ToLower() == "null") &&
                                // (status6.ToLower() == "approved" || status6.ToLower() == "null") &&
                                // (status7.ToLower() == "approved" || status7.ToLower() == "null") &&
                                // (status8.ToLower() == "approved" || status8.ToLower() == "null") &&
                                // (status9.ToLower() == "approved" || status9.ToLower() == "null"))
                                //{

                                //    long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(TestingSheetData, "Is All Deviations Approved?");
                                //    Cell cellToUpdateB = new() { ColumnId = honorariumSubmittedColumnId, Value = "Sales Deviations Approved" };
                                //    Row updateRow = new() { Id = targetRowId.Id, Cells = new Cell[] { cellToUpdateB } };
                                //    Cell? cellToUpdate = targetRowId.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                                //    if (cellToUpdate != null) { cellToUpdate.Value = "Sales Deviations Approved"; }

                                //    smartsheet.SheetResources.RowResources.UpdateRows(TestingSheetData.Id.Value, new Row[] { updateRow });
                                //}
                                //else if (status3.ToLower() == "approved" || status3.ToLower() == "null")
                                ////TriggerStatus.ToLower() != "30 days deviation pending" ||
                                ////TriggerStatus.ToLower() != "Less than 5 invitees pending" ||

                                ////TriggerStatus.ToLower() != "cin pending" ||
                                ////TriggerStatus.ToLower() != "cis pending" ||
                                ////TriggerStatus.ToLower() != "cit pending" ||
                                ////TriggerStatus.ToLower() != "anc pending" ||
                                ////TriggerStatus.ToLower() != "snc is pending" ||
                                ////TriggerStatus.ToLower() != "od pending")
                                //{
                                //    long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(TestingSheetData, "Is All Deviations Approved?");
                                //    Cell cellToUpdateB = new() { ColumnId = honorariumSubmittedColumnId, Value = "Finance Deviations Approved" };
                                //    Row updateRow = new() { Id = targetRowId.Id, Cells = new Cell[] { cellToUpdateB } };
                                //    Cell? cellToUpdate = targetRowId.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                                //    if (cellToUpdate != null) { cellToUpdate.Value = "Finance Deviations Approved"; }

                                //    smartsheet.SheetResources.RowResources.UpdateRows(TestingSheetData.Id.Value, new Row[] { updateRow });
                                //}
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
                Serilog.Log.Error($"Error occured on Webhook apicontroller Attachementfile method {ex.Message} at {DateTime.Now}");
                Serilog.Log.Error(ex.StackTrace);
            }

        }

        private async void AgreementsTrigger(Root RequestWebhook)
        {
            try
            {

                if (RequestWebhook != null && RequestWebhook.events != null)
                {
                    foreach (var WebHookEvent in RequestWebhook.events)
                    {

                        if (WebHookEvent.eventType.ToLower() == "updated" /*|| WebHookEvent.eventType.ToLower() == "created"*/)
                        {
                            //var DataInSheet = smartsheet.SheetResources.GetSheet(sheet_SpeakerCode.Id.Value, null, null, new List<long> { WebHookEvent.rowId }, null, null, null, null, null, null).Rows;



                            Row targetRowId = sheet_SpeakerCode.Rows.FirstOrDefault(row => row.Id == WebHookEvent.rowId);


                            long ColumnId = SheetHelper.GetColumnIdByName(sheet_SpeakerCode, "Agreement Trigger");
                            Cell updatedCell = new Cell
                            {
                                ColumnId = ColumnId,
                                Value = "Yes"
                            };
                            Row updatedRow = new Row
                            {
                                Id = targetRowId.Id,
                                Cells = new List<Cell> { updatedCell }
                            };



                            smartsheet.SheetResources.RowResources.UpdateRows(sheet_SpeakerCode.Id.Value, new Row[] { updatedRow });





                        }
                    }
                }
            }

            catch (Exception ex)
            {
                Serilog.Log.Error($"Error occured on Webhook apicontroller Attachementfile method {ex.Message} at {DateTime.Now}");
                Serilog.Log.Error(ex.StackTrace);
            }

        }

        private async void Attachementfile(Root RequestWebhook)
        {
            try
            {

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
                            Column processIdColumn4 = processSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "EventType", StringComparison.OrdinalIgnoreCase));


                            Row targetRowId = processSheetData.Rows.FirstOrDefault(row => row.Id == WebHookEvent.rowId);



                            if (processIdColumn1 != null && processIdColumn2 != null)
                            {
                                var columnValue = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == processIdColumn1.Id)?.Value.ToString();
                                var status = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == processIdColumn2.Id)?.Value.ToString();
                                var meetingType = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == processIdColumn3.Id)?.Value;
                                var EventType = targetRowId.Cells.FirstOrDefault(cell => cell.ColumnId == processIdColumn4.Id)?.Value.ToString();
                                if (EventType == "Class I" || EventType == "Webinar")
                                {
                                    if (status != null && (status == "Approved" || status == "Waiting for Finance Treasury Approval"))

                                    {
                                        int timeInterval = 250000;
                                        await Task.Delay(timeInterval);
                                        if (meetingType != null)
                                        {
                                            if (meetingType.ToString() == "Other | ")
                                            {

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
                                else if (status != null && status == "Approved" || status == "Waiting for Finance Treasury Approval")
                                {
                                    int timeInterval = 250000;
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
                Serilog.Log.Error($"Error occured on Webhook apicontroller Attachementfile method {ex.Message} at {DateTime.Now}");
                Serilog.Log.Error(ex.StackTrace);
            }

        }

        private void moveAttachments(string EventID, long rowId)
        {
            List<Attachment> attachments = new List<Attachment>();


            foreach (Row row in sheet_SpeakerCode.Rows)
            {
                Cell matchingCell = row.Cells.FirstOrDefault(cell => cell.DisplayValue == EventID);

                if (matchingCell != null && matchingCell.Value != null)
                {
                    var Id = (long)row.Id;
                    string eventId = matchingCell.Value.ToString();
                    if (!string.IsNullOrEmpty(eventId) && eventId.Equals(EventID, StringComparison.OrdinalIgnoreCase))
                    {
                        var a = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments((long)sheet_SpeakerCode.Id, Id, null);
                        var url = "";
                        var name = "";
                        foreach (var x in a.Data)
                        {
                            if (x != null)
                            {
                                var AID = (long)x.Id;
                                var file = smartsheet.SheetResources.AttachmentResources.GetAttachment((long)sheet_SpeakerCode.Id, AID);
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
                                    var f = Path.Combine("Resources", "Images");
                                    var ps = Path.Combine(Directory.GetCurrentDirectory(), f);
                                    if (!Directory.Exists(ps))
                                    {
                                        Directory.CreateDirectory(ps);
                                    }
                                    string ft = SheetHelper.GetFileType(xy);
                                    string fileName = name;
                                    string fp = Path.Combine(ps, fileName);

                                    System.IO.File.WriteAllBytes(fp, xy);
                                    string type = SheetHelper.GetContentType(ft);
                                    var z = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile((long)processSheetData.Id, rowId, fp, "application/msword");

                                }
                                url = "";
                                var bs64 = "";
                            }
                        }
                    }
                }
            }
        }

        private void GenerateSummaryPDF(string EventID, long rowId)
        {
            try
            {

                var EventCode = "";
                var EventName = "";
                var EventDate = "";
                var EventVenue = "";
                DateTime parsedDate;
                List<string> Speakers = new List<string>();




                Column SpecialityColumn = sheet1.Columns.FirstOrDefault(column => string.Equals(column.Title, "EventId/EventRequestId", StringComparison.OrdinalIgnoreCase));
                Column targetColumn1 = sheet1.Columns.FirstOrDefault(column => string.Equals(column.Title, "Event Topic", StringComparison.OrdinalIgnoreCase));
                Column targetColumn2 = sheet1.Columns.FirstOrDefault(column => string.Equals(column.Title, "EventDate", StringComparison.OrdinalIgnoreCase));
                Column targetColumn3 = sheet1.Columns.FirstOrDefault(column => string.Equals(column.Title, "VenueName", StringComparison.OrdinalIgnoreCase));

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
                            string columnName = sheet.Columns
                                .FirstOrDefault(c => c.Id == cell.ColumnId)?.Title;
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
                var folderName = Path.Combine("Resources", "Images");
                var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                if (!Directory.Exists(pathToSave))
                {
                    Directory.CreateDirectory(pathToSave);
                }
                string fileType = SheetHelper.GetFileType(fileBytes);
                string filePath = Path.Combine(pathToSave, filename);
                System.IO.File.WriteAllBytes(filePath, fileBytes);
                var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile((long)processSheetData.Id, rowId, filePath, "application/msword");
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

        public class Webhook
        {
            public string? smartsheetHookResponse { get; set; }
            public string challenge { get; set; }
            public string webhookId { get; set; }

        }



    }

}
