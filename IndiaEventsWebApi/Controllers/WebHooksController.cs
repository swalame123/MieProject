using IndiaEventsWebApi.Helper;
using IndiaEventsWebApi.Models.RequestSheets;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Serilog;
using NLog.Fluent;
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
                                    if (status != null && status == "Approved")
                                    {
                                        int timeInterval = 250000;
                                        await Task.Delay(timeInterval);
                                        if (meetingType != null)
                                        {
                                            if (meetingType.ToString() == "Other |")
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
                                else if (status != null && status == "Approved")
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



        //private void GenerateSummaryPDF(string EventID, long rowId)
        //{
        //    try
        //    {

        //        var EventCode = "";
        //        var EventName = "";
        //        var EventDate = "";
        //        var EventVenue = "";
        //        List<string> Speakers = new List<string>();




        //        Column SpecialityColumn = sheet1.Columns.FirstOrDefault(column => string.Equals(column.Title, "EventId/EventRequestId", StringComparison.OrdinalIgnoreCase));
        //        Column targetColumn1 = sheet1.Columns.FirstOrDefault(column => string.Equals(column.Title, "Event Topic", StringComparison.OrdinalIgnoreCase));
        //        Column targetColumn2 = sheet1.Columns.FirstOrDefault(column => string.Equals(column.Title, "EventDate", StringComparison.OrdinalIgnoreCase));
        //        Column targetColumn3 = sheet1.Columns.FirstOrDefault(column => string.Equals(column.Title, "VenueName", StringComparison.OrdinalIgnoreCase));

        //        if (SpecialityColumn != null)
        //        {
        //            Row targetRow = sheet1.Rows
        //             .FirstOrDefault(row => row.Cells?.Any(cell => cell.ColumnId == SpecialityColumn.Id && cell.Value?.ToString() == EventID) == true);

        //            if (targetRow != null)
        //            {

        //                EventCode = targetRow.Cells.FirstOrDefault(cell => cell.ColumnId == SpecialityColumn.Id)?.Value?.ToString();
        //                EventName = targetRow.Cells.FirstOrDefault(cell => cell.ColumnId == targetColumn1.Id)?.Value?.ToString();
        //                EventDate = targetRow.Cells.FirstOrDefault(cell => cell.ColumnId == targetColumn2.Id)?.Value?.ToString();
        //                EventVenue = targetRow.Cells.FirstOrDefault(cell => cell.ColumnId == targetColumn3.Id)?.Value?.ToString();
        //            }
        //        }

        //        List<string> requiredColumns = new List<string> { "HCPName", "MISCode", "Speciality", "HCP Type" };
        //        List<string> MenariniColumns = new List<string> { "HCPName", "MISCode", "Speciality" };

        //        List<Column> selectedColumns = sheet_SpeakerCode.Columns.Where(column => requiredColumns.Contains(column.Title, StringComparer.OrdinalIgnoreCase)).ToList();
        //        List<Column> selectedMenariniColumns = sheet.Columns.Where(column => requiredColumns.Contains(column.Title, StringComparer.OrdinalIgnoreCase)).ToList();

        //        DataTable dtMai = new DataTable();
        //        dtMai.Columns.Add("S.No", typeof(int));
        //        foreach (Column column in selectedColumns)
        //        {
        //            dtMai.Columns.Add(column.Title);
        //        }
        //        dtMai.Columns.Add("Sign");

        //        DataTable MenariniTable = new DataTable();
        //        MenariniTable.Columns.Add("S.No", typeof(int));
        //        foreach (Column column in selectedMenariniColumns)
        //        {
        //            MenariniTable.Columns.Add(column.Title);
        //        }
        //        MenariniTable.Columns.Add("Sign");

        //        int Sr_No = 1;
        //        int m_no = 1;
        //        foreach (Row row in sheet_SpeakerCode.Rows)
        //        {
        //            string eventId = row.Cells.FirstOrDefault(cell => sheet_SpeakerCode.Columns.FirstOrDefault(c => c.Id == cell.ColumnId)?.Title == "EventId/EventRequestId")?.DisplayValue;
        //            if (!string.IsNullOrEmpty(eventId) && eventId.Equals(EventID, StringComparison.OrdinalIgnoreCase))
        //            {
        //                DataRow newRow = dtMai.NewRow();
        //                newRow["S.No"] = Sr_No;
        //                foreach (Cell cell in row.Cells)
        //                {
        //                    string columnName = sheet_SpeakerCode.Columns.FirstOrDefault(c => c.Id == cell.ColumnId)?.Title;
        //                    if (requiredColumns.Contains(columnName, StringComparer.OrdinalIgnoreCase))
        //                    {
        //                        if (columnName == "HCPName")
        //                        {
        //                            var val = cell.DisplayValue;
        //                            Speakers.Add(val);
        //                        }
        //                        newRow[columnName] = cell.DisplayValue;
        //                    }
        //                }
        //                dtMai.Rows.Add(newRow);
        //                Sr_No++;
        //            }
        //        }

        //        foreach (Row row in sheet.Rows)
        //        {
        //            string eventId = row.Cells.FirstOrDefault(cell => sheet.Columns.FirstOrDefault(c => c.Id == cell.ColumnId)?.Title == "EventId/EventRequestId")?.DisplayValue;
        //            string InviteeSource = row.Cells.FirstOrDefault(cell => sheet.Columns.FirstOrDefault(c => c.Id == cell.ColumnId)?.Title == "Invitee Source")?.DisplayValue;
        //            if (!string.IsNullOrEmpty(eventId) && eventId.Equals(EventID, StringComparison.OrdinalIgnoreCase) && !InviteeSource.Equals("Menarini Employees", StringComparison.OrdinalIgnoreCase))
        //            {
        //                DataRow newRow = dtMai.NewRow();
        //                newRow["S.No"] = Sr_No;
        //                foreach (Cell cell in row.Cells)
        //                {
        //                    string columnName = sheet.Columns
        //                        .FirstOrDefault(c => c.Id == cell.ColumnId)?.Title;
        //                    if (requiredColumns.Contains(columnName, StringComparer.OrdinalIgnoreCase))
        //                    {
        //                        newRow[columnName] = cell.DisplayValue;
        //                    }
        //                }
        //                dtMai.Rows.Add(newRow);
        //                Sr_No++;
        //            }
        //            else if (!string.IsNullOrEmpty(eventId) && eventId.Equals(EventID, StringComparison.OrdinalIgnoreCase) && InviteeSource.Equals("Menarini Employees", StringComparison.OrdinalIgnoreCase))
        //            {
        //                DataRow newRow = MenariniTable.NewRow();
        //                newRow["S.No"] = m_no;
        //                foreach (Cell cell in row.Cells)
        //                {
        //                    string columnName = sheet.Columns
        //                        .FirstOrDefault(c => c.Id == cell.ColumnId)?.Title;
        //                    if (requiredColumns.Contains(columnName, StringComparer.OrdinalIgnoreCase))
        //                    {
        //                        newRow[columnName] = cell.DisplayValue;
        //                    }
        //                }
        //                MenariniTable.Rows.Add(newRow);
        //                m_no++;
        //            }






        //        }
        //        string resultString = string.Join(", ", Speakers);

        //        byte[] fileBytes = SheetHelper.exportAttendencepdfnew(dtMai,MenariniTable, EventCode, EventName, EventDate, EventVenue, resultString);
        //        string filename = "Attendance Sheet_" + EventID + ".pdf";
        //        var folderName = Path.Combine("Resources", "Images");
        //        var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
        //        if (!Directory.Exists(pathToSave))
        //        {
        //            Directory.CreateDirectory(pathToSave);
        //        }
        //        string fileType = SheetHelper.GetFileType(fileBytes);
        //        string filePath = Path.Combine(pathToSave, filename);
        //        System.IO.File.WriteAllBytes(filePath, fileBytes);
        //        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile((long)processSheetData.Id, rowId, filePath, "application/msword");
        //        if (System.IO.File.Exists(filePath))
        //        {
        //            System.IO.File.Delete(filePath);
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        //return BadRequest(ex.Message);
        //    }
        //}

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
                                url = file.Url;
                                name = file.Name;
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

                        EventCode = targetRow.Cells.FirstOrDefault(cell => cell.ColumnId == SpecialityColumn.Id)?.Value?.ToString();
                        EventName = targetRow.Cells.FirstOrDefault(cell => cell.ColumnId == targetColumn1.Id)?.Value?.ToString();
                        EventDate = targetRow.Cells.FirstOrDefault(cell => cell.ColumnId == targetColumn2.Id)?.Value?.ToString();
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



        private byte[] exportpdf(DataTable dtEmployee, string EventCode, string EventName, string EventDate, string EventVenue, string speakers)
        {
            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            iTextSharp.text.Rectangle rec = new iTextSharp.text.Rectangle(PageSize.A4);
            rec.BackgroundColor = new BaseColor(System.Drawing.Color.Olive);
            Document doc = new Document(rec);
            doc.SetPageSize(iTextSharp.text.PageSize.A4);
            PdfWriter writer = PdfWriter.GetInstance(doc, ms);
            doc.Open();
            BaseFont bfntHead = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            iTextSharp.text.Font fntHead = new iTextSharp.text.Font(bfntHead, 16, 1, iTextSharp.text.BaseColor.BLUE);
            Paragraph prgHeading = new Paragraph();
            prgHeading.Alignment = Element.ALIGN_LEFT;
            prgHeading.Add(new Chunk("Attendance Sheet".ToUpper(), fntHead));
            doc.Add(prgHeading);
            Paragraph prgGeneratedBY = new Paragraph();
            BaseFont btnAuthor = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            iTextSharp.text.Font fntAuthor = new iTextSharp.text.Font(btnAuthor, 8, 2, iTextSharp.text.BaseColor.BLUE);
            prgGeneratedBY.Alignment = Element.ALIGN_RIGHT;
            doc.Add(prgGeneratedBY);
            Paragraph p = new Paragraph(new Chunk(new iTextSharp.text.pdf.draw.LineSeparator(0.0F, 100.0F, iTextSharp.text.BaseColor.BLACK, Element.ALIGN_LEFT, 1)));
            doc.Add(p);
            doc.Add(new Chunk("\n", fntHead));

            Paragraph pBody = new Paragraph();
            pBody.Add(new Chunk("Event Code:" + EventCode));
            pBody.Add(new Chunk("\nEvent Name:" + EventName));
            pBody.Add(new Chunk("\nEvent Date:" + EventDate));
            pBody.Add(new Chunk("\nEvent Venue:" + EventVenue));
            pBody.Add(new Chunk("\n\nSpeakers: "));

            //foreach(DataRow row in dtMai.Rows)
            //{
            //    string hcpName = row["HCPName"].ToString();
            //    pBody.Add(new Chunk(" " + hcpName));
            //}

            string hcpNames = speakers;//string.Join(", ", dtMai.AsEnumerable().Select(row => row["HCPName"].ToString()));
            pBody.Add(new Chunk(" " + hcpNames));
            doc.Add(pBody);
            doc.Add(new Paragraph("\n "));
            PdfPTable table = new PdfPTable(dtEmployee.Columns.Count);
            table.WidthPercentage = 100;
            float[] columnWidths = Enumerable.Range(0, dtEmployee.Columns.Count).Select(i => i == dtEmployee.Columns.IndexOf("HCPName") ? 2f : 1f).ToArray(); /*Count).Select(i => 1f).ToArray();*/
            table.SetWidths(columnWidths);

            for (int i = 0; i < dtEmployee.Columns.Count; i++)
            {
                string cellText = dtEmployee.Columns[i].ColumnName;
                PdfPCell cell = new PdfPCell();
                cell.Phrase = new Phrase(cellText, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 8, 1, new BaseColor(System.Drawing.ColorTranslator.FromHtml("#000000"))));
                cell.BackgroundColor = new BaseColor(System.Drawing.ColorTranslator.FromHtml("#C8C8C8"));
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.PaddingBottom = 5;
                table.AddCell(cell);
            }
            for (int i = 0; i < dtEmployee.Rows.Count; i++)
            {
                for (int j = 0; j < dtEmployee.Columns.Count; j++)
                {
                    table.AddCell(dtEmployee.Rows[i][j].ToString());
                }
            }
            doc.Add(table);
            doc.Close();
            byte[] result = ms.ToArray();
            return result;
        }

    }

}
//string BaseHref => $"{HttpContext.Request.Scheme}://{HttpContext.Request.Host}";


//[HttpPost(Name = "WebHook")]
//public async Task<IActionResult> PostData()
//{

//    try
//    {

//        Dictionary<string, string> requestHeaders = ExtractRequestHeaders();
//        string rawContent = ExtractRawContent();


//        var webhookPayload = JsonConvert.DeserializeObject<Root>(rawContent);


//        var attachedFile = Attachementfile(webhookPayload);


//        var challenge = requestHeaders.GetValueOrDefault("challenge");


//        LogWebhookRequest(requestHeaders);


//        return Ok(new Webhook { smartsheetHookResponse = challenge });
//    }
//    catch (Exception ex)
//    {

//        Log.Error($"Error occurred in Webhook API controller PostData method: {ex.Message} at {DateTime.Now}");


//        return BadRequest(ex.StackTrace);
//    }

//}




//private Dictionary<string, string> ExtractRequestHeaders()
//{
//    Dictionary<string, string> requestHeaders = new Dictionary<string, string>();

//    foreach (var header in Request.Headers)
//    {
//        requestHeaders.Add(header.Key, header.Value.ToString());
//    }

//    return requestHeaders;
//}

//private string ExtractRawContent()
//{
//    using (var reader = new StreamReader(Request.Body, encoding: Encoding.UTF8, detectEncodingFromByteOrderMarks: false))
//    {
//        return reader.ReadToEndAsync().Result;
//    }
//}

//private void LogWebhookRequest(Dictionary<string, string> requestHeaders)
//{

//    Log.Info(string.Join(";", requestHeaders.Select(x => x.Key + "=" + x.Value).ToArray()));

//}

//private string Attachementfile(Root webhookPayload)
//{


//    return "Path to attached file";
//}

//private class Root
//{
//}
//public class Webhook
//{
//    public string smartsheetHookResponse { get; set; }
//    public long? Id { get; internal set; }
//    public string Name { get; internal set; }
//}









































//    [HttpPost("WebhookHandler")]
//    public IActionResult WebhookHandler([FromBody] WebhookPayload payload)
//    {
//        try
//        {
//            // Validate payload, check event type, etc.
//            if (payload.EventType == "UPDATE" && payload.ResourceType == "ROW")
//            {
//                // Check if the updated column is "Status" and the new value is "completed"
//                if (payload.Change != null &&
//                    payload.Change.ColumnId == "your_status_column_id" &&
//                    payload.Change.NewValue == "completed")
//                {
//                    // Extract necessary information from the payload, e.g., EventID
//                    string EventID = payload.RowId.ToString(); // Replace with the actual way to get EventID

//                    // Call your existing logic to generate PDF
//                    GenerateSummaryPDF(EventID);
//                }
//            }

//            return Ok();
//        }
//        catch (Exception ex)
//        {
//            // Log and handle exceptions
//            // You might want to return an error response here if needed
//            return BadRequest(ex.Message);
//        }
//    }
//}

//// This class represents the structure of the Smartsheet webhook payload.
//// Please replace it with the actual structure provided by Smartsheet.
//public class WebhookPayload
//{
//    public string EventType { get; set; }
//    public string ResourceType { get; set; }
//    public long? RowId { get; set; }
//    public Change Change { get; set; }
//}

//public class Change
//{
//    public string ColumnId { get; set; }
//    public string NewValue { get; set; }

