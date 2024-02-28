using IndiaEventsWebApi.Helper;
using IndiaEventsWebApi.Models.RequestSheets;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using NLog.Fluent;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Data;
using System.Text;

namespace IndiaEventsWebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class WebHooksController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;
        private readonly SmartsheetClient _smartsheetClient;
        private readonly ILogger<WebHooksController> _logger;
        public WebHooksController(IConfiguration configuration, ILogger<WebHooksController> logger, SmartsheetClient smartsheetClient)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
            _smartsheetClient = smartsheetClient;
            _logger = logger;
        }

        [HttpGet("GenerateSummaryPDF")]
        public void GenerateSummaryPDF(string EventID)
        {
            try
            {

                var EventCode = "";
                var EventName = "";
                var EventDate = "";
                var EventVenue = "";

                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

                string sheetId_SpeakerCode = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
                long.TryParse(sheetId_SpeakerCode, out long parsedSheetId_SpeakerCode);
                Sheet sheet_SpeakerCode = smartsheet.SheetResources.GetSheet(parsedSheetId_SpeakerCode, null, null, null, null, null, null, null);

                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
                long.TryParse(sheetId, out long parsedSheetId);
                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);

                string sheetId1 = configuration.GetSection("SmartsheetSettings:Class1").Value;
                long.TryParse(sheetId1, out long parsedSheetId1);
                Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId1, null, null, null, null, null, null, null);

                string processSheet = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;
                long.TryParse(processSheet, out long parsedProcessSheet);
                Sheet processSheetData = smartsheet.SheetResources.GetSheet(parsedProcessSheet, null, null, null, null, null, null, null);

                long rowId = 0;

                Column processIdColumn = processSheetData.Columns.FirstOrDefault(column => string.Equals(column.Title, "EventId/EventRequestId", StringComparison.OrdinalIgnoreCase));
                if (processIdColumn != null)
                {
                    // Find all rows with the specified speciality
                    List<Row> targetRows = processSheetData.Rows
                        .Where(row => row.Cells.Any(cell => cell.ColumnId == processIdColumn.Id && cell.Value.ToString() == EventID))
                        .ToList();
                    if (targetRows.Any())
                    {
                        var rowIds = targetRows.Select(row => row.Id).ToList();
                        rowId = (long)rowIds[0];
                    }
                }
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
                List<Column> selectedColumns = sheet_SpeakerCode.Columns
                    .Where(column => requiredColumns.Contains(column.Title, StringComparer.OrdinalIgnoreCase)).ToList();
                DataTable dtMai = new DataTable();
                dtMai.Columns.Add("S.No", typeof(int));
                foreach (Column column in selectedColumns)
                {
                    dtMai.Columns.Add(column.Title);
                }
                dtMai.Columns.Add("Sign");
                int Sr_No = 1;
                foreach (Row row in sheet_SpeakerCode.Rows)
                {
                    string eventId = row.Cells
                        .FirstOrDefault(cell => sheet_SpeakerCode.Columns.FirstOrDefault(c => c.Id == cell.ColumnId)?.Title == "EventId/EventRequestId")?.DisplayValue;
                    if (!string.IsNullOrEmpty(eventId) && eventId.Equals(EventID, StringComparison.OrdinalIgnoreCase))
                    {
                        DataRow newRow = dtMai.NewRow();
                        newRow["S.No"] = Sr_No;
                        foreach (Cell cell in row.Cells)
                        {
                            string columnName = sheet_SpeakerCode.Columns.FirstOrDefault(c => c.Id == cell.ColumnId)?.Title;
                            if (requiredColumns.Contains(columnName, StringComparer.OrdinalIgnoreCase))
                            {
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
                    if (!string.IsNullOrEmpty(eventId) && eventId.Equals(EventID, StringComparison.OrdinalIgnoreCase))
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
                }

                byte[] fileBytes = SheetHelper.exportpdf(dtMai, EventCode, EventName, EventDate, EventVenue, dtMai);
                string filename = "Sample_PDF_" + EventID + ".pdf";
                var folderName = Path.Combine("Resources", "Images");
                var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                if (!Directory.Exists(pathToSave))
                {
                    Directory.CreateDirectory(pathToSave);
                }
                string fileType = SheetHelper.GetFileType(fileBytes);
                string filePath = Path.Combine(pathToSave, filename);
                System.IO.File.WriteAllBytes(filePath, fileBytes);
                var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(parsedProcessSheet, rowId, filePath, "application/msword");
                if (System.IO.File.Exists(filePath))
                {
                    System.IO.File.Delete(filePath);
                }
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
                            var a = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(parsedSheetId_SpeakerCode, Id, null);
                            var url = "";
                            foreach (var x in a.Data)
                            {
                                if (x != null)
                                {
                                    var AID = (long)x.Id;
                                    var file = smartsheet.SheetResources.AttachmentResources.GetAttachment(parsedSheetId_SpeakerCode, AID);
                                    url = file.Url;
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
                                        string fileName = eventId + "-" + x + " AttachedFile." + ft;
                                        string fp = Path.Combine(ps, fileName);
                                        var addedRow = rowId;
                                        System.IO.File.WriteAllBytes(fp, xy);
                                        string type = SheetHelper.GetContentType(ft);
                                        var z = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(parsedProcessSheet, addedRow, fp, "application/msword");
                                    }
                                    var bs64 = "";
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //return BadRequest(ex.Message);
            }
        }
        //private string GetContentType(string fileExtension)
        //{
        //    switch (fileExtension.ToLower())
        //    {
        //        case "jpg":
        //        case "jpeg":
        //            return "image/jpeg";
        //        case "pdf":
        //            return "application/pdf";
        //        case "gif":
        //            return "image/gif";
        //        case "png":
        //            return "image/png";
        //        case "webp":
        //            return "image/webp";
        //        case "doc":
        //            return "application/msword";
        //        case "docx":
        //            return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
        //        default:
        //            return "application/octet-stream";
        //    }
        //}

        //private string GetFileType(byte[] bytes)
        //{

        //    if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xD8)
        //    {
        //        return "jpg";
        //    }
        //    else if (bytes.Length >= 4 && Encoding.UTF8.GetString(bytes, 0, 4) == "%PDF")
        //    {
        //        return "pdf";
        //    }
        //    else if (bytes.Length >= 3 && Encoding.UTF8.GetString(bytes, 0, 3) == "GIF")
        //    {
        //        return "gif";
        //    }
        //    else if (bytes.Length >= 8 && Encoding.UTF8.GetString(bytes, 0, 8) == "PNG\r\n\x1A\n")
        //    {
        //        return "png";
        //    }
        //    else if (bytes.Length >= 4 && Encoding.UTF8.GetString(bytes, 0, 4) == "RIFF" && Encoding.UTF8.GetString(bytes, 8, 4) == "WEBP")
        //    {
        //        return "webp";
        //    }
        //    else if (bytes.Length >= 4 && (bytes[0] == 0xD0 && bytes[1] == 0xCF && bytes[2] == 0x11 && bytes[3] == 0xE0))
        //    {
        //        return "doc";
        //    }
        //    else if (bytes.Length >= 4 && (bytes[0] == 0x50 && bytes[1] == 0x4B && bytes[2] == 0x03 && bytes[3] == 0x04))
        //    {
        //        return "docx";
        //    }
        //    else
        //    {
        //        return "unknown";
        //    }
        //}




















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

