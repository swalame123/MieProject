using IndiaEvents.Models.Models.EventTypeSheets;
using IndiaEventsWebApi.Helper;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Data;
using System.Runtime.CompilerServices;
using System.Text;

namespace IndiaEventsWebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class AttendenceSheetController : ControllerBase
    {

        private readonly string accessToken;
        private readonly IConfiguration configuration;
        private readonly SmartsheetClient smartsheet;
        private readonly Sheet sheet_SpeakerCode;
        private readonly Sheet sheet;
        private readonly Sheet sheet1;
        private readonly Sheet processSheetData;
        public AttendenceSheetController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

            string sheetId_SpeakerCode = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
            string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;

            string sheetId1 = configuration.GetSection("SmartsheetSettings:Class1").Value;
            string processSheet = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;
            smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            processSheetData = SheetHelper.GetSheetById(smartsheet, processSheet);
            sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
            sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
            sheet_SpeakerCode = SheetHelper.GetSheetById(smartsheet, sheetId_SpeakerCode);
        }
        [HttpGet("GenerateAttendencePDF")]
        public IActionResult GenerateAttendencePDF(string EventID)
        {
            try
            {

                var EventCode = "";
                var EventName = "";
                var EventDate = "";
                var EventVenue = "";
                List<string> Speakers = new List<string>();

                // SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

                //string sheetId_SpeakerCode = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
                //long.TryParse(sheetId_SpeakerCode, out long parsedSheetId_SpeakerCode);
                //Sheet sheet_SpeakerCode = smartsheet.SheetResources.GetSheet(parsedSheetId_SpeakerCode, null, null, null, null, null, null, null);

                //string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
                //long.TryParse(sheetId, out long parsedSheetId);
                //Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);

                //string sheetId1 = configuration.GetSection("SmartsheetSettings:Class1").Value;
                //long.TryParse(sheetId1, out long parsedSheetId1);
                //Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId1, null, null, null, null, null, null, null);


                //string processSheet = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;
                //long.TryParse(processSheet, out long parsedProcessSheet);
                //Sheet processSheetData = smartsheet.SheetResources.GetSheet(parsedProcessSheet, null, null, null, null, null, null, null);

                long rowId = 0;
                Column IdColumn = sheet1.Columns.FirstOrDefault(column => string.Equals(column.Title, "EventId/EventRequestId", StringComparison.OrdinalIgnoreCase));
                if (IdColumn != null)
                {
                    // Find all rows with the specified speciality
                    List<Row> targetRows = sheet1.Rows
                        .Where(row => row.Cells.Any(cell => cell.ColumnId == IdColumn.Id && cell.Value.ToString() == EventID))
                        .ToList();

                    if (targetRows.Any())
                    {
                        var rowIds = targetRows.Select(row => row.Id).ToList();
                        rowId = (long)rowIds[0];
                    }

                }
                var a = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments((long)sheet1.Id, rowId, null);





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
                var url = "";
                foreach (var x in a.Data)
                {
                    long Id = (long)x.Id;
                    var Fullname = x.Name.Split("-");
                    var splitName = Fullname[0];

                    if (splitName == "AttendanceSheet")
                    {

                        var ExistingFile = smartsheet.SheetResources.AttachmentResources.GetAttachment((long)sheet1.Id, Id);
                        url = ExistingFile.Url;

                        //smartsheet.SheetResources.AttachmentResources.DeleteAttachment(
                        //  (long)sheet1.Id,           // sheetId
                        //  Id            // attachmentId
                        //);

                    }
                }
                if (url != "")
                {
                    return Ok(new { url });
                }
                else
                {
                    string resultString = string.Join(", ", Speakers);
                    byte[] fileBytes = SheetHelper.exportAttendencepdfnew(dtMai, MenariniTable, EventCode, EventName, EventDate, EventVenue, resultString);
                    string filename = "AttendanceSheet-" + EventID + ".pdf";
                    var folderName = Path.Combine("Resources", "Images");
                    var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                    if (!Directory.Exists(pathToSave))
                    {
                        Directory.CreateDirectory(pathToSave);
                    }
                    string fileType = SheetHelper.GetFileType(fileBytes);
                    string filePath = Path.Combine(pathToSave, filename);
                    System.IO.File.WriteAllBytes(filePath, fileBytes);

                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile((long)sheet1.Id, rowId, filePath, "application/msword");

                    //var FileUrl = "";
                    var AID = (long)attachment.Id;
                    var file = smartsheet.SheetResources.AttachmentResources.GetAttachment((long)sheet1.Id, AID);
                    url = file.Url;

                    //string downloadLink = GetDownloadLink(filename);
                    if (System.IO.File.Exists(filePath))
                    {
                        System.IO.File.Delete(filePath);
                    }


                    return Ok(new { url });
                }



            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }


        [HttpGet("AttendencesheetBase64value")]
        public IActionResult AttendencesheetBase64value(string EventID)
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:HonorariumPayment").Value;
                Sheet sheetAs = SheetHelper.GetSheetById(smartsheet, sheetId);
                Column SpecialityColumn = sheetAs.Columns.FirstOrDefault(column => string.Equals(column.Title, "EventId/EventRequestId", StringComparison.OrdinalIgnoreCase));

                Row existingRow = sheetAs.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == EventID));
                var data = "";
                if (existingRow != null)
                {
                    var attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheetAs.Id.Value, existingRow.Id.Value, null);
                    var url = "";
                    var FileName = "";
                    foreach (var attachment in attachments.Data)
                    {
                        if (attachment != null)
                        {
                            var name = attachment.Name;
                            var s = name.Split(".")[0];
                            //if (s == (EventID + "-AttendanceSheet")) //RQID757-AttendenceSheet
                            if(name.ToLower().Contains("attendancesheet"))
                            {
                                var AID = (long)attachment.Id;
                                var file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheetAs.Id.Value, AID);
                                FileName = file.Name;
                                url = file.Url;
                            }


                        }
                    }
                    using (HttpClient client = new HttpClient())
                    {
                        var fileContent = client.GetByteArrayAsync(url).Result;
                        var base64String = Convert.ToBase64String(fileContent);


                        data = $"{FileName}:{base64String}";


                    }
                }





                return Ok(new { data });

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }





    }

}
