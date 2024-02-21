using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Data;
using System.Text;

namespace IndiaEventsWebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class AttendenceSheetController : ControllerBase
    {

        private readonly string accessToken;
        private readonly IConfiguration configuration;
        public AttendenceSheetController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
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
                var a = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(parsedSheetId1, rowId, null);





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
                foreach (var x in a.Data)
                {
                    long Id = (long)x.Id;
                    var Fullname = x.Name.Split("-");
                    var splitName = Fullname[0];

                    if (splitName == "AttendenceSheet")
                    {

                        smartsheet.SheetResources.AttachmentResources.DeleteAttachment(
                          parsedSheetId1,           // sheetId
                          Id            // attachmentId
                        );

                    }
                }

                byte[] fileBytes = exportpdf(dtMai, EventCode, EventName, EventDate, EventVenue, dtMai);
                string filename = "AttendenceSheet-" + EventID + ".pdf";
                var folderName = Path.Combine("Resources", "Images");
                var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                if (!Directory.Exists(pathToSave))
                {
                    Directory.CreateDirectory(pathToSave);
                }
                string fileType = GetFileType(fileBytes);
                string filePath = Path.Combine(pathToSave, filename);
                System.IO.File.WriteAllBytes(filePath, fileBytes);

                var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(parsedSheetId1, rowId, filePath, "application/msword");

                var FileUrl = "";
                var AID = (long)attachment.Id;
                var file = smartsheet.SheetResources.AttachmentResources.GetAttachment(parsedSheetId1, AID);
                FileUrl = file.Url;

                //string downloadLink = GetDownloadLink(filename);
                if (System.IO.File.Exists(filePath))
                {
                    System.IO.File.Delete(filePath);
                }


                return Ok(new { FileUrl});


            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        private byte[] exportpdf(DataTable dtEmployee, string EventCode, string EventName, string EventDate, string EventVenue, DataTable dtMai)
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
            pBody.Add(new Chunk("\nEvent Vanue:" + EventVenue));
            pBody.Add(new Chunk("\n\nSpeakers: "));

            //foreach(DataRow row in dtMai.Rows)
            //{
            //    string hcpName = row["HCPName"].ToString();
            //    pBody.Add(new Chunk(" " + hcpName));
            //}

            string hcpNames = string.Join(", ", dtMai.AsEnumerable().Select(row => row["HCPName"].ToString()));
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


        private string GetFileType(byte[] bytes)
        {

            if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xD8)
            {
                return "jpg";
            }
            else if (bytes.Length >= 4 && Encoding.UTF8.GetString(bytes, 0, 4) == "%PDF")
            {
                return "pdf";
            }
            else if (bytes.Length >= 3 && Encoding.UTF8.GetString(bytes, 0, 3) == "GIF")
            {
                return "gif";
            }
            else if (bytes.Length >= 8 && Encoding.UTF8.GetString(bytes, 0, 8) == "PNG\r\n\x1A\n")
            {
                return "png";
            }
            else if (bytes.Length >= 4 && Encoding.UTF8.GetString(bytes, 0, 4) == "RIFF" && Encoding.UTF8.GetString(bytes, 8, 4) == "WEBP")
            {
                return "webp";
            }
            else if (bytes.Length >= 4 && (bytes[0] == 0xD0 && bytes[1] == 0xCF && bytes[2] == 0x11 && bytes[3] == 0xE0))
            {
                return "doc";
            }
            else if (bytes.Length >= 4 && (bytes[0] == 0x50 && bytes[1] == 0x4B && bytes[2] == 0x03 && bytes[3] == 0x04))
            {
                return "docx";
            }
            else
            {
                return "unknown";
            }
        }



    }

}
