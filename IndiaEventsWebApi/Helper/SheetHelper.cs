using iTextSharp.text.pdf;
using iTextSharp.text;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Data;
using System.Text;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using System.Web.Http.Results;
using Aspose.Pdf.Operators;
using IndiaEventsWebApi.Helper;
using NPOI.HPSF;

namespace IndiaEventsWebApi.Helper
{
    public class SheetHelper
    {

        public static long GetColumnIdByName(Sheet sheet, string columnName)
        {
            foreach (var column in sheet.Columns)
            {
                if (column.Title == columnName)
                {
                    return column.Id.Value;
                }
            }
            return 0;
        }
        internal static string SQlFileinsertion(string base64, string name)
        {
            byte[] fileBytes = Convert.FromBase64String(base64);
            //var fileSize = (fileBytes.Length) / 1048576;
            var folderName = Path.Combine("Resources", "Images");
            var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
            if (!Directory.Exists(pathToSave))
            {
                Directory.CreateDirectory(pathToSave);
            }
            string fileType = GetFileType(fileBytes);
            string fileName = name + DateTime.Now.ToString("ddMMyyyymmss") + "." + fileType;
            string filePath = Path.Combine(pathToSave, fileName);
            File.WriteAllBytes(filePath, fileBytes);
            return fileName;
        }

        public static string GetValueByColumnName(Row row, List<string> columnNames, string columnName)
        {
            int columnIndex = columnNames.IndexOf(columnName);
            return columnIndex != -1 ? row.Cells[columnIndex].Value?.ToString() : null;
        }


        internal static List<Dictionary<string, object>> GetAttachmentsForRow(SmartsheetClient smartsheet, long sheetId, long row)
        {
            List<Dictionary<string, object>> attachmentsList = new List<Dictionary<string, object>>();
            var attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheetId, row, null);

            if (attachments.Data != null && attachments.Data.Count > 0)
            {
                foreach (var attachment in attachments.Data)
                {
                    long AID = (long)attachment.Id;
                    //Attachment file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheetId, AID);
                    Dictionary<string, object> attachmentInfo = new Dictionary<string, object>
                    {
                        { "Name", attachment.Name },
                        { "Id", AID },
                       // { "Url", file.Url },
                       // { "base64", UrlToBaseValue(file.Url) },
                        {"SheetId",sheetId }
                    };
                    attachmentsList.Add(attachmentInfo);
                }
            }
            return attachmentsList;
        }
        internal static List<Dictionary<string, object>> GetAttachmentsIdForRow(SmartsheetClient smartsheet, long sheetId, long row)
        {
            List<Dictionary<string, object>> attachmentsList = new List<Dictionary<string, object>>();
            PaginatedResult<Attachment>? attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheetId, row, null);

            if (attachments.Data != null && attachments.Data.Count > 0)
            {
                foreach (var attachment in attachments.Data)
                {

                    Dictionary<string, object> attachmentInfo = new Dictionary<string, object>
                    {
                        { "Name", attachment.Name },
                        { "Id", attachment.Id },
                         {"SheetId",sheetId }

                    };
                    attachmentsList.Add(attachmentInfo);
                }
            }
            return attachmentsList;
        }


        //UrlToBase64

        public static string UrlToBaseValue(string url)
        {
            using HttpClient client = new();
            byte[] fileContent = client.GetByteArrayAsync(url).Result;
            string base64String = Convert.ToBase64String(fileContent);
            string fileType = GetFileType(fileContent);
            string prefix = $"data:application/{fileType};base64,";
            string prefixedBase64String = prefix + base64String;
            return prefixedBase64String;
        }

        private static string GetFileType(string url)
        {
            string filename = Path.GetFileName(url);
            string extension = Path.GetExtension(filename);
            return extension.TrimStart('.').ToLower();
        }


        // number check 
        public static double NumCheck(string val)
        {
            if (string.IsNullOrEmpty(val))
            {
                return 0;
            }
            //return Convert.ToInt32(val);
            double result;
            return double.TryParse(val, out result) ? result : 0;
        }

        public static object MisCodeCheck(string val)
        {
            if (string.IsNullOrEmpty(val))
            {
                return 0;
            }

            int result;
            if (int.TryParse(val, out result))
            {
                return result;
            }
            else
            {
                return val;
            }
            //return int.TryParse(val, out result) ? result : 0;
        }

        // get sheet using sheetId
        internal static Sheet GetSheetById(SmartsheetClient smartsheet, string sheetId)
        {
            try
            {
                long.TryParse(sheetId, out long parsedSheetId);
                return smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);

            }
            catch (Exception ex)
            {
                return GetSheetById(smartsheet, sheetId);
            }

        }

        // get sheet data using sheet access
        internal static List<Dictionary<string, object>> GetSheetData(Sheet sheet)
        {
            List<Dictionary<string, object>> sheetData = new List<Dictionary<string, object>>();
            List<string> columnNames = sheet.Columns.Select(column => column.Title).ToList();

            foreach (Row row in sheet.Rows)
            {
                Dictionary<string, object> rowData = new Dictionary<string, object>();

                for (int i = 0; i < row.Cells.Count && i < columnNames.Count; i++)
                {
                    rowData[columnNames[i]] = row.Cells[i].Value;
                }

                sheetData.Add(rowData);
            }

            return sheetData;
        }


        // file convert from base 64 to file  and store in local
        internal static string testingFile(string base64, string name)
        {
            byte[] fileBytes = Convert.FromBase64String(base64);
            //var fileSize = (fileBytes.Length) / 1048576;
            var folderName = Path.Combine("Resources", "Images");
            var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
            if (!Directory.Exists(pathToSave))
            {
                Directory.CreateDirectory(pathToSave);
            }
            string fileType = GetFileType(fileBytes);
            string fileName =  name + DateTime.Now.ToString("ddMMyyyymmss") + "." + fileType;
            string filePath = Path.Combine(pathToSave, fileName);
            File.WriteAllBytes(filePath, fileBytes);
            return filePath;
        }

        // delete local file if file exists
        internal static string DeleteFile(string filePath)
        {
            System.IO.File.Delete(filePath);
            return "ok";
        }

        //internal static byte[] exportpdf(DataTable dtEmployee, string EventCode, string EventName, string EventDate, string EventVenue, DataTable dtMai)
        //{
        //    System.IO.MemoryStream ms = new System.IO.MemoryStream();
        //    iTextSharp.text.Rectangle rec = new iTextSharp.text.Rectangle(PageSize.A4);
        //    rec.BackgroundColor = new BaseColor(System.Drawing.Color.Olive);
        //    Document doc = new Document(rec);
        //    doc.SetPageSize(iTextSharp.text.PageSize.A4);
        //    PdfWriter writer = PdfWriter.GetInstance(doc, ms);
        //    doc.Open();
        //    BaseFont bfntHead = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //    iTextSharp.text.Font fntHead = new iTextSharp.text.Font(bfntHead, 16, 1, iTextSharp.text.BaseColor.BLUE);
        //    Paragraph prgHeading = new Paragraph();
        //    prgHeading.Alignment = Element.ALIGN_LEFT;
        //    prgHeading.Add(new Chunk("Attendance Sheet".ToUpper(), fntHead));
        //    doc.Add(prgHeading);
        //    Paragraph prgGeneratedBY = new Paragraph();
        //    BaseFont btnAuthor = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //    iTextSharp.text.Font fntAuthor = new iTextSharp.text.Font(btnAuthor, 8, 2, iTextSharp.text.BaseColor.BLUE);
        //    prgGeneratedBY.Alignment = Element.ALIGN_RIGHT;
        //    doc.Add(prgGeneratedBY);
        //    Paragraph p = new Paragraph(new Chunk(new iTextSharp.text.pdf.draw.LineSeparator(0.0F, 100.0F, iTextSharp.text.BaseColor.BLACK, Element.ALIGN_LEFT, 1)));
        //    doc.Add(p);
        //    doc.Add(new Chunk("\n", fntHead));

        //    Paragraph pBody = new Paragraph();
        //    pBody.Add(new Chunk("Event Code:" + EventCode));
        //    pBody.Add(new Chunk("\nEvent Name:" + EventName));
        //    pBody.Add(new Chunk("\nEvent Date:" + EventDate));
        //    pBody.Add(new Chunk("\nEvent Venue:" + EventVenue));
        //    pBody.Add(new Chunk("\n\nSpeakers: "));


        //    string hcpNames = string.Join(", ", dtMai.AsEnumerable().Select(row => row["HCPName"].ToString()));
        //    pBody.Add(new Chunk(" " + hcpNames));
        //    doc.Add(pBody);
        //    doc.Add(new Paragraph("\n "));
        //    PdfPTable table = new PdfPTable(dtEmployee.Columns.Count);
        //    table.WidthPercentage = 100;
        //    float[] columnWidths = Enumerable.Range(0, dtEmployee.Columns.Count).Select(i => i == dtEmployee.Columns.IndexOf("HCPName") ? 2f : 1f).ToArray(); /*Count).Select(i => 1f).ToArray();*/
        //    table.SetWidths(columnWidths);

        //    for (int i = 0; i < dtEmployee.Columns.Count; i++)
        //    {
        //        string cellText = dtEmployee.Columns[i].ColumnName;
        //        PdfPCell cell = new PdfPCell();
        //        cell.Phrase = new Phrase(cellText, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 8, 1, new BaseColor(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        //        cell.BackgroundColor = new BaseColor(System.Drawing.ColorTranslator.FromHtml("#C8C8C8"));
        //        cell.HorizontalAlignment = Element.ALIGN_CENTER;
        //        cell.PaddingBottom = 5;
        //        table.AddCell(cell);
        //    }
        //    for (int i = 0; i < dtEmployee.Rows.Count; i++)
        //    {
        //        for (int j = 0; j < dtEmployee.Columns.Count; j++)
        //        {
        //            table.AddCell(dtEmployee.Rows[i][j].ToString());
        //        }
        //    }
        //    doc.Add(table);
        //    doc.Close();
        //    byte[] result = ms.ToArray();
        //    return result;
        //}

        internal static string GetContentType(string fileExtension)
        {
            switch (fileExtension.ToLower())
            {
                case "jpg":
                case "jpeg":
                    return "image/jpeg";
                case "pdf":
                    return "application/pdf";
                case "gif":
                    return "image/gif";
                case "png":
                    return "image/png";
                case "webp":
                    return "image/webp";
                case "doc":
                    return "application/msword";
                case "docx":
                    return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                default:
                    return "application/octet-stream";
            }
        }

        internal static string GetFileType(byte[] bytes)
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
            else if (bytes.Length >= 8 && bytes[0] == 0x89 &&
                bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47 && bytes[4] == 0x0D &&
                bytes[5] == 0x0A && bytes[6] == 0x1A && bytes[7] == 0x0A
                /*Encoding.UTF8.GetString(bytes, 0, 8) == "PNG"*/)
            {
                return "png";
            }
            else if (bytes.Length >= 4 && Encoding.UTF8.GetString(bytes, 0, 4) == "RIFF" && Encoding.UTF8.GetString(bytes, 8, 4) == "WEBP")
            {
                return "webp";
            }
            else if (bytes.Length >= 4 && (bytes[0] == 0xD0 && bytes[1] == 0xCF && bytes[2] == 0x11 && bytes[3] == 0xE0))
            {
                return "doc"; // .doc format
            }
            else if (bytes.Length >= 4 && (bytes[0] == 0x50 && bytes[1] == 0x4B && bytes[2] == 0x03 && bytes[3] == 0x04))
            {
                return "docx"; // .docx format
            }
            else if (IsLikelyText(bytes))
            {
                return "txt"; // .txt format
            }
            else
            {
                return "unknown";
            }
        }

        private static bool IsLikelyText(byte[] bytes)
        {
            for (int i = 0; i < bytes.Length; i++)
            {
                byte b = bytes[i];
                if (b < 7 || (b > 14 && b < 32 && b != 10 && b != 13))
                {
                    return false;
                }

            }
            return true;
        }

        //internal static byte[] exportAttendencepdf(DataTable dtEmployee, string EventCode, string EventName, string EventDate, string EventVenue, string speakers)
        //{
        //    System.IO.MemoryStream ms = new System.IO.MemoryStream();
        //    iTextSharp.text.Rectangle rec = new iTextSharp.text.Rectangle(PageSize.A4);
        //    rec.BackgroundColor = new BaseColor(System.Drawing.Color.Olive);
        //    Document doc = new Document(rec);
        //    doc.SetPageSize(iTextSharp.text.PageSize.A4);
        //    PdfWriter writer = PdfWriter.GetInstance(doc, ms);
        //    doc.Open();
        //    BaseFont bfntHead = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //    iTextSharp.text.Font fntHead = new iTextSharp.text.Font(bfntHead, 16, 1, iTextSharp.text.BaseColor.BLUE);
        //    Paragraph prgHeading = new Paragraph();
        //    prgHeading.Alignment = Element.ALIGN_LEFT;
        //    prgHeading.Add(new Chunk("Attendance Sheet".ToUpper(), fntHead));
        //    doc.Add(prgHeading);
        //    Paragraph prgGeneratedBY = new Paragraph();
        //    BaseFont btnAuthor = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //    iTextSharp.text.Font fntAuthor = new iTextSharp.text.Font(btnAuthor, 8, 2, iTextSharp.text.BaseColor.BLUE);
        //    prgGeneratedBY.Alignment = Element.ALIGN_RIGHT;
        //    doc.Add(prgGeneratedBY);
        //    Paragraph p = new Paragraph(new Chunk(new iTextSharp.text.pdf.draw.LineSeparator(0.0F, 100.0F, iTextSharp.text.BaseColor.BLACK, Element.ALIGN_LEFT, 1)));
        //    doc.Add(p);
        //    doc.Add(new Chunk("\n", fntHead));

        //    Paragraph pBody = new Paragraph();
        //    pBody.Add(new Chunk("Event Code:" + EventCode));
        //    pBody.Add(new Chunk("\nEvent Name:" + EventName));
        //    pBody.Add(new Chunk("\nEvent Date:" + EventDate));
        //    pBody.Add(new Chunk("\nEvent Venue:" + EventVenue));
        //    pBody.Add(new Chunk("\n\nSpeakers: "));


        //    string hcpNames = speakers;//string.Join(", ", dtMai.AsEnumerable().Select(row => row["HCPName"].ToString()));
        //    pBody.Add(new Chunk(" " + hcpNames));
        //    doc.Add(pBody);
        //    doc.Add(new Paragraph("\n "));
        //    PdfPTable table = new PdfPTable(dtEmployee.Columns.Count);
        //    table.WidthPercentage = 100;
        //    float[] columnWidths = Enumerable.Range(0, dtEmployee.Columns.Count).Select(i => i == dtEmployee.Columns.IndexOf("HCPName") ? 2f : 1f).ToArray(); /*Count).Select(i => 1f).ToArray();*/
        //    table.SetWidths(columnWidths);

        //    for (int i = 0; i < dtEmployee.Columns.Count; i++)
        //    {
        //        string cellText = dtEmployee.Columns[i].ColumnName;
        //        PdfPCell cell = new PdfPCell();
        //        cell.Phrase = new Phrase(cellText, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 8, 1, new BaseColor(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        //        cell.BackgroundColor = new BaseColor(System.Drawing.ColorTranslator.FromHtml("#C8C8C8"));
        //        cell.HorizontalAlignment = Element.ALIGN_CENTER;
        //        cell.PaddingBottom = 5;
        //        table.AddCell(cell);
        //    }
        //    for (int i = 0; i < dtEmployee.Rows.Count; i++)
        //    {
        //        for (int j = 0; j < dtEmployee.Columns.Count; j++)
        //        {
        //            table.AddCell(dtEmployee.Rows[i][j].ToString());
        //        }
        //    }
        //    doc.Add(table);
        //    doc.Close();
        //    byte[] result = ms.ToArray();
        //    return result;
        //}

        //internal static byte[] exportAttendencepdfnew(DataTable dtEmployee,DataTable MenariniEmployee, string EventCode, string EventName, string EventDate, string EventVenue, string speakers )
        //{
        //    System.IO.MemoryStream ms = new System.IO.MemoryStream();
        //    iTextSharp.text.Rectangle rec = new iTextSharp.text.Rectangle(PageSize.A4);
        //    rec.BackgroundColor = new BaseColor(System.Drawing.Color.Olive);
        //    Document doc = new Document(rec);
        //    doc.SetPageSize(iTextSharp.text.PageSize.A4);
        //    PdfWriter writer = PdfWriter.GetInstance(doc, ms);
        //    doc.Open();
        //    BaseFont bfntHead = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //    iTextSharp.text.Font fntHead = new iTextSharp.text.Font(bfntHead, 16, 1, iTextSharp.text.BaseColor.BLUE);
        //    Paragraph prgHeading = new Paragraph();
        //    prgHeading.Alignment = Element.ALIGN_LEFT;
        //    prgHeading.Add(new Chunk("Attendance Sheet".ToUpper(), fntHead));
        //    doc.Add(prgHeading);
        //    Paragraph prgGeneratedBY = new Paragraph();
        //    BaseFont btnAuthor = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //    iTextSharp.text.Font fntAuthor = new iTextSharp.text.Font(btnAuthor, 8, 2, iTextSharp.text.BaseColor.BLUE);
        //    prgGeneratedBY.Alignment = Element.ALIGN_RIGHT;
        //    doc.Add(prgGeneratedBY);
        //    Paragraph p = new Paragraph(new Chunk(new iTextSharp.text.pdf.draw.LineSeparator(0.0F, 100.0F, iTextSharp.text.BaseColor.BLACK, Element.ALIGN_LEFT, 1)));
        //    doc.Add(p);
        //    doc.Add(new Chunk("\n", fntHead));

        //    Paragraph pBody = new Paragraph();
        //    pBody.Add(new Chunk("Event Code:" + EventCode));
        //    pBody.Add(new Chunk("\nEvent Name:" + EventName));
        //    pBody.Add(new Chunk("\nEvent Date:" + EventDate));
        //    pBody.Add(new Chunk("\nEvent Venue:" + EventVenue));
        //    pBody.Add(new Chunk("\n\nSpeakers: "));


        //    string hcpNames = speakers;//string.Join(", ", dtMai.AsEnumerable().Select(row => row["HCPName"].ToString()));
        //    pBody.Add(new Chunk(" " + hcpNames));
        //    doc.Add(pBody);
        //    doc.Add(new Paragraph("\n "));

        //    PdfPTable table = new PdfPTable(dtEmployee.Columns.Count);
        //    table.WidthPercentage = 100;
        //    float[] columnWidths = Enumerable.Range(0, dtEmployee.Columns.Count).Select(i => i == dtEmployee.Columns.IndexOf("HCPName") ? 2f : 1f).ToArray(); /*Count).Select(i => 1f).ToArray();*/
        //    table.SetWidths(columnWidths);

        //    for (int i = 0; i < dtEmployee.Columns.Count; i++)
        //    {
        //        string cellText = dtEmployee.Columns[i].ColumnName;
        //        PdfPCell cell = new PdfPCell();
        //        cell.Phrase = new Phrase(cellText, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 8, 1, new BaseColor(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        //        cell.BackgroundColor = new BaseColor(System.Drawing.ColorTranslator.FromHtml("#C8C8C8"));
        //        cell.HorizontalAlignment = Element.ALIGN_CENTER;
        //        cell.PaddingBottom = 5;
        //        table.AddCell(cell);
        //    }
        //    for (int i = 0; i < dtEmployee.Rows.Count; i++)
        //    {
        //        for (int j = 0; j < dtEmployee.Columns.Count; j++)
        //        {
        //            table.AddCell(dtEmployee.Rows[i][j].ToString());
        //        }
        //    }
        //    doc.Add(table);
        //    doc.Add(new Paragraph("\n "));

        //    PdfPTable Menarinitable = new PdfPTable(MenariniEmployee.Columns.Count);
        //    Menarinitable.WidthPercentage = 100;
        //    float[] columnWidth = Enumerable.Range(0, MenariniEmployee.Columns.Count).Select(i => i == MenariniEmployee.Columns.IndexOf("HCPName") ? 2f : 1f).ToArray(); /*Count).Select(i => 1f).ToArray();*/
        //    Menarinitable.SetWidths(columnWidths);

        //    for (int i = 0; i < MenariniEmployee.Columns.Count; i++)
        //    {
        //        string cellText = MenariniEmployee.Columns[i].ColumnName;
        //        PdfPCell cell = new PdfPCell();
        //        cell.Phrase = new Phrase(cellText, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 8, 1, new BaseColor(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        //        cell.BackgroundColor = new BaseColor(System.Drawing.ColorTranslator.FromHtml("#C8C8C8"));
        //        cell.HorizontalAlignment = Element.ALIGN_CENTER;
        //        cell.PaddingBottom = 5;
        //        Menarinitable.AddCell(cell);
        //    }
        //    for (int i = 0; i < MenariniEmployee.Rows.Count; i++)
        //    {
        //        for (int j = 0; j < MenariniEmployee.Columns.Count; j++)
        //        {
        //            Menarinitable.AddCell(MenariniEmployee.Rows[i][j].ToString());
        //        }
        //    }
        //    doc.Add(Menarinitable);



        //    doc.Close();
        //    byte[] result = ms.ToArray();
        //    return result;
        //}



        internal static byte[] exportAttendencepdfnew(DataTable dtEmployee, DataTable MenariniEmployee, string EventCode, string EventName, string EventDate, string EventVenue, string speakers)
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
            pBody.Add(new Chunk("\nSpeakers: "));


            string hcpNames = speakers;//string.Join(", ", dtMai.AsEnumerable().Select(row => row["HCPName"].ToString()));
            pBody.Add(new Chunk(" " + hcpNames));
            doc.Add(pBody);
            doc.Add(new Paragraph("\n "));

            ////Paragraph hcpTableTitle = new Paragraph("HCP Table");
            //hcpTableTitle.Alignment = Element.ALIGN_LEFT;
            //doc.Add(hcpTableTitle);
            //doc.Add(new Paragraph("\n "));
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
            doc.Add(new Paragraph("\n "));

            Paragraph MenariniTableTitle = new Paragraph("MIPL Employees");
            MenariniTableTitle.Alignment = Element.ALIGN_LEFT;
            doc.Add(MenariniTableTitle);
            doc.Add(new Paragraph("\n"));

            //Paragraph paragraph = new Paragraph("\n").setma;
            //paragraph.SetMarginBottom(20); // Adjust the bottom margin as needed for padding
            //doc.Add(paragraph);
            PdfPTable Menarinitable = new PdfPTable(MenariniEmployee.Columns.Count);
            Menarinitable.WidthPercentage = 100;
            float[] columnWidth = Enumerable.Range(0, MenariniEmployee.Columns.Count).Select(i => i == MenariniEmployee.Columns.IndexOf("HCPName") ? 2f : 1f).ToArray(); /*Count).Select(i => 1f).ToArray();*/
            Menarinitable.SetWidths(columnWidth);

            Dictionary<string, string> columnTitleMappings = new Dictionary<string, string>
            {
                {"S.No","S.No" },
                { "HCPName", "Name" },
                { "Employee Code", "Employee Code" },
                { "Designation", "Designation" },
                { "Sign", "Sign" },
                // Add more mappings if needed
            };

            for (int i = 0; i < MenariniEmployee.Columns.Count; i++)
            {
                string columnName = MenariniEmployee.Columns[i].ColumnName;
                string mappedTitle = columnTitleMappings.ContainsKey(columnName) ? columnTitleMappings[columnName] : columnName;

                PdfPCell cell = new PdfPCell();
                cell.Phrase = new Phrase(mappedTitle, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 8, 1, new BaseColor(System.Drawing.ColorTranslator.FromHtml("#000000"))));
                cell.BackgroundColor = new BaseColor(System.Drawing.ColorTranslator.FromHtml("#C8C8C8"));
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.PaddingBottom = 5;
                Menarinitable.AddCell(cell);
            }
            for (int i = 0; i < MenariniEmployee.Rows.Count; i++)
            {
                for (int j = 0; j < MenariniEmployee.Columns.Count; j++)
                {
                    Menarinitable.AddCell(MenariniEmployee.Rows[i][j].ToString());
                }
            }

            //for (int i = 0; i < MenariniEmployee.Columns.Count; i++)
            //{
            //    string cellText = MenariniEmployee.Columns[i].ColumnName;
            //    PdfPCell cell = new PdfPCell();
            //    cell.Phrase = new Phrase(cellText, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 8, 1, new BaseColor(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            //    cell.BackgroundColor = new BaseColor(System.Drawing.ColorTranslator.FromHtml("#C8C8C8"));
            //    cell.HorizontalAlignment = Element.ALIGN_CENTER;
            //    cell.PaddingBottom = 5;
            //    Menarinitable.AddCell(cell);
            //}
            //for (int i = 0; i < MenariniEmployee.Rows.Count; i++)
            //{
            //    for (int j = 0; j < MenariniEmployee.Columns.Count; j++)
            //    {
            //        Menarinitable.AddCell(MenariniEmployee.Rows[i][j].ToString());
            //    }
            //}
            doc.Add(Menarinitable);



            doc.Close();
            byte[] result = ms.ToArray();
            return result;
        }



        //internal static Row Deviations(Sheet sheet2,object formdata)
        //{
        //    Row newRow = new()
        //    {
        //        Cells = new List<Cell>()
        //        {
        //                    new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "% Allocation"), Value = formdata.PercentAllocation },
        //                    new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Brands"), Value = formdata.BrandName },
        //                    new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Project ID"), Value = formdata.ProjectId },
        //                    new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "EventId/EventRequestId"), Value = val }
        //                }
        //    };
        //    return newRow;
        //}



        //           if (formDataList.class1.EventOpen30days == "Yes" || formDataList.class1.EventWithin7days == "Yes" || formDataList.class1.FB_Expense_Excluding_Tax == "Yes" || formDataList.class1.IsDeviationUpload == "Yes")
        //                {
        //                    List<string> DeviationNames = new List<string>();
        //                    foreach (var p in formDataList.class1.DeviationFiles)
        //                    {
        //                        string[] words = p.Split(':');
        //        var r = words[0];

        //        DeviationNames.Add(r);
        //                    }
        //                    foreach (var deviationname in DeviationNames)
        //                    {
        //                        var file = deviationname.Split(".")[0];
        //    var eventId = val;
        //                        try
        //                        {
        //                            var newRow7 = new Row();
        //    newRow7.Cells = new List<Cell>();

        //                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId
        //});
        //newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.class1.EventTopic });
        //newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formDataList.class1.EventType });
        //newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formDataList.class1.EventDate });
        //newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "StartTime"), Value = formDataList.class1.StartTime });
        //newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EndTime"), Value = formDataList.class1.EndTime });
        //newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "VenueName"), Value = formDataList.class1.VenueName });
        //newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "City"), Value = formDataList.class1.City });
        //newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "State"), Value = formDataList.class1.State });


        //if (file == "30DaysDeviationFile")
        //{
        //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Outstanding with initiator for more than 45 days" });
        //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = formDataList.class1.EventOpen30days });
        //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = SheetHelper.NumCheck(formDataList.class1.EventOpen30dayscount) });
        //}
        //else if (file == "7DaysDeviationFile")
        //{
        //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "5 days from the Event Date" });
        //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = formDataList.class1.EventWithin7days });

        //}
        //else if (file == "ExpenseExcludingTax")
        //{
        //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Food and Beverages expense exceeds 1500" });
        //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "PRE-F&B Expense Excluding Tax"), Value = formDataList.class1.FB_Expense_Excluding_Tax });
        //}
        //else if (file == "Travel_Accomodation3LExceededFile")
        //{
        //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Travel/AccomodationAggregate Limit of 3,00,000 is Exceeded" });
        //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Travel/Accomodation 3,00,000 Exceeded Trigger"), Value = "Yes" });//formDataList.class1.FB_Expense_Excluding_Tax });
        //}
        //else if (file == "TrainerHonorarium12LExceededFile")
        //{
        //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Honorarium Aggregate Limit of 12,00,000 is Exceeded" });
        //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Trainer Honorarium 12,00,000 Exceeded Trigger"), Value = "Yes" }); //formDataList.class1.FB_Expense_Excluding_Tax });
        //}
        //else if (file == "HCPHonorarium6LExceededFile")
        //{
        //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Honorarium Aggregate Limit of 6,00,000 is Exceeded" });
        //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HCP Honorarium 6,00,000 Exceeded Trigger"), Value = "Yes" }); // formDataList.class1.FB_Expense_Excluding_Tax });
        //}
        //newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.class1.Sales_Head });
        //newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formDataList.class1.FinanceHead });
        //newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formDataList.class1.InitiatorName });
        //newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.class1.Initiator_Email });


        //var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId7, new Row[] { newRow7 });




        //var j = 1;
        //foreach (var p in formDataList.class1.DeviationFiles)
        //{
        //    string[] words = p.Split(':');
        //    var r = words[0];
        //    var q = words[1];
        //    if (deviationname == r)
        //    {

        //        var name = r.Split(".")[0];

        //        var filePath = SheetHelper.testingFile(q, val, name);


        //        //byte[] fileBytes = Convert.FromBase64String(q);
        //        //var folderName = Path.Combine("Resources", "Images");
        //        //var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
        //        //if (!Directory.Exists(pathToSave))
        //        //{
        //        //    Directory.CreateDirectory(pathToSave);
        //        //}

        //        //string fileType = SheetHelper.GetFileType(fileBytes);
        //        //string fileName = r;
        //        //// string fileName = val+x + ": AttachedFile." + fileType;
        //        //string filePath = Path.Combine(pathToSave, fileName);


        //        var addedRow = addeddeviationrow[0];

        //        //System.IO.File.WriteAllBytes(filePath, fileBytes);
        //        //string type = SheetHelper.GetContentType(fileType);
        //        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
        //                parsedSheetId7, addedRow.Id.Value, filePath, "application/msword");
        //        var attachmentinmain = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
        //               parsedSheetId1, addedRows[0].Id.Value, filePath, "application/msword");
        //        j++;
        //        if (System.IO.File.Exists(filePath))
        //        {
        //            SheetHelper.DeleteFile(filePath);
        //        }
        //    }


        //}
        //                        }
        //                        catch (Exception ex)
        //                        {
        //    return BadRequest(ex.Message);
        //}
        //                    }

        //                }

        internal static List<object> ConvertToJsonObject(string data)
        {
            string[] lines = data.Split('\n');
            List<object> resultList = new List<object>();

            foreach (var line in lines)
            {
                string[] values = line.Split('|');

                if (values.Length > 0)
                {
                    Dictionary<string, string> item = new Dictionary<string, string>();

                    foreach (var value in values)
                    {
                        string[] keyValue = value.Trim().Split(':');
                        if (keyValue.Length == 2)
                        {
                            item.Add(keyValue[0].Trim(), keyValue[1].Trim());
                        }
                    }

                    resultList.Add(item);
                }
            }

            return resultList;
        }



    }

}
