using iTextSharp.text.pdf;
using iTextSharp.text;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Data;
using System.Text;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using System.Web.Http.Results;

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

        internal static string testingFile(string Invoice_QuotationUpload, string eventId , string name)
        {
            byte[] fileBytes = Convert.FromBase64String(Invoice_QuotationUpload);
            var folderName = Path.Combine("Resources", "Images");
            var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
            if (!Directory.Exists(pathToSave))
            {
                Directory.CreateDirectory(pathToSave);
            }
            string fileType = GetFileType(fileBytes);
            string fileName = eventId + "-"  +name+ "." + fileType;
            string filePath = Path.Combine(pathToSave, fileName);           
            System.IO.File.WriteAllBytes(filePath, fileBytes);
           

            return filePath;           

           
        }

        internal static string DeleteFile(string filePath)
        {
            System.IO.File.Delete(filePath);
            return "ok";
        }
    

        internal static byte[] exportpdf(DataTable dtEmployee, string EventCode, string EventName, string EventDate, string EventVenue, DataTable dtMai)
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










    }

}
