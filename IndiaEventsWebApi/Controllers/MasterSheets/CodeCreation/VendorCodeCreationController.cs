using IndiaEventsWebApi.Helper;
using IndiaEventsWebApi.Models;
using IndiaEventsWebApi.Models.MasterSheets.CodeCreation;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Text;
using static Org.BouncyCastle.Bcpg.Attr.ImageAttrib;

namespace IndiaEventsWebApi.Controllers.MasterSheets.CodeCreation
{
    [Route("api/[controller]")]
    [ApiController]
    public class VendorCodeCreationController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;

        public VendorCodeCreationController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

        }
        [HttpGet("GetVendorCodeCreationData")]
        public IActionResult GetSpeakerCodeCreationData()
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:VendorCodeCreation").Value;
                long.TryParse(sheetId, out long parsedSheetId);
                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
                List<Dictionary<string, object>> sheetData = new List<Dictionary<string, object>>();
                List<string> columnNames = new List<string>();
                foreach (Column column in sheet.Columns)
                {
                    columnNames.Add(column.Title);
                }
                foreach (Row row in sheet.Rows)
                {
                    Dictionary<string, object> rowData = new Dictionary<string, object>();
                    for (int i = 0; i < row.Cells.Count && i < columnNames.Count; i++)
                    {
                        rowData[columnNames[i]] = row.Cells[i].Value;

                    }
                    sheetData.Add(rowData);
                }
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        [HttpPost("AddVendorData")]
        public IActionResult AddVendorData(VendorCodeGeneration formData)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

                string sheetId = configuration.GetSection("SmartsheetSettings:VendorCodeCreation").Value;
                long.TryParse(sheetId, out long parsedSheetId);

                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
                string[] sheetIds = {
                //configuration.GetSection("SmartsheetSettings:HcpMaster").Value,
                //configuration.GetSection("SmartsheetSettings:HcpMaster1").Value,
                //configuration.GetSection("SmartsheetSettings:HcpMaster2").Value,
                //configuration.GetSection("SmartsheetSettings:HcpMaster3").Value,
                //configuration.GetSection("SmartsheetSettings:HcpMaster4").Value,
                //configuration.GetSection("SmartsheetSettings:ApprovedSpeakers").Value,
                //configuration.GetSection("SmartsheetSettings:ApprovedTrainers").Value,
                configuration.GetSection("SmartsheetSettings:VendorMasterSheet").Value
                };
                var mis = "";
                var sheetval = "";
                foreach (string i in sheetIds)
                {
                    long.TryParse(i, out long p);
                    Sheet sheeti = smartsheet.SheetResources.GetSheet(p, null, null, null, null, null, null, null);

                    Column misCodeColumn = sheeti.Columns.FirstOrDefault(column => column.Title == "MisCode");

                    if (misCodeColumn != null)
                    {
                        Row existingRow = sheeti.Rows.FirstOrDefault(row =>
                        row.Cells != null &&

                        row.Cells.Any(cell =>
                            cell.ColumnId == misCodeColumn.Id && cell.Value != null && cell.Value.ToString() == formData.MisCode
                        )
                        );
                        if (existingRow != null)
                        {
                            mis = formData.MisCode;
                            sheetval = sheeti.Name;


                        }
                    }
                }
                if (mis != "")
                {
                    return Ok($"MIS Code: {formData.MisCode} already exist in sheetname:{sheetval}");
                }

                else
                {

                    var IsPanCardDocument = !string.IsNullOrEmpty(formData.PanCardDocument) ? "Yes" : "No";
                    var IsChequeDocument = !string.IsNullOrEmpty(formData.ChequeDocument) ? "Yes" : "No";
                    var IsTaxResidenceCertificate = !string.IsNullOrEmpty(formData.TaxResidenceCertificate) ? "Yes" : "No";
                    //var FCPA = !string.IsNullOrEmpty(formDataList.HcpConsultant.FcpaFile) ? "Yes" : "No";


                   // var IsPanCardDocument = "";
                   // var IsChequeDocument = "";
                   // var IsTaxResidenceCertificate = "";
                    //if (formData.PanCardDocument != "")
                    //{
                    //    IsPanCardDocument = "Yes";
                    //}
                    //else
                    //{
                    //    IsPanCardDocument = "No";
                    //}
                    //if (formData.ChequeDocument != "")
                    //{
                    //    IsChequeDocument = "Yes";
                    //}
                    //else
                    //{
                    //    IsChequeDocument = "No";
                    //}
                    //if (formData.TaxResidenceCertificate != "")
                    //{
                    //    IsTaxResidenceCertificate = "Yes";
                    //}
                    //else
                    //{
                    //    IsTaxResidenceCertificate = "No";
                    //}
                    var newRow = new Row();
                    newRow.Cells = new List<Cell>();

                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Initiator Name"), Value = formData.InitiatorNameName });

                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Initiator Email"), Value = formData.InitiatorEmail });


                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "VendorAccount"), Value = formData.VendorAccount });

                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "MisCode"), Value = formData.MisCode });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "BeneficiaryName"), Value = formData.BenificiaryName });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "PanCardName"), Value = formData.PanCardName });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "PanNumber"), Value = formData.PanNumber });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "BankAccountNumber"), Value = formData.BankAccountNumber });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "IfscCode"), Value = formData.IfscCode });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Swift Code"), Value = formData.SwiftCode });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "IBN Number"), Value = formData.IbnNumber });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Email "), Value = formData.Email });


                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Cheque Document"), Value = IsChequeDocument });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Pancard Document"), Value = IsPanCardDocument });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Tax Residence Certificate"), Value = IsTaxResidenceCertificate });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance Checker"), Value = formData.FinanceChecker });

                    var addedRows = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId, new Row[] { newRow });
                    var RowId = addedRows[0].Id.Value;

                    if (IsChequeDocument == "Yes")
                    {

                        byte[] fileBytes = Convert.FromBase64String(formData.ChequeDocument);
                        var folderName = Path.Combine("Resources", "Images");
                        var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                        if (!Directory.Exists(pathToSave))
                        {
                            Directory.CreateDirectory(pathToSave);
                        }

                        string fileType = SheetHelper.GetFileType(fileBytes);
                        string fileName = " ChequeDocument." + fileType;

                        string filePath = Path.Combine(pathToSave, fileName);


                        var addedRow = addedRows[0];

                        System.IO.File.WriteAllBytes(filePath, fileBytes);
                        // string type = SheetHelper.GetContentType(fileType);
                        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                parsedSheetId, addedRow.Id.Value, filePath, "application/msword");

                        if (System.IO.File.Exists(filePath))
                        {
                            System.IO.File.Delete(filePath);
                        }
                    }
                    if (IsPanCardDocument == "Yes")
                    {
                        //var addFile = AddFile(formData.TrainerCV,RowId );
                        byte[] fileBytes = Convert.FromBase64String(formData.PanCardDocument);
                        var folderName = Path.Combine("Resources", "Images");
                        var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                        if (!Directory.Exists(pathToSave))
                        {
                            Directory.CreateDirectory(pathToSave);
                        }

                        string fileType = SheetHelper.GetFileType(fileBytes);
                        string fileName = " PanCardDocument." + fileType;
                        // string fileName = val+x + ": AttachedFile." + fileType;
                        string filePath = Path.Combine(pathToSave, fileName);


                        var addedRow = addedRows[0];

                        System.IO.File.WriteAllBytes(filePath, fileBytes);
                        // string type = SheetHelper.GetContentType(fileType);
                        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                parsedSheetId, addedRow.Id.Value, filePath, "application/msword");
                        if (System.IO.File.Exists(filePath))
                        {
                            System.IO.File.Delete(filePath);
                        }

                    }
                    if (IsTaxResidenceCertificate == "Yes")
                    {
                        //var addFile = AddFile(formData.TrainerCV,RowId );
                        byte[] fileBytes = Convert.FromBase64String(formData.TaxResidenceCertificate);
                        var folderName = Path.Combine("Resources", "Images");
                        var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                        if (!Directory.Exists(pathToSave))
                        {
                            Directory.CreateDirectory(pathToSave);
                        }

                        string fileType = SheetHelper.GetFileType(fileBytes);
                        string fileName = " TaxResidenceCertificate." + fileType;
                        // string fileName = val+x + ": AttachedFile." + fileType;
                        string filePath = Path.Combine(pathToSave, fileName);


                        var addedRow = addedRows[0];

                        System.IO.File.WriteAllBytes(filePath, fileBytes);
                        // string type = SheetHelper.GetContentType(fileType);
                        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                parsedSheetId, addedRow.Id.Value, filePath, "application/msword");
                        if (System.IO.File.Exists(filePath))
                        {
                            System.IO.File.Delete(filePath);
                        }

                    }

                }


                return Ok(new
                { Message = "Data added successfully." });

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }


        [HttpPut("UpdateVendorDatausingVendorId")]
        public IActionResult UpdateVendorDatausingVendorId(UpdateVendorCodeGeneration formData)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:VendorMasterSheet").Value;
                //string sheetId1 = configuration.GetSection("SmartsheetSettings:HonorariumPayment").Value;
                long.TryParse(sheetId, out long parsedSheetId);
                //long.TryParse(sheetId1, out long parsedSheetId1);

                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
                //Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId1, null, null, null, null, null, null, null);           

                Row existingRow = GetRowById(smartsheet, parsedSheetId, formData.VendorId);
                Row updateRow = new Row { Id = existingRow.Id, Cells = new List<Cell>() };

                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Initiator Name"), Value = formData.InitiatorNameName });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Initiator Email"), Value = formData.InitiatorEmail });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "VendorAccount"), Value = formData.VendorAccount });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "MisCode"), Value = formData.MisCode });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "BeneficiaryName"), Value = formData.BenificiaryName });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "PanCardName"), Value = formData.PanCardName });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "PanNumber"), Value = formData.PanNumber });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "BankAccountNumber"), Value = formData.BankAccountNumber });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "IfscCode"), Value = formData.IfscCode });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Swift Code"), Value = formData.SwiftCode });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "IBN Number"), Value = formData.IbnNumber });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Email "), Value = formData.Email });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Requestor Name"), Value = formData.InitiatorNameName });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Requestor"), Value = formData.InitiatorEmail });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Tax Residence Certificate Date"), Value = formData.TaxResidenceCertificateDate });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance Checker"), Value = formData.FinanceChecker });

                var updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId, new Row[] { updateRow });

                var attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(parsedSheetId, updatedRow[0].Id.Value, null);
                List<string> Names = new List<string>();
                foreach (var attachment in attachments.Data)
                {
                    var Id = attachment.Id;
                    var Fullname = attachment.Name.Split(".");
                    var splitName = Fullname[0];
                    var data = $"{splitName}:{Id}";
                    Names.Add(data);
                }

                var IsPanCardDocument = !string.IsNullOrEmpty(formData.PanCardDocument) ? "Yes" : "No";
                var IsChequeDocument = !string.IsNullOrEmpty(formData.ChequeDocument) ? "Yes" : "No";
                var IsTaxResidenceCertificate = !string.IsNullOrEmpty(formData.TaxResidenceCertificate) ? "Yes" : "No";

                if (IsChequeDocument == "Yes")
                {
                    foreach (var a in attachments.Data)
                    {
                        long Id = (long)a.Id;
                        var Fullname = a.Name.Split(".");
                        var splitName = Fullname[0];
                      
                        if(splitName == " ChequeDocument")
                        {

                            smartsheet.SheetResources.AttachmentResources.DeleteAttachment(
                              parsedSheetId,           // sheetId
                              Id            // attachmentId
                            );

                        }
                    }


                    byte[] fileBytes = Convert.FromBase64String(formData.ChequeDocument);
                    var folderName = Path.Combine("Resources", "Images");
                    var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                    if (!Directory.Exists(pathToSave))
                    {
                        Directory.CreateDirectory(pathToSave);
                    }

                    string fileType = SheetHelper.GetFileType(fileBytes);
                    string fileName = " ChequeDocument." + fileType;

                    string filePath = Path.Combine(pathToSave, fileName);


                    var addedRow = updatedRow[0];

                    System.IO.File.WriteAllBytes(filePath, fileBytes);
                    // string type = SheetHelper.GetContentType(fileType);
                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                            parsedSheetId, addedRow.Id.Value, filePath, "application/msword");

                    if (System.IO.File.Exists(filePath))
                    {
                        System.IO.File.Delete(filePath);
                    }
                }
                if (IsPanCardDocument == "Yes")
                {
                    foreach (var a in attachments.Data)
                    {
                        long Id = (long)a.Id;
                        var Fullname = a.Name.Split(".");
                        var splitName = Fullname[0];
                       
                       
                        if (splitName == " PanCardDocument")
                        {

                            smartsheet.SheetResources.AttachmentResources.DeleteAttachment(
                              parsedSheetId,           // sheetId
                              Id            // attachmentId
                            );

                        }
                    }

                    //var addFile = AddFile(formData.TrainerCV,RowId );
                    byte[] fileBytes = Convert.FromBase64String(formData.PanCardDocument);
                    var folderName = Path.Combine("Resources", "Images");
                    var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                    if (!Directory.Exists(pathToSave))
                    {
                        Directory.CreateDirectory(pathToSave);
                    }

                    string fileType = SheetHelper.GetFileType(fileBytes);
                    string fileName = " PanCardDocument." + fileType;
                    // string fileName = val+x + ": AttachedFile." + fileType;
                    string filePath = Path.Combine(pathToSave, fileName);


                    var addedRow = updatedRow[0];

                    System.IO.File.WriteAllBytes(filePath, fileBytes);
                    // string type = SheetHelper.GetContentType(fileType);
                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                            parsedSheetId, addedRow.Id.Value, filePath, "application/msword");
                    if (System.IO.File.Exists(filePath))
                    {
                        System.IO.File.Delete(filePath);
                    }

                }
                if (IsTaxResidenceCertificate == "Yes")
                {
                    foreach (var a in attachments.Data)
                    {
                        long Id = (long)a.Id;
                        var Fullname = a.Name.Split(".");
                        var splitName = Fullname[0];
                        
                      
                        if (splitName == " TaxResidenceCertificate")
                        {

                            smartsheet.SheetResources.AttachmentResources.DeleteAttachment(
                              parsedSheetId,           // sheetId
                              Id            // attachmentId
                            );

                        }
                    }
                    //var addFile = AddFile(formData.TrainerCV,RowId );
                    byte[] fileBytes = Convert.FromBase64String(formData.TaxResidenceCertificate);
                    var folderName = Path.Combine("Resources", "Images");
                    var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                    if (!Directory.Exists(pathToSave))
                    {
                        Directory.CreateDirectory(pathToSave);
                    }

                    string fileType = SheetHelper.GetFileType(fileBytes);
                    string fileName = " TaxResidenceCertificate." + fileType;
                    // string fileName = val+x + ": AttachedFile." + fileType;
                    string filePath = Path.Combine(pathToSave, fileName);


                    var addedRow = updatedRow[0];

                    System.IO.File.WriteAllBytes(filePath, fileBytes);
                    // string type = SheetHelper.GetContentType(fileType);
                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                            parsedSheetId, addedRow.Id.Value, filePath, "application/msword");
                    if (System.IO.File.Exists(filePath))
                    {
                        System.IO.File.Delete(filePath);
                    }

                }



                return Ok(new { Message = "Data Updated successfully." });

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }

        }






        [HttpGet("Getbase64UsingVendorId")]
        public IActionResult Getbase64UsingVendorId(string vendorId)
        {
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            string[] sheetIds = {
                configuration.GetSection("SmartsheetSettings:VendorMasterSheet").Value
            };
            foreach (string i in sheetIds)
            {
                long.TryParse(i, out long p);
                Sheet sheeti = smartsheet.SheetResources.GetSheet(p, null, null, null, null, null, null, null);
                Column misCodeColumn = sheeti.Columns.FirstOrDefault(column => column.Title == "VendorId");
                Column DateColumn = sheeti.Columns.FirstOrDefault(column => column.Title == "Tax Residence Certificate Date");
                if (misCodeColumn != null)
                {
                    Row existingRow = sheeti.Rows.FirstOrDefault(row =>
                        row.Cells != null &&
                        row.Cells.Any(cell =>
                            cell.ColumnId == misCodeColumn.Id && cell.Value != null && cell.Value.ToString() == vendorId
                        )
                    );
                    if (existingRow != null)
                    {
                        Cell VendorCell = existingRow.Cells.FirstOrDefault(cell => cell.ColumnId == misCodeColumn.Id);
                        Cell DateCell = existingRow.Cells.FirstOrDefault(cell => cell.ColumnId == DateColumn.Id);
                        var date = DateCell?.Value?.ToString();


                        var TaxResidenceCertificateDate = !string.IsNullOrEmpty(date) ? $"{"TaxResidenceCertificateDate"}:{date}" : "No";


                        var attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(p, existingRow.Id.Value, null);
                        var url = "";

                        var dataArray = new List<string>();
                        Dictionary<string, object> rowData = new Dictionary<string, object>();
                        List<string> Base64Strings = new List<string>();
                        foreach (var attachment in attachments.Data)
                        {
                            if (attachment != null)
                            {
                                var AID = (long)attachment.Id;
                                var Name = attachment.Name;
                                var file = smartsheet.SheetResources.AttachmentResources.GetAttachment(p, AID);
                                url = file.Url;                                
                                using (HttpClient client = new HttpClient())
                                {
                                    var fileContent = client.GetByteArrayAsync(url).Result;
                                    var base64String = Convert.ToBase64String(fileContent);
                                    Base64Strings.Add(base64String);
                                    rowData[Name] = base64String;
                                    var Data = $"{Name}:{base64String}";
                                    dataArray.Add(Data);
                                }
                            }
                        }
                        return Ok(new
                        {
                            TaxResidenceCertificateDate,
                            dataArray
                        });
                    }
                }
            }

            return Ok("False");
        }






        private Row GetRowById(SmartsheetClient smartsheet, long sheetId, string email)
        {
            Sheet sheet = smartsheet.SheetResources.GetSheet(sheetId, null, null, null, null, null, null, null);



            Column idColumn = sheet.Columns.FirstOrDefault(col => col.Title == "VendorId");

            if (idColumn != null)
            {
                foreach (var row in sheet.Rows)
                {
                    var cell = row.Cells.FirstOrDefault(c => c.ColumnId == idColumn.Id && c.Value.ToString() == email);

                    if (cell != null)
                    {
                        return row;
                    }
                }
            }

            return null;
        }


        private long GetColumnIdByName(Sheet sheet, string columnname)
        {
            foreach (var column in sheet.Columns)
            {
                if (column.Title == columnname)
                {
                    return column.Id.Value;
                }
            }
            return 0;
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
                return "doc"; // .doc format
            }
            else if (bytes.Length >= 4 && (bytes[0] == 0x50 && bytes[1] == 0x4B && bytes[2] == 0x03 && bytes[3] == 0x04))
            {
                return "docx"; // .docx format
            }
            else
            {
                return "unknown";
            }
        }
    }
}
