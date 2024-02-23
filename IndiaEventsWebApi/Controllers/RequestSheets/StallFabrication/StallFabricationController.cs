using IndiaEventsWebApi.Models.EventTypeSheets;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Globalization;
using System.Text;
using static Org.BouncyCastle.Bcpg.Attr.ImageAttrib;

namespace IndiaEventsWebApi.Controllers.RequestSheets.StallFabrication
{
    [Route("api/[controller]")]
    [ApiController]
    public class StallFabricationController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;

        public StallFabricationController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

        }

        [HttpPost("PreEventData")]
        public IActionResult PreEventData(AllStallFabrication formDataList)
        {

            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            string sheetId1 = configuration.GetSection("SmartsheetSettings:Class1").Value;
            string sheetId2 = configuration.GetSection("SmartsheetSettings:EventRequestBrandsList").Value;
            string sheetId6 = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
            string sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;
            long.TryParse(sheetId1, out long parsedSheetId1);
            long.TryParse(sheetId2, out long parsedSheetId2);
            long.TryParse(sheetId6, out long parsedSheetId6);
            long.TryParse(sheetId7, out long parsedSheetId7);
            Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId1, null, null, null, null, null, null, null);
            Sheet sheet2 = smartsheet.SheetResources.GetSheet(parsedSheetId2, null, null, null, null, null, null, null);
            Sheet sheet6 = smartsheet.SheetResources.GetSheet(parsedSheetId6, null, null, null, null, null, null, null);
            Sheet sheet7 = smartsheet.SheetResources.GetSheet(parsedSheetId7, null, null, null, null, null, null, null);
            StringBuilder addedBrandsData = new StringBuilder();
            StringBuilder addedExpences = new StringBuilder();
            int addedBrandsDataNo = 1;
            int addedExpencesNo = 1;
            var TotalExpenseAmount = 0;
            CultureInfo hindi = new CultureInfo("hi-IN");


            var uploadDeviationForTableContainsData = "";
            var EventWithin7Days = "";
            var BrouchereUpload = "";
            var InvoiceUpload = "";
            if (formDataList.StallFabrication.EventWithin7daysUpload != "")
            {
                EventWithin7Days = "Yes";
            }
            else
            {
                EventWithin7Days = "No";
            }
            if (formDataList.StallFabrication.TableContainsDataUpload != "")
            {
                uploadDeviationForTableContainsData = "Yes";
            }
            else
            {
                uploadDeviationForTableContainsData = "No";
            }
            if (formDataList.StallFabrication.EventBrouchereUpload != "")
            {
                BrouchereUpload = "Yes";
            }
            else
            {
                BrouchereUpload = "No";
            }
            if (formDataList.StallFabrication.Invoice_QuotationUpload != "")
            {
                InvoiceUpload = "Yes";
            }
            else
            {
                InvoiceUpload = "No";
            }

            foreach (var formdata in formDataList.ExpenseSheets)
            {
                string rowData = $"{addedExpencesNo}. {formdata.Expense} | AmountExcludingTax: {formdata.AmountExcludingTax}| Amount: {formdata.Amount} | {formdata.BtcorBte}";
                addedExpences.AppendLine(rowData);
                addedExpencesNo++;
                var amount = int.Parse(formdata.Amount);
                TotalExpenseAmount = TotalExpenseAmount + amount;

            }
            string Expense = addedExpences.ToString();


            foreach (var formdata in formDataList.EventBrands)
            {

                string rowData = $"{addedBrandsDataNo}. {formdata.BrandName} | {formdata.ProjectId} | {formdata.PercentAllocation}";
                addedBrandsData.AppendLine(rowData);
                addedBrandsDataNo++;
            }
            string brand = addedBrandsData.ToString();

            var FormattedTotalExpenseAmount = string.Format(hindi, "{0:#,#}", TotalExpenseAmount);

            var total = TotalExpenseAmount;

            var FormattedTotal = string.Format(hindi, "{0:#,#}", total);



            try
            {

                var newRow = new Row();
                newRow.Cells = new List<Cell>();
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.StallFabrication.EventName });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "EventType"), Value = formDataList.StallFabrication.EventType });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "EventDate"), Value = formDataList.StallFabrication.StartDate });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Event End Date"), Value = formDataList.StallFabrication.EndDate });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Class III Event Code"), Value = formDataList.StallFabrication.Class_III_EventCode });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Brands"), Value = brand });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Expenses"), Value = Expense });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "InitiatorName"), Value = formDataList.StallFabrication.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Total Expense"), Value = TotalExpenseAmount });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Total Budget"), Value = total });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "IsAdvanceRequired"), Value = formDataList.StallFabrication.IsAdvanceRequired });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Initiator Email"), Value = formDataList.StallFabrication.Initiator_Email });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "RBM/BM"), Value = formDataList.StallFabrication.RBMorBM });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Sales Head"), Value = formDataList.StallFabrication.Sales_Head });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Sales Coordinator"), Value = formDataList.StallFabrication.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Marketing Head"), Value = formDataList.StallFabrication.Marketing_Head });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Compliance"), Value = formDataList.StallFabrication.ComplianceEmail });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Finance Accounts"), Value = formDataList.StallFabrication.FinanceAccountsEmail });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Finance Treasury"), Value = formDataList.StallFabrication.Finance });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Reporting Manager"), Value = formDataList.StallFabrication.ReportingManagerEmail });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "1 Up Manager"), Value = formDataList.StallFabrication.FirstLevelEmail });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Medical Affairs Head"), Value = formDataList.StallFabrication.MedicalAffairsEmail });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Role"), Value = formDataList.StallFabrication.Role });

                var addedRows = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId1, new Row[] { newRow });
                var eventIdColumnId = GetColumnIdByName(sheet1, "EventId/EventRequestId");
                var eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                var val = eventIdCell.DisplayValue;

                var x = 1;
                if (BrouchereUpload == "Yes")
                {



                    byte[] fileBytes = Convert.FromBase64String(formDataList.StallFabrication.EventBrouchereUpload);
                    var folderName = Path.Combine("Resources", "Images");
                    var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                    if (!Directory.Exists(pathToSave))
                    {
                        Directory.CreateDirectory(pathToSave);
                    }

                    string fileType = GetFileType(fileBytes);
                    string fileName = val + "-" + x + " Brochure." + fileType;
                    // string fileName = val+x + ": AttachedFile." + fileType;
                    string filePath = Path.Combine(pathToSave, fileName);


                    var addedRow = addedRows[0];

                    System.IO.File.WriteAllBytes(filePath, fileBytes);
                    string type = GetContentType(fileType);
                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                            parsedSheetId1, addedRow.Id.Value, filePath, "application/msword");
                    x++;

                    if (System.IO.File.Exists(filePath))
                    {
                        System.IO.File.Delete(filePath);
                    }
                }
                if (InvoiceUpload == "Yes")
                {



                    byte[] fileBytes = Convert.FromBase64String(formDataList.StallFabrication.Invoice_QuotationUpload);
                    var folderName = Path.Combine("Resources", "Images");
                    var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                    if (!Directory.Exists(pathToSave))
                    {
                        Directory.CreateDirectory(pathToSave);
                    }

                    string fileType = GetFileType(fileBytes);
                    string fileName = val + "-" + x + " Invoice_Quotation." + fileType;

                    string filePath = Path.Combine(pathToSave, fileName);


                    var addedRow = addedRows[0];

                    System.IO.File.WriteAllBytes(filePath, fileBytes);
                    string type = GetContentType(fileType);
                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                            parsedSheetId1, addedRow.Id.Value, filePath, "application/msword");
                    x++;

                    if (System.IO.File.Exists(filePath))
                    {
                        System.IO.File.Delete(filePath);
                    }
                }




                if (formDataList.StallFabrication.IsDeviationUpload == "Yes")
                {
                    var eventId = val;
                    List<string> list = new List<string>
                    {
                        "30days","7days"
                    };
                    foreach (var item in list)
                    {
                        if (uploadDeviationForTableContainsData == "Yes" && item == "30days")
                        {
                            try
                            {
                                var newRow7 = new Row();
                                newRow7.Cells = new List<Cell>();
                                newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                                newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.StallFabrication.EventName });
                                newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "EventType"), Value = formDataList.StallFabrication.EventType });
                                newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "EventDate"), Value = formDataList.StallFabrication.StartDate });
                                newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "Event End Date"), Value = formDataList.StallFabrication.EndDate });
                                newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "EventOpen30days"), Value = uploadDeviationForTableContainsData });
                                newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.StallFabrication.Sales_Head });
                                newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "InitiatorName"), Value = formDataList.StallFabrication.InitiatorName });
                                newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.StallFabrication.Initiator_Email });
                                newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "Deviation Type"), Value = "Outstanding with intiator for more than 30 days" });
                                newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "Outstanding Events"), Value = formDataList.StallFabrication.EventOpen30dayscount });

                                var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId7, new Row[] { newRow7 });

                                byte[] fileBytes = Convert.FromBase64String(formDataList.StallFabrication.TableContainsDataUpload);
                                var folderName = Path.Combine("Resources", "Images");
                                var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                                if (!Directory.Exists(pathToSave))
                                {
                                    Directory.CreateDirectory(pathToSave);
                                }
                                string fileType = GetFileType(fileBytes);
                                string fileName = "30DaysDeviationFile." + fileType;
                                // string fileName = val+x + ": AttachedFile." + fileType;
                                string filePath = Path.Combine(pathToSave, fileName);
                                var addedRow = addeddeviationrow[0];
                                System.IO.File.WriteAllBytes(filePath, fileBytes);
                                string type = GetContentType(fileType);
                                var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                        parsedSheetId7, addedRow.Id.Value, filePath, "application/msword");

                                if (System.IO.File.Exists(filePath))
                                {
                                    System.IO.File.Delete(filePath);
                                }

                            }
                            catch (Exception ex)
                            {
                                return BadRequest(ex.Message);
                            }
                        }
                        else if (EventWithin7Days == "Yes" && item == "7days")
                        {
                            try
                            {

                                var newRow7 = new Row();
                                newRow7.Cells = new List<Cell>();
                                newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                                newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.StallFabrication.EventName });
                                newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "EventType"), Value = formDataList.StallFabrication.EventType });
                                newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "EventDate"), Value = formDataList.StallFabrication.StartDate });
                                //newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "StartTime"), Value = formDataList.StallFabrication.StartTime });
                                newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "Event End Date"), Value = formDataList.StallFabrication.EndDate });
                                //newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "EventOpen30days"), Value = uploadDeviationForTableContainsData });
                                newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "EventWithin7days"), Value = EventWithin7Days });
                                newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.StallFabrication.Sales_Head });
                                newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "InitiatorName"), Value = formDataList.StallFabrication.InitiatorName });
                                newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.StallFabrication.Initiator_Email });
                                newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "Deviation Type"), Value = "7 days from the Event Date" });


                                var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId7, new Row[] { newRow7 });

                                byte[] fileBytes = Convert.FromBase64String(formDataList.StallFabrication.EventWithin7daysUpload);
                                var folderName = Path.Combine("Resources", "Images");
                                var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                                if (!Directory.Exists(pathToSave))
                                {
                                    Directory.CreateDirectory(pathToSave);
                                }
                                string fileType = GetFileType(fileBytes);
                                string fileName = "7DaysDeviationFile." + fileType;
                                string filePath = Path.Combine(pathToSave, fileName);
                                var addedRow = addeddeviationrow[0];
                                System.IO.File.WriteAllBytes(filePath, fileBytes);
                                string type = GetContentType(fileType);
                                var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                        parsedSheetId7, addedRow.Id.Value, filePath, "application/msword");

                                if (System.IO.File.Exists(filePath))
                                {
                                    System.IO.File.Delete(filePath);
                                }

                            }
                            catch (Exception ex)
                            {
                                return BadRequest(ex.Message);
                            }
                        }
                    }
                  
                }


                foreach (var formdata in formDataList.EventBrands)
                {
                    var newRow2 = new Row();
                    newRow2.Cells = new List<Cell>();
                    newRow2.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet2, "% Allocation"), Value = formdata.PercentAllocation });
                    newRow2.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet2, "Brands"), Value = formdata.BrandName });
                    newRow2.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet2, "Project ID"), Value = formdata.ProjectId });
                    newRow2.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet2, "EventId/EventRequestId"), Value = val });

                    smartsheet.SheetResources.RowResources.AddRows(parsedSheetId2, new Row[] { newRow2 });

                }


                foreach (var formdata in formDataList.ExpenseSheets)
                {
                    var newRow6 = new Row();
                    newRow6.Cells = new List<Cell>();

                    newRow6.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet6, "Expense"), Value = formdata.Expense });
                    newRow6.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet6, "EventId/EventRequestID"), Value = val });
                    newRow6.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet6, "AmountExcludingTax?"), Value = formdata.AmountExcludingTax });
                    newRow6.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet6, "Amount"), Value = formdata.Amount });
                    newRow6.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet6, "BTC/BTE"), Value = formdata.BtcorBte });
                    newRow6.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet6, "BudgetAmount"), Value = formdata.BudgetAmount });
                    newRow6.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet6, "BTCAmount"), Value = formdata.BtcAmount });
                    newRow6.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet6, "BTEAmount"), Value = formdata.BteAmount });

                    smartsheet.SheetResources.RowResources.AddRows(parsedSheetId6, new Row[] { newRow6 });
                }

                return Ok(new
                { Message = " Success!" });



            }



            catch (Exception ex)
            {
                return BadRequest($"Could not find {ex.Message}");
            }












        }


        private string GetContentType(string fileExtension)
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
        private Row GetRowById(SmartsheetClient smartsheet, long sheetId, string val)
        {
            Sheet sheet = smartsheet.SheetResources.GetSheet(sheetId, null, null, null, null, null, null, null);

            // Assuming you have a column named "Id"

            Column idColumn = sheet.Columns.FirstOrDefault(col => col.Title == "Honorarium Submitted?");

            if (idColumn != null)
            {
                foreach (var row in sheet.Rows)
                {
                    var cell = row.Cells.FirstOrDefault(c => c.ColumnId == idColumn.Id && c.Value.ToString() == val);

                    if (cell != null)
                    {
                        return row;
                    }
                }
            }

            return null;
        }
    }
}
