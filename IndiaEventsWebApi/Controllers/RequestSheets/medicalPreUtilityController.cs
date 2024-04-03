using IndiaEventsWebApi.Helper;
using IndiaEventsWebApi.Models.EventTypeSheets;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Globalization;
using System.Text;
using static Org.BouncyCastle.Bcpg.Attr.ImageAttrib;
//using static IndiaEventsWebApi.Models.EventTypeSheets.MedicalUtility;

namespace IndiaEventsWebApi.Controllers.RequestSheets
{
    [Route("api/[controller]")]
    [ApiController]
    //[Authorize]
    public class medicalPreUtilityController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;

        public medicalPreUtilityController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

        }

        [HttpPost("PreEventData"), DisableRequestSizeLimit]
        public IActionResult PreEventData(MedicalUtilityPreEventPayload formDataList)
        {

            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            string sheetId1 = configuration.GetSection("SmartsheetSettings:Class1").Value;
            string sheetId2 = configuration.GetSection("SmartsheetSettings:EventRequestBrandsList").Value;
            string sheetId4 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
            string sheetId6 = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
            string sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;



            long.TryParse(sheetId1, out long parsedSheetId1);
            long.TryParse(sheetId2, out long parsedSheetId2);
            long.TryParse(sheetId4, out long parsedSheetId4);
            long.TryParse(sheetId6, out long parsedSheetId6);
            long.TryParse(sheetId7, out long parsedSheetId7);

            Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId1, null, null, null, null, null, null, null);
            Sheet sheet2 = smartsheet.SheetResources.GetSheet(parsedSheetId2, null, null, null, null, null, null, null);
            Sheet sheet4 = smartsheet.SheetResources.GetSheet(parsedSheetId4, null, null, null, null, null, null, null);
            Sheet sheet6 = smartsheet.SheetResources.GetSheet(parsedSheetId6, null, null, null, null, null, null, null);
            Sheet sheet7 = smartsheet.SheetResources.GetSheet(parsedSheetId7, null, null, null, null, null, null, null);

            StringBuilder addedBrandsData = new StringBuilder();
            StringBuilder addedHcpData = new StringBuilder();
            StringBuilder addedExpences = new StringBuilder();

            int addedHcpDataNo = 1;
            int addedBrandsDataNo = 1;
            int addedExpencesNo = 1;


            var TotalExpenseAmount = 0;

            CultureInfo hindi = new CultureInfo("hi-IN");

            //var EventOpen30Days = !string.IsNullOrEmpty(formDataList.HcpConsultant.EventOpen30days) ? "Yes" : "No";
            var EventOpen30Days = !string.IsNullOrEmpty(formDataList.MedicalUtilityData.EventOpen30daysFile) ? "Yes" : "No";
            var EventWithin7Days = !string.IsNullOrEmpty(formDataList.MedicalUtilityData.EventWithin7daysFile) ? "Yes" : "No";
            var UploadDeviationFile = !string.IsNullOrEmpty(formDataList.MedicalUtilityData.UploadDeviationFile) ? "Yes" : "No";
            var FCPA = "";


            //if (formDataList.MedicalUtilityData.EventWithin7daysFile != "")
            //{
            //    EventWithin7Days = "Yes";
            //}
            //else
            //{
            //    EventWithin7Days = "No";
            //}
            //if (formDataList.MedicalUtilityData.EventOpen30daysFile != "")
            //{
            //    EventOpen30Days = "Yes";
            //}
            //else
            //{
            //    EventOpen30Days = "No";
            //}
            //if (formDataList.MedicalUtilityData.UploadDeviationFile != "")
            //{
            //    UploadDeviationFile = "Yes";
            //}
            //else
            //{
            //    UploadDeviationFile = "No";
            //}




            foreach (var formdata in formDataList.ExpenseSheet)
            {
                string rowData = $"{addedExpencesNo}. {formdata.Expense} | TotalAmount: {formdata.TotalExpenseAmount}| {formdata.BTC_BTE}";
                addedExpences.AppendLine(rowData);
                addedExpencesNo++;
                var amount = SheetHelper.NumCheck(formdata.TotalExpenseAmount);
                TotalExpenseAmount = TotalExpenseAmount + amount;

            }
            string Expense = addedExpences.ToString();



            foreach (var formdata in formDataList.BrandsList)
            {

                string rowData = $"{addedBrandsDataNo}. {formdata.BrandName} | {formdata.ProjectId} | {formdata.PercentAllocation}";
                addedBrandsData.AppendLine(rowData);
                addedBrandsDataNo++;
            }
            string brand = addedBrandsData.ToString();




            foreach (var formdata in formDataList.HcpList)
            {
                string rowData = $"{addedHcpDataNo}. {formdata.MisCode} |{formdata.HcpName} |Speciality: {formdata.Speciality} |Tier: {formdata.Tier} ";

                addedHcpData.AppendLine(rowData);
                addedHcpDataNo++;

            }
            string HCP = addedHcpData.ToString();


            //var BTE = SheetHelper.NumCheck(formDataList.MedicalUtilityData.TotalExpenseBTE);
            //var BTC = SheetHelper.NumCheck(formDataList.MedicalUtilityData.TotalExpenseBTC);

            //var total = BTC + BTE;
            var total = TotalExpenseAmount;






            try
            {

                var newRow = new Row();
                newRow.Cells = new List<Cell>();
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.MedicalUtilityData.EventTopic });

                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventType"), Value = formDataList.MedicalUtilityData.EventType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventDate"), Value = formDataList.MedicalUtilityData.EventDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Valid From"), Value = formDataList.MedicalUtilityData.ValidFrom });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Valid To"), Value = formDataList.MedicalUtilityData.ValidTill });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Utility Type"), Value = formDataList.MedicalUtilityData.MedicalUtilityType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Utility Description"), Value = formDataList.MedicalUtilityData.MedicalUtilityDescription });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "IsAdvanceRequired"), Value = formDataList.MedicalUtilityData.IsAdvanceRequired });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Advance Amount"), Value = SheetHelper.NumCheck(formDataList.MedicalUtilityData.AdvanceAmount )});
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Brands"), Value = brand });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Expenses"), Value = Expense });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Panelists"), Value = HCP });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "InitiatorName"), Value = formDataList.MedicalUtilityData.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense"), Value = TotalExpenseAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Budget"), Value = total });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Email"), Value = formDataList.MedicalUtilityData.Initiator_Email });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "RBM/BM"), Value = formDataList.MedicalUtilityData.RBMorBM });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Head"), Value = formDataList.MedicalUtilityData.Sales_Head });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Coordinator"), Value = formDataList.MedicalUtilityData.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Marketing Head"), Value = formDataList.MedicalUtilityData.Marketing_Head });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Compliance"), Value = formDataList.MedicalUtilityData.ComplianceEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Accounts"), Value = formDataList.MedicalUtilityData.FinanceAccountsEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Treasury"), Value = formDataList.MedicalUtilityData.Finance });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Reporting Manager"), Value = formDataList.MedicalUtilityData.ReportingManagerEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "1 Up Manager"), Value = formDataList.MedicalUtilityData.FirstLevelEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Affairs Head"), Value = formDataList.MedicalUtilityData.MedicalAffairsEmail });
                //newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role"), Value = formDataList.MedicalUtilityData.Role });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, " Total Expense BTC"), Value = SheetHelper.NumCheck(formDataList.MedicalUtilityData.TotalExpenseBTC) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense BTE"), Value = SheetHelper.NumCheck(formDataList.MedicalUtilityData.TotalExpenseBTE) });


                var addedRows = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId1, new Row[] { newRow });

                var eventIdColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                var eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                var val = eventIdCell.DisplayValue;


                List<string> DeviationNames = new List<string>();
                DeviationNames.Add($"30DaysDeviationFile:{formDataList.MedicalUtilityData.EventOpen30daysFile}");
                DeviationNames.Add($"7DaysDeviationFile:{formDataList.MedicalUtilityData.EventWithin7daysFile}");
                DeviationNames.Add($"AgregateSpendDeviationFile:{formDataList.MedicalUtilityData.UploadDeviationFile}");

                if (EventOpen30Days == "Yes" || EventWithin7Days == "Yes" || UploadDeviationFile == "Yes")
                {
                    var eventId = val;
                    foreach (var name in DeviationNames)
                    {
                        var y = name.Split(':');
                        var fn = y[0];
                        var bs = y[1];

                        if (bs != "")
                        {
                            try
                            {
                                var newRow7 = new Row();
                                newRow7.Cells = new List<Cell>();
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.MedicalUtilityData.EventTopic });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formDataList.MedicalUtilityData.EventType });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formDataList.MedicalUtilityData.EventDate });
                                if (fn == "30DaysDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = EventOpen30Days });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = formDataList.MedicalUtilityData.EventOpen30dayscount });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Outstanding with initiator for more than 45 days" });

                                }
                                else if (fn == "7DaysDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = EventWithin7Days });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "5 days from the Event Date" });

                                }
                                else if (fn == "AgregateSpendDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HCP exceeds 1,00,000 Trigger"), Value = UploadDeviationFile });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Aggregate Limit of 1,00,000 is Exceeded" });
                                }
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.MedicalUtilityData.Sales_Head });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = /*"sybil.oommen@menariniapac.com"*/formDataList.MedicalUtilityData.FinanceHead });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formDataList.MedicalUtilityData.InitiatorName });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.MedicalUtilityData.Initiator_Email });


                                var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId7, new Row[] { newRow7 });

                                var filename = fn;
                                var filePath = SheetHelper.testingFile(bs, val, filename);



                                var addedRow = addeddeviationrow[0];

                                var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                        parsedSheetId7, addedRow.Id.Value, filePath, "application/msword"); 
                                var attachmentintoMain = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                        parsedSheetId1, addedRows[0].Id.Value, filePath, "application/msword");

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





                foreach (var formData in formDataList.HcpList)
                {
                    var newRow1 = new Row();
                    newRow1.Cells = new List<Cell>();
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCPName"), Value = formData.HcpName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "MISCode"), Value = formData.MisCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCP Type"), Value = formData.HcpType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Speciality"), Value = formData.Speciality });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Tier"), Value = formData.Tier });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Medical Utility Cost"), Value = SheetHelper.NumCheck(formData.MedicalUtilityCostAmount) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Medical Utility Type"), Value = formDataList.MedicalUtilityData.MedicalUtilityType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Medical Utility description"), Value = formDataList.MedicalUtilityData.MedicalUtilityDescription });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Legitimate Need"), Value = formData.Legitimate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Objective Criteria"), Value = formData.Objective });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Rationale"), Value = formData.Rationale });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Topic"), Value = formDataList.MedicalUtilityData.EventTopic });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Type"), Value = formDataList.MedicalUtilityData.EventType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Date Start"), Value = formDataList.MedicalUtilityData.ValidFrom });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event End Date"), Value = formDataList.MedicalUtilityData.ValidTill });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "EventId/EventRequestId"), Value = val });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "ExpenseType"), Value = formData.ExpenseType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "FCPA Date"), Value = formData.Fcpadate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Request Date"), Value = formData.HCPRequestDate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Valid From"), Value = formDataList.MedicalUtilityData.ValidFrom });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Valid To"), Value = formDataList.MedicalUtilityData.ValidTill });

                    //Request Date,fcpa date,Expense type
                    var addeddatarows = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId4, new Row[] { newRow1 });


                    var FCPAFile = !string.IsNullOrEmpty(formData.UploadFCPA) ? "Yes" : "No";
                    var UploadWrittenRequestDate = !string.IsNullOrEmpty(formData.UploadWrittenRequestDate) ? "Yes" : "No";
                    var Invoice_Brouchere_Quotation = !string.IsNullOrEmpty(formData.Invoice_Brouchere_Quotation) ? "Yes" : "No";



                    var columnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                    var Cell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == columnId);
                    var value = Cell.DisplayValue;

                    if (FCPAFile == "Yes")
                    {

                        var filename =" FCPA";
                        var filePath = SheetHelper.testingFile(formData.UploadFCPA, val, filename);
                        var addedRow = addedRows[0];                      
                        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                parsedSheetId1, addedRow.Id.Value, filePath, "application/msword");
                        if (System.IO.File.Exists(filePath))
                        {
                            SheetHelper.DeleteFile(filePath);
                        }
                    }
                    if (UploadWrittenRequestDate == "Yes")
                    {
                        var filename = " UploadWrittenRequestDate";
                        var filePath = SheetHelper.testingFile(formData.UploadWrittenRequestDate, val, filename);
                        var addedRow = addedRows[0];
                        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                parsedSheetId1, addedRow.Id.Value, filePath, "application/msword");
                        if (System.IO.File.Exists(filePath))
                        {
                           SheetHelper.DeleteFile(filePath);
                        }
                    }
                   
                    if (Invoice_Brouchere_Quotation == "Yes")
                    {
                        var filename = " Invoice_Brouchere_Quotation";
                        var filePath = SheetHelper.testingFile(formData.Invoice_Brouchere_Quotation, val, filename);
                        var addedRow = addedRows[0];
                        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                parsedSheetId1, addedRow.Id.Value, filePath, "application/msword");
                        if (System.IO.File.Exists(filePath))
                        {
                            System.IO.File.Delete(filePath);
                        }
                    }





                }

                foreach (var formdata in formDataList.BrandsList)
                {
                    var newRow2 = new Row();
                    newRow2.Cells = new List<Cell>();
                    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "% Allocation"), Value = formdata.PercentAllocation });
                    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Brands"), Value = formdata.BrandName });
                    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Project ID"), Value = formdata.ProjectId });
                    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "EventId/EventRequestId"), Value = val });

                    smartsheet.SheetResources.RowResources.AddRows(parsedSheetId2, new Row[] { newRow2 });

                }


                foreach (var formdata in formDataList.ExpenseSheet)
                {
                    var newRow6 = new Row();
                    newRow6.Cells = new List<Cell>();
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "MisCode"), Value = formdata.MisCode });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Expense"), Value = formdata.Expense });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "EventId/EventRequestID"), Value = val });

                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formdata.BTC_BTE });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount"), Value = SheetHelper.NumCheck(formdata.TotalExpenseAmount) });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount Excluding Tax"), Value = formdata.TotalExpenseAmountExcludingTax });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTCAmount"), Value = SheetHelper.NumCheck(formdata.BtcAmount) });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTEAmount"), Value = SheetHelper.NumCheck(formdata.BteAmount )});
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Topic"), Value = formDataList.MedicalUtilityData.EventTopic });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Type"), Value = formDataList.MedicalUtilityData.EventType });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Date Start"), Value = formDataList.MedicalUtilityData.ValidFrom });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event End Date"), Value = formDataList.MedicalUtilityData.ValidTill });

                    smartsheet.SheetResources.RowResources.AddRows(parsedSheetId6, new Row[] { newRow6 });
                }


                var s = addedRows[0];
                long ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role");
                var UpdateB = new Cell { ColumnId = ColumnId, Value = formDataList.MedicalUtilityData.Role };
                Row updateRows = new Row { Id = s.Id, Cells = new Cell[] { UpdateB } };
                var cellsToUpdate = s.Cells.FirstOrDefault(c => c.ColumnId == ColumnId);
                if (cellsToUpdate != null) { cellsToUpdate.Value = formDataList.MedicalUtilityData.Role; }

                smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId1, new Row[] { updateRows });


                return Ok(new
                { Message = " Success!" });
            }

            catch (Exception ex)
            {
                return BadRequest($"Could not find {ex.Message}");
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
        //        return "doc"; // .doc format
        //    }
        //    else if (bytes.Length >= 4 && (bytes[0] == 0x50 && bytes[1] == 0x4B && bytes[2] == 0x03 && bytes[3] == 0x04))
        //    {
        //        return "docx"; // .docx format
        //    }
        //    else
        //    {
        //        return "unknown";
        //    }
        //}

        //private long GetColumnIdByName(Sheet sheet, string columnname)
        //{
        //    foreach (var column in sheet.Columns)
        //    {
        //        if (column.Title == columnname)
        //        {
        //            return column.Id.Value;
        //        }
        //    }
        //    return 0;
        //}
    }
}
