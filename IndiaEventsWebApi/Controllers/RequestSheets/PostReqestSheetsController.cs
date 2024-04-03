using Microsoft.AspNetCore.Components.Forms;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using IndiaEventsWebApi.Models;
using IndiaEventsWebApi.Models.EventTypeSheets;
using IndiaEventsWebApi.Models.RequestSheets;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using Smartsheet.Api.OAuth;
using System.Text;
using System.Globalization;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using static Org.BouncyCastle.Bcpg.Attr.ImageAttrib;
using IndiaEventsWebApi.Helper;
using IndiaEvents.Models.Models.Webhook;
using Smartsheet.Core.Definitions;
using Microsoft.AspNetCore.Authorization;



namespace IndiaEventsWebApi.Controllers.RequestSheets
{
    [Route("api/[controller]")]
    [ApiController]
    //[Authorize]
    public class PostReqestSheetsController : ControllerBase
    {

        private readonly string accessToken;
        private readonly IConfiguration configuration;
        private readonly SmartsheetClient smartsheet;

        public PostReqestSheetsController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
            smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
        }


        [HttpPost("AllObjModelsData"), DisableRequestSizeLimit]
        public IActionResult AllObjModelsData(AllObjModels formDataList)
        {


            string sheetId1 = configuration.GetSection("SmartsheetSettings:Class1").Value;
            string sheetId2 = configuration.GetSection("SmartsheetSettings:EventRequestBrandsList").Value;
            string sheetId3 = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
            string sheetId4 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
            string sheetId5 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpSlideKit").Value;
            string sheetId6 = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
            string sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;

            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
            Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
            Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
            Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
            Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
            Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);


            long.TryParse(sheetId1, out long parsedSheetId1);
            long.TryParse(sheetId2, out long parsedSheetId2);
            long.TryParse(sheetId3, out long parsedSheetId3);
            long.TryParse(sheetId4, out long parsedSheetId4);
            long.TryParse(sheetId5, out long parsedSheetId5);
            long.TryParse(sheetId6, out long parsedSheetId6);
            long.TryParse(sheetId7, out long parsedSheetId7);




            StringBuilder addedBrandsData = new StringBuilder();
            StringBuilder addedInviteesData = new StringBuilder();
            StringBuilder addedMEnariniInviteesData = new StringBuilder();
            StringBuilder addedHcpData = new StringBuilder();
            StringBuilder addedSlideKitData = new StringBuilder();
            StringBuilder addedExpences = new StringBuilder();

            int addedSlideKitDataNo = 1;
            int addedHcpDataNo = 1;
            int addedInviteesDataNo = 1;
            int addedInviteesDataNoforMenarini = 1;
            int addedBrandsDataNo = 1;
            int addedExpencesNo = 1;

            var TotalHonorariumAmount = 0;
            var TotalTravelAmount = 0;
            var TotalAccomodateAmount = 0;
            var TotalHCPLcAmount = 0;
            var TotalInviteesLcAmount = 0;
            var TotalExpenseAmount = 0;

            CultureInfo hindi = new CultureInfo("hi-IN");





            foreach (var formdata in formDataList.EventRequestExpenseSheet)
            {
                string rowData = $"{addedExpencesNo}. {formdata.Expense} | AmountExcludingTax: {formdata.AmountExcludingTax}| Amount: {formdata.Amount} | {formdata.BtcorBte}";
                addedExpences.AppendLine(rowData);
                addedExpencesNo++;
                var amount = SheetHelper.NumCheck(formdata.Amount);
                TotalExpenseAmount = TotalExpenseAmount + amount;

            }
            string Expense = addedExpences.ToString();

            foreach (var formdata in formDataList.EventRequestHCPSlideKits)
            {

                string rowData = $"{addedSlideKitDataNo}. {formdata.HcpName} | {formdata.SlideKitType}";
                addedSlideKitData.AppendLine(rowData);
                addedSlideKitDataNo++;
            }
            string slideKit = addedSlideKitData.ToString();

            foreach (var formdata in formDataList.RequestBrandsList)
            {

                string rowData = $"{addedBrandsDataNo}. {formdata.BrandName} | {formdata.ProjectId} | {formdata.PercentAllocation}";
                addedBrandsData.AppendLine(rowData);
                addedBrandsDataNo++;
            }
            string brand = addedBrandsData.ToString();

            foreach (var formdata in formDataList.EventRequestInvitees)
            {




                // string rowData = $"{addedInviteesDataNo}. Name: {formdata.InviteeName} | MIS Code: {formdata.MISCode} | LocalConveyance: {formdata.LocalConveyance} ";
                if (formdata.InviteedFrom == "Menarini Employees")
                {
                    string row = $"{addedInviteesDataNoforMenarini}. {formdata.InviteeName}";
                    addedMEnariniInviteesData.AppendLine(row);
                    addedInviteesDataNoforMenarini++;
                }
                else
                {
                    string rowData = $"{addedInviteesDataNo}. {formdata.InviteeName}";
                    addedInviteesData.AppendLine(rowData);
                    addedInviteesDataNo++;
                }

                TotalInviteesLcAmount = TotalInviteesLcAmount + SheetHelper.NumCheck(formdata.LcAmount);
            }
            string Invitees = addedInviteesData.ToString();
            string MenariniInvitees = addedMEnariniInviteesData.ToString();


            foreach (var formdata in formDataList.EventRequestHcpRole)
            {

                var HM = SheetHelper.NumCheck(formdata.HonarariumAmount);
                var x = string.Format(hindi, "{0:#,#}", HM);
                var t = SheetHelper.NumCheck(formdata.Travel) + SheetHelper.NumCheck(formdata.Accomdation);
                var y = string.Format(hindi, "{0:#,#}", t);
                //string rowData = $"{addedHcpDataNo}. {formdata.HcpRole} |Name: {formdata.HcpName} | Honr.Amt: {formdata.HonarariumAmount} |Trav.Amt: {formdata.Travel} |Acco.Amt: {formdata.Accomdation} ";
                string rowData = $"{addedHcpDataNo}. {formdata.HcpRole} |{formdata.HcpName} | Honr.Amt: {x} |Trav.&Acc.Amt: {y} ";

                addedHcpData.AppendLine(rowData);
                addedHcpDataNo++;
                TotalHonorariumAmount = TotalHonorariumAmount + SheetHelper.NumCheck(formdata.HonarariumAmount);
                TotalTravelAmount = TotalTravelAmount + SheetHelper.NumCheck(formdata.Travel);
                TotalAccomodateAmount = TotalAccomodateAmount + SheetHelper.NumCheck(formdata.Accomdation);
                TotalHCPLcAmount = TotalHCPLcAmount + SheetHelper.NumCheck(formdata.LocalConveyance);
            }
            string HCP = addedHcpData.ToString();



            var FormattedTotalHonorariumAmount = string.Format(hindi, "{0:#,#}", TotalHonorariumAmount);
            var FormattedTotalTravelAmount = string.Format(hindi, "{0:#,#}", TotalTravelAmount);
            var FormattedTotalAccomodateAmount = string.Format(hindi, "{0:#,#}", TotalAccomodateAmount);
            var FormattedTotalHCPLcAmount = string.Format(hindi, "{0:#,#}", TotalHCPLcAmount);
            var FornattedTotalInviteesLcAmount = string.Format(hindi, "{0:#,#}", TotalInviteesLcAmount);
            var FormattedTotalExpenseAmount = string.Format(hindi, "{0:#,#}", TotalExpenseAmount);

            var c = TotalHCPLcAmount + TotalInviteesLcAmount;
            var FormattedTotalLC = string.Format(hindi, "{0:#,#}", c);


            //var BTE = SheetHelper.NumCheck(formDataList.class1.TotalExpenseBTE);
            //var BTC = SheetHelper.NumCheck(formDataList.class1.TotalExpenseBTC);
            //var total = BTC + BTE;
            var total = TotalHonorariumAmount + TotalTravelAmount + TotalAccomodateAmount + TotalHCPLcAmount + TotalInviteesLcAmount + TotalExpenseAmount;

            var FormattedTotal = string.Format(hindi, "{0:#,#}", total);
            var s = (TotalTravelAmount + TotalAccomodateAmount);
            var FormattedTotalTAAmount = string.Format(hindi, "{0:#,#}", s);




            try
            {

                var newRow = new Row();
                newRow.Cells = new List<Cell>();
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.class1.EventTopic });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "StartTime"), Value = formDataList.class1.StartTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EndTime"), Value = formDataList.class1.EndTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "VenueName"), Value = formDataList.class1.VenueName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "City"), Value = formDataList.class1.City });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "State"), Value = formDataList.class1.State });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Panelists"), Value = HCP });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Invitees"), Value = Invitees });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "MIPL Invitees"), Value = MenariniInvitees });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventType"), Value = formDataList.class1.EventType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventDate"), Value = formDataList.class1.EventDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Brands"), Value = brand });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Expenses"), Value = Expense });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "SlideKits"), Value = slideKit });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "IsAdvanceRequired"), Value = formDataList.class1.IsAdvanceRequired });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventOpen30days"), Value = formDataList.class1.EventOpen30days });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventWithin7days"), Value = formDataList.class1.EventWithin7days });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "InitiatorName"), Value = formDataList.class1.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Advance Amount"), Value = SheetHelper.NumCheck(formDataList.class1.AdvanceAmount) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, " Total Expense BTC"), Value = SheetHelper.NumCheck(formDataList.class1.TotalExpenseBTC) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense BTE"), Value = SheetHelper.NumCheck(formDataList.class1.TotalExpenseBTE) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Honorarium Amount"), Value = TotalHonorariumAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel Amount"), Value = TotalTravelAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel & Accommodation Amount"), Value = s });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Accomodation Amount"), Value = TotalAccomodateAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Budget"), Value = total });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Local Conveyance"), Value = c });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense"), Value = TotalExpenseAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Email"), Value = formDataList.class1.Initiator_Email });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "RBM/BM"), Value = formDataList.class1.RBMorBM });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Head"), Value = formDataList.class1.Sales_Head });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Coordinator"), Value = formDataList.class1.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Marketing Head"), Value = formDataList.class1.Marketing_Head });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Treasury"), Value = formDataList.class1.Finance });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Compliance"), Value = formDataList.class1.ComplianceEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Accounts"), Value = formDataList.class1.FinanceAccountsEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Reporting Manager"), Value = formDataList.class1.ReportingManagerEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "1 Up Manager"), Value = formDataList.class1.FirstLevelEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Affairs Head"), Value = formDataList.class1.MedicalAffairsEmail });
                //newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role"), Value = formDataList.class1.Role });


                var addedRows = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId1, new Row[] { newRow });

                var eventIdColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                var eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                var val = eventIdCell.DisplayValue;

                var x = 1;
                foreach (var p in formDataList.class1.Files)
                {
                    string[] words = p.Split(':');
                    var r = words[0];
                    var q = words[1];
                    var name = r.Split(".")[0];
                    var filePath = SheetHelper.testingFile(q, val, name);
                    var addedRow = addedRows[0];
                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                           parsedSheetId1, addedRow.Id.Value, filePath, "application/msword");
                    x++;

                    // ////////////////////////////////////////////////
                    if (System.IO.File.Exists(filePath))
                    {
                        SheetHelper.DeleteFile(filePath);
                    }
                }







                if (formDataList.class1.EventOpen30days == "Yes" || formDataList.class1.EventWithin7days == "Yes" || formDataList.class1.FB_Expense_Excluding_Tax == "Yes" || formDataList.class1.IsDeviationUpload == "Yes")
                {
                    List<string> DeviationNames = new List<string>();
                    foreach (var p in formDataList.class1.DeviationFiles)
                    {
                        string[] words = p.Split(':');
                        var r = words[0];

                        DeviationNames.Add(r);
                    }
                    foreach (var deviationname in DeviationNames)
                    {
                        var file = deviationname.Split(".")[0];
                        var eventId = val;
                        try
                        {
                            var newRow7 = new Row();
                            newRow7.Cells = new List<Cell>();

                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.class1.EventTopic });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formDataList.class1.EventType });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formDataList.class1.EventDate });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "StartTime"), Value = formDataList.class1.StartTime });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EndTime"), Value = formDataList.class1.EndTime });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "VenueName"), Value = formDataList.class1.VenueName });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "City"), Value = formDataList.class1.City });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "State"), Value = formDataList.class1.State });


                            if (file == "30DaysDeviationFile")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Outstanding with initiator for more than 45 days" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = formDataList.class1.EventOpen30days });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = SheetHelper.NumCheck(formDataList.class1.EventOpen30dayscount) });
                            }
                            else if (file == "7DaysDeviationFile")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "5 days from the Event Date" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = formDataList.class1.EventWithin7days });

                            }
                            else if (file == "ExpenseExcludingTax")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Food and Beverages expense exceeds 1500" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "PRE-F&B Expense Excluding Tax"), Value = formDataList.class1.FB_Expense_Excluding_Tax });
                            }
                            else if (file == "Travel_Accomodation3LExceededFile")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Travel/AccomodationAggregate Limit of 3,00,000 is Exceeded" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Travel/Accomodation 3,00,000 Exceeded Trigger"), Value = "Yes" });//formDataList.class1.FB_Expense_Excluding_Tax });
                            }
                            else if (file == "TrainerHonorarium12LExceededFile")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Honorarium Aggregate Limit of 12,00,000 is Exceeded" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Trainer Honorarium 12,00,000 Exceeded Trigger"), Value = "Yes" }); //formDataList.class1.FB_Expense_Excluding_Tax });
                            }
                            else if (file == "HCPHonorarium6LExceededFile")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Honorarium Aggregate Limit of 6,00,000 is Exceeded" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HCP Honorarium 6,00,000 Exceeded Trigger"), Value = "Yes" }); // formDataList.class1.FB_Expense_Excluding_Tax });
                            }
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.class1.Sales_Head });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formDataList.class1.FinanceHead });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formDataList.class1.InitiatorName });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.class1.Initiator_Email });


                            var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId7, new Row[] { newRow7 });




                            var j = 1;
                            foreach (var p in formDataList.class1.DeviationFiles)
                            {
                                string[] words = p.Split(':');
                                var r = words[0];
                                var q = words[1];
                                if (deviationname == r)
                                {

                                    var name = r.Split(".")[0];

                                    var filePath = SheetHelper.testingFile(q, val, name);


                                    //byte[] fileBytes = Convert.FromBase64String(q);
                                    //var folderName = Path.Combine("Resources", "Images");
                                    //var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                                    //if (!Directory.Exists(pathToSave))
                                    //{
                                    //    Directory.CreateDirectory(pathToSave);
                                    //}

                                    //string fileType = SheetHelper.GetFileType(fileBytes);
                                    //string fileName = r;
                                    //// string fileName = val+x + ": AttachedFile." + fileType;
                                    //string filePath = Path.Combine(pathToSave, fileName);


                                    var addedRow = addeddeviationrow[0];

                                    //System.IO.File.WriteAllBytes(filePath, fileBytes);
                                    //string type = SheetHelper.GetContentType(fileType);
                                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                            parsedSheetId7, addedRow.Id.Value, filePath, "application/msword");
                                    var attachmentinmain = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                           parsedSheetId1, addedRows[0].Id.Value, filePath, "application/msword");
                                    j++;
                                    if (System.IO.File.Exists(filePath))
                                    {
                                        SheetHelper.DeleteFile(filePath);
                                    }
                                }


                            }
                        }
                        catch (Exception ex)
                        {
                            return BadRequest(ex.Message);
                        }
                    }

                }



                foreach (var formData in formDataList.EventRequestHcpRole)
                {

                    var newRow1 = new Row();
                    newRow1.Cells = new List<Cell>();
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HcpRole"), Value = formData.HcpRole });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "MISCode"), Value = formData.MisCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Travel"), Value = SheetHelper.NumCheck(formData.Travel) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TotalSpend"), Value = SheetHelper.NumCheck(formData.FinalAmount) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Accomodation"), Value = SheetHelper.NumCheck(formData.Accomdation) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "LocalConveyance"), Value = SheetHelper.NumCheck(formData.LocalConveyance) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "SpeakerCode"), Value = formData.SpeakerCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TrainerCode"), Value = formData.TrainerCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HonorariumRequired"), Value = formData.HonorariumRequired });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "AgreementAmount"), Value = SheetHelper.NumCheck(formData.AgreementAmount) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Speciality"), Value = formData.Speciality });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Topic"), Value = formDataList.class1.EventTopic });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Type"), Value = formDataList.class1.EventType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Venue name"), Value = formDataList.class1.VenueName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Date Start"), Value = formDataList.class1.EventDate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event End Date"), Value = formDataList.class1.EventEndDate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCPName"), Value = formData.HcpName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PAN card name"), Value = formData.PanCardName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "ExpenseType"), Value = formData.ExpenseType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Bank Account Number"), Value = formData.BankAccountNumber });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Bank Name"), Value = formData.BankName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "IFSC Code"), Value = formData.IFSCCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "FCPA Date"), Value = formData.Fcpadate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Currency"), Value = formData.Currency });

                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Honorarium Amount Excluding Tax"), Value = formData.HonarariumAmountExcludingTax });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Travel Excluding Tax"), Value = formData.TravelExcludingTax });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Accomodation Excluding Tax"), Value = formData.AccomdationExcludingTax });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Local Conveyance Excluding Tax"), Value = formData.LocalConveyanceExcludingTax });

                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "LC BTC/BTE"), Value = formData.LcBtcorBte });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Travel BTC/BTE"), Value = formData.TravelBtcorBte });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Accomodation BTC/BTE"), Value = formData.AccomodationBtcorBte });


                    if (formData.Currency == "Others")
                    {
                        newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Other Currency"), Value = formData.OtherCurrencyType });
                    }
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Beneficiary Name"), Value = formData.BeneficiaryName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Pan Number"), Value = formData.PanNumber });



                    if (formData.HcpRole == "Others")
                    {

                        newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Other Type"), Value = formData.OthersType });
                    }

                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Tier"), Value = formData.Tier });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCP Type"), Value = formData.GOorNGO });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PresentationDuration"), Value = SheetHelper.NumCheck(formData.PresentationDuration) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PanelSessionPreparationDuration"), Value = SheetHelper.NumCheck(formData.PanelSessionPreperationDuration) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PanelDiscussionDuration"), Value = SheetHelper.NumCheck(formData.PanelDisscussionDuration) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "QASessionDuration"), Value = SheetHelper.NumCheck(formData.QASessionDuration) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "BriefingSession"), Value = SheetHelper.NumCheck(formData.BriefingSession) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TotalSessionHours"), Value = SheetHelper.NumCheck(formData.TotalSessionHours) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Rationale"), Value = formData.Rationale });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "EventId/EventRequestId"), Value = val });
                    //newRow1.Cells.Add(new Cell
                    //{
                    //    ColumnId = SheetHelper.SheetHelper.GetColumnIdByName(sheet4, "HCPName"),
                    //    Value = formData.HcpName
                    //});
                    // ///////////////////////////////////////////////////


                    smartsheet.SheetResources.RowResources.AddRows(parsedSheetId4, new Row[] { newRow1 });







                }

                foreach (var formdata in formDataList.RequestBrandsList)
                {
                    var newRow2 = new Row();
                    newRow2.Cells = new List<Cell>();
                    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "% Allocation"), Value = formdata.PercentAllocation });
                    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Brands"), Value = formdata.BrandName });
                    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Project ID"), Value = formdata.ProjectId });
                    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "EventId/EventRequestId"), Value = val });

                    smartsheet.SheetResources.RowResources.AddRows(parsedSheetId2, new Row[] { newRow2 });

                }

                foreach (var formdata in formDataList.EventRequestInvitees)
                {
                    var newRow3 = new Row();
                    newRow3.Cells = new List<Cell>();
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "HCPName"), Value = formdata.InviteeName });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Designation"), Value = formdata.Designation });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Employee Code"), Value = formdata.EmployeeCode });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "LocalConveyance"), Value = formdata.LocalConveyance });

                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "BTC/BTE"), Value = formdata.BtcorBte });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "LcAmount"), Value = SheetHelper.NumCheck(formdata.LcAmount) });

                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Lc Amount Excluding Tax"), Value = formdata.LcAmountExcludingTax });

                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "EventId/EventRequestId"), Value = val });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Invitee Source"), Value = formdata.InviteedFrom });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "HCP Type"), Value = formdata.HCPType });
                    if (formdata.InviteedFrom == "Others")
                    {
                        newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Speciality"), Value = formdata.Speciality });


                    }
                    else
                    {
                        newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "MISCode"), Value = formdata.MISCode });
                    }

                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Topic"), Value = formDataList.class1.EventTopic });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Type"), Value = formDataList.class1.EventType });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Venue name"), Value = formDataList.class1.VenueName });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Date Start"), Value = formDataList.class1.EventDate });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event End Date"), Value = formDataList.class1.EventDate });

                    smartsheet.SheetResources.RowResources.AddRows(parsedSheetId3, new Row[] { newRow3 });
                }


                foreach (var formdata in formDataList.EventRequestHCPSlideKits)
                {
                    var newRow5 = new Row();
                    newRow5.Cells = new List<Cell>();

                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "MIS"), Value = formdata.MIS });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "Slide Kit Type"), Value = formdata.SlideKitType });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "SlideKit Document"), Value = formdata.SlideKitDocument });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "EventId/EventRequestId"), Value = val });


                    smartsheet.SheetResources.RowResources.AddRows(parsedSheetId5, new Row[] { newRow5 });
                }

                foreach (var formdata in formDataList.EventRequestExpenseSheet)
                {
                    var newRow6 = new Row();
                    newRow6.Cells = new List<Cell>();

                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Expense"), Value = formdata.Expense });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "EventId/EventRequestID"), Value = val });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "AmountExcludingTax?"), Value = formdata.AmountExcludingTax });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount Excluding Tax"), Value = formdata.ExcludingTaxAmount });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount"), Value = SheetHelper.NumCheck(formdata.Amount) });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formdata.BtcorBte });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BudgetAmount"), Value = SheetHelper.NumCheck(formdata.BudgetAmount) });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTCAmount"), Value = SheetHelper.NumCheck(formdata.BtcAmount) });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTEAmount"), Value = SheetHelper.NumCheck(formdata.BteAmount) });

                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Topic"), Value = formDataList.class1.EventTopic });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Type"), Value = formDataList.class1.EventType });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Venue name"), Value = formDataList.class1.VenueName });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Date Start"), Value = formDataList.class1.EventDate });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event End Date"), Value = formDataList.class1.EventDate });

                    smartsheet.SheetResources.RowResources.AddRows(parsedSheetId6, new Row[] { newRow6 });
                }

                var targetRow = addedRows[0];
                long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role");
                var cellToUpdateB = new Cell { ColumnId = honorariumSubmittedColumnId, Value = formDataList.class1.Role };
                Row updateRow = new Row { Id = targetRow.Id, Cells = new Cell[] { cellToUpdateB } };
                var cellToUpdate = targetRow.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                if (cellToUpdate != null) { cellToUpdate.Value = formDataList.class1.Role; }

                smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId1, new Row[] { updateRow });

                return Ok(new
                { Message = " Success!" });

            }
            catch (Exception ex)
            {
                return BadRequest($"Could not find {ex.Message}");
            }

        }


        [HttpPost("AddHonorariumData"), DisableRequestSizeLimit]
        public IActionResult AddHonorariumData(HonorariumPaymentList formData)
        {
            try
            {

                string sheetId = configuration.GetSection("SmartsheetSettings:HonorariumPayment").Value;
                string sheetId1 = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;
                string sheetId2 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;
                string sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
                Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
                Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);

                long.TryParse(sheetId, out long parsedSheetId);
                long.TryParse(sheetId1, out long parsedSheetId1);
                long.TryParse(sheetId2, out long parsedSheetId2);
                long.TryParse(sheetId7, out long parsedSheetId7);


                StringBuilder addedHcpData = new StringBuilder();

                int addedHcpDataNo = 1;



                CultureInfo hindi = new CultureInfo("hi-IN");
                foreach (var i in formData.HcpRoles)
                {

                    string rowData = $"{addedHcpDataNo}. Name:{i.HcpName} | Role:{i.HcpRole} |Code:{i.MisCode} | HCP Type:{i.GOorNGO}| Including GST:{i.IsInclidingGst}| Agreement Amount:{i.AgreementAmount} ";

                    addedHcpData.AppendLine(rowData);
                    addedHcpDataNo++;
                }
                string panalist = addedHcpData.ToString();





                var newRow = new Row();
                newRow.Cells = new List<Cell>();


                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "SlideKits"), Value = formData.RequestHonorariumList.slideKits });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventId/EventRequestId"), Value = formData.RequestHonorariumList.EventId });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Event Type"), Value = formData.RequestHonorariumList.EventType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Event Date"), Value = formData.RequestHonorariumList.EventDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Event Topic"), Value = formData.RequestHonorariumList.EventTopic });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "City"), Value = formData.RequestHonorariumList.City });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "State"), Value = formData.RequestHonorariumList.State });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Start Time"), Value = formData.RequestHonorariumList.StartTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "End Time"), Value = formData.RequestHonorariumList.EndTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Venue Name"), Value = formData.RequestHonorariumList.VenueName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Travel & Accommodation Amount"), Value = SheetHelper.NumCheck(formData.RequestHonorariumList.TotalTravelAndAccomodationSpend) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Honorarium Amount"), Value = SheetHelper.NumCheck(formData.RequestHonorariumList.TotalHonorariumSpend) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Budget"), Value = SheetHelper.NumCheck(formData.RequestHonorariumList.TotalSpend) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Expenses"), Value = formData.RequestHonorariumList.Expenses });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Travel Amount"), Value = SheetHelper.NumCheck(formData.RequestHonorariumList.TotalTravelSpend) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Accommodation Amount"), Value = SheetHelper.NumCheck(formData.RequestHonorariumList.TotalAccomodationSpend) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Expense"), Value = SheetHelper.NumCheck(formData.RequestHonorariumList.TotalExpenses) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Local Conveyance"), Value = SheetHelper.NumCheck(formData.RequestHonorariumList.TotalLocalConveyance) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Brands"), Value = formData.RequestHonorariumList.Brands });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Invitees"), Value = formData.RequestHonorariumList.Invitees });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Panelists"), Value = formData.RequestHonorariumList.Panelists });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Initiator Name"), Value = formData.RequestHonorariumList.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Panelists & Agreements"), Value = panalist });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Honorarium Submitted?"), Value = formData.RequestHonorariumList.HonarariumSubmitted });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Initiator Email"), Value = formData.RequestHonorariumList.InitiatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "RBM/BM"), Value = formData.RequestHonorariumList.RBMorBM });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Sales Head"), Value = formData.RequestHonorariumList.SalesHeadEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Sales Coordinator"), Value = formData.RequestHonorariumList.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Marketing Head"), Value = formData.RequestHonorariumList.MarketingHeadEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Compliance"), Value = formData.RequestHonorariumList.Compliance });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance Accounts"), Value = formData.RequestHonorariumList.FinanceAccounts });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance Treasury"), Value = formData.RequestHonorariumList.FinanceTreasury });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Reporting Manager"), Value = formData.RequestHonorariumList.ReportingManagerEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "1 Up Manager"), Value = formData.RequestHonorariumList.FirstLevelEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Medical Affairs Head"), Value = formData.RequestHonorariumList.MedicalAffairsEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Role"), Value = formData.RequestHonorariumList.Role });









                var addedRows = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId, new Row[] { newRow });
                var eventId = formData.RequestHonorariumList.EventId;
                var x = 1;
                foreach (var p in formData.RequestHonorariumList.Files)
                {

                    string[] words = p.Split(':');
                    var r = words[0];
                    var q = words[1];

                    var name = r.Split(".")[0];

                    var filePath = SheetHelper.testingFile(q, eventId, name);



                    var addedRow = addedRows[0];

                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                            parsedSheetId, addedRow.Id.Value, filePath, "application/msword");
                    x++;


                    if (System.IO.File.Exists(filePath))
                    {
                        SheetHelper.DeleteFile(filePath);
                    }
                }



                if (formData.RequestHonorariumList.IsDeviationUpload == "Yes")
                {

                    try
                    {

                        var newRow7 = new Row();
                        newRow7.Cells = new List<Cell>();

                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formData.RequestHonorariumList.EventTopic });

                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formData.RequestHonorariumList.EventType });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formData.RequestHonorariumList.EventDate });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "StartTime"), Value = formData.RequestHonorariumList.StartTime });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EndTime"), Value = formData.RequestHonorariumList.EndTime });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "VenueName"), Value = formData.RequestHonorariumList.VenueName });

                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "City"), Value = formData.RequestHonorariumList.City });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "State"), Value = formData.RequestHonorariumList.State });

                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Honorarium documentation not uploaded within 5 working days" });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HON-5Workingdays Deviation Date Trigger"), Value = formData.RequestHonorariumList.IsDeviationUpload });

                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formData.RequestHonorariumList.SalesHeadEmail });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formData.RequestHonorariumList.FinanceHead });

                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formData.RequestHonorariumList.InitiatorName });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formData.RequestHonorariumList.InitiatorEmail });


                        var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId7, new Row[] { newRow7 });




                        var j = 1;
                        foreach (var p in formData.RequestHonorariumList.DeviationFiles)
                        {

                            var name = "2WorkingDaysAfterEvent";

                            var filePath = SheetHelper.testingFile(p, eventId, name);


                            var addedRow = addeddeviationrow[0];

                            var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                    parsedSheetId7, addedRow.Id.Value, filePath, "application/msword");
                            var attachmentintoMain = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                    parsedSheetId, addedRows[0].Id.Value, filePath, "application/msword");
                            j++;

                            if (System.IO.File.Exists(filePath))
                            {
                                SheetHelper.DeleteFile(filePath);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        return BadRequest(ex.Message);
                    }
                }










                var targetRow = sheet1.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == eventId));

                if (targetRow != null)
                {
                    long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(sheet1, "Honorarium Submitted?");
                    var cellToUpdateB = new Cell
                    {
                        ColumnId = honorariumSubmittedColumnId,
                        Value = "Yes"
                    };
                    Row updateRow = new Row { Id = targetRow.Id, Cells = new Cell[] { cellToUpdateB } };
                    var cellToUpdate = targetRow.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                    if (cellToUpdate != null)
                    {
                        cellToUpdate.Value = "Yes";
                    }

                    smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId1, new Row[] { updateRow });


                }

                var s = addedRows[0];
                long ColumnId = SheetHelper.GetColumnIdByName(sheet, "Role");
                var UpdateB = new Cell { ColumnId = ColumnId, Value = formData.RequestHonorariumList.Role };
                Row updateRows = new Row { Id = s.Id, Cells = new Cell[] { UpdateB } };
                var cellsToUpdate = s.Cells.FirstOrDefault(c => c.ColumnId == ColumnId);
                if (cellsToUpdate != null) { cellsToUpdate.Value = formData.RequestHonorariumList.Role; }

                smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId, new Row[] { updateRows });



                return Ok(new
                { Message = "Data added successfully." });

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }




        [HttpPost("AddEventSettlementData"), DisableRequestSizeLimit]
        public IActionResult AddEventSettlementData(EventSettlement formData)
        {

            try
            {

                string sheetId = configuration.GetSection("SmartsheetSettings:EventSettlement").Value;
                string sheetId1 = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;
                string sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;
                long.TryParse(sheetId, out long parsedSheetId);
                long.TryParse(sheetId1, out long parsedSheetId1);
                long.TryParse(sheetId7, out long parsedSheetId7);

                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);

                Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);



                StringBuilder addedInviteesData = new StringBuilder();
                StringBuilder addedExpences = new StringBuilder();
                int addedInviteesDataNo = 1;
                int addedExpencesNo = 1;



                foreach (var formdata in formData.expenseSheets)
                {

                    string rowData = $"{addedExpencesNo}. {formdata.Expense} | AmountExcludingTax: {formdata.AmountExcludingTax}| Amount: {formdata.Amount} | {formdata.BtcorBte}";
                    addedExpences.AppendLine(rowData);
                    addedExpencesNo++;
                }
                string Expense = addedExpences.ToString();


                foreach (var formdata in formData.Invitee)
                {

                    string rowData = $"{addedInviteesDataNo}. {formdata.InviteeName} | {formdata.MISCode} | {formdata.LocalConveyance}";
                    addedInviteesData.AppendLine(rowData);
                    addedInviteesDataNo++;
                }
                string Invitee = addedInviteesData.ToString();


                // /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                var newRow = new Row();
                newRow.Cells = new List<Cell>();
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventId/EventRequestId"), Value = formData.EventId });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventTopic"), Value = formData.EventTopic });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventType"), Value = formData.EventType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventDate"), Value = formData.EventDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "StartTime"), Value = formData.StartTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EndTime"), Value = formData.EndTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "VenueName"), Value = formData.VenueName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "City"), Value = formData.City });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "State"), Value = formData.State });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Attended"), Value = formData.Attended });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "InviteesParticipated"), Value = Invitee });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "ExpenseDetails"), Value = Expense });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "TotalExpenseDetails"), Value = formData.TotalExpense });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "AdvanceDetails"), Value = formData.Advance });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "InitiatorName"), Value = formData.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Brands"), Value = formData.Brands });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Invitees"), Value = Invitee });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Panelists"), Value = formData.Panalists });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "SlideKits"), Value = formData.SlideKits });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Expenses"), Value = Expense });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Invitees"), Value = SheetHelper.NumCheck(formData.totalInvitees) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Attendees"), Value = SheetHelper.NumCheck(formData.TotalAttendees) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "IsAdvanceRequired"), Value = formData.IsAdvanceRequired });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "PostEventSubmitted?"), Value = formData.PostEventSubmitted });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Medical Affairs Head"), Value = formData.MedicalAffairsHead });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance Head"), Value = formData.FinanceHead });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance"), Value = formData.FinanceHead });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Initiator Email"), Value = formData.InitiatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "RBM/BM"), Value = formData.RBMorBM });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Sales Head"), Value = formData.SalesHead });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Sales Coordinator"), Value = formData.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Marketing Head"), Value = formData.MarkeringHead });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Compliance"), Value = formData.Compliance });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance Accounts"), Value = formData.FinanceAccounts });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance Treasury"), Value = formData.FinanceTreasury });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Reporting Manager"), Value = formData.ReportingManagerEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "1 Up Manager"), Value = formData.FirstLevelEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Role"), Value = formData.Role });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Travel Amount"), Value = SheetHelper.NumCheck(formData.TotalTravelSpend) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Accommodation Amount"), Value = SheetHelper.NumCheck(formData.TotalAccomodationSpend) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Expense"), Value = SheetHelper.NumCheck(formData.TotalExpenses) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Travel & Accommodation Amount"), Value = SheetHelper.NumCheck(formData.TotalTravelAndAccomodationSpend) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Honorarium Amount"), Value = SheetHelper.NumCheck(formData.TotalHonorariumSpend) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Budget"), Value = SheetHelper.NumCheck(formData.TotalSpend) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Actual"), Value = SheetHelper.NumCheck(formData.TotalActuals) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Advance Utilized For Event"), Value = SheetHelper.NumCheck(formData.AdvanceUtilizedForEvents) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Pay Back Amount To Company"), Value = SheetHelper.NumCheck(formData.PayBackAmountToCompany) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Additional Amount Needed To Pay For Initiator"), Value = SheetHelper.NumCheck(formData.AdditionalAmountNeededToPayForInitiator) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Local Conveyance"), Value = SheetHelper.NumCheck(formData.TotalLocalConveyance) });


                var addedRows = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId, new Row[] { newRow });

                var eventIdColumnId = SheetHelper.GetColumnIdByName(sheet, "EventId/EventRequestId");
                var eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                var val = eventIdCell.DisplayValue;

                var x = 1;
                foreach (var p in formData.Files)
                {

                    string[] words = p.Split(':');
                    var r = words[0];
                    var q = words[1];

                    var name = r.Split(".")[0];

                    var filePath = SheetHelper.testingFile(q, val, name);




                    var addedRow = addedRows[0];

                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                            parsedSheetId, addedRow.Id.Value, filePath, "application/msword");
                    x++;


                    if (System.IO.File.Exists(filePath))
                    {
                        SheetHelper.DeleteFile(filePath);
                    }
                }
                if (formData.EventOpen30Days == "Yes" || formData.EventLessThan5Days == "Yes" || formData.IsDeviationUpload == "Yes")
                {
                    List<string> DeviationNames = new List<string>();
                    foreach (var p in formData.DeviationFiles)
                    {
                        string[] words = p.Split(':');
                        var r = words[0];
                        DeviationNames.Add(r);
                    }
                    foreach (var deviationname in DeviationNames)
                    {
                        var file = deviationname.Split(".")[0];
                        var eventId = val;

                        try
                        {

                            var newRow7 = new Row();
                            newRow7.Cells = new List<Cell>();

                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = val });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formData.EventTopic });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formData.EventType });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formData.EventDate });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "StartTime"), Value = formData.StartTime });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EndTime"), Value = formData.EndTime });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "VenueName"), Value = formData.VenueName });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "City"), Value = formData.City });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "State"), Value = formData.State });
                            if (file == "30DaysDeviationFile")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "POST- Beyond45Days Deviation Date Trigger"), Value = formData.EventOpen30Days });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Outstanding with initiator for more than 45 days" });
                            }
                            else if (file == "Lessthan5InviteesDeviationFile")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "POST-Lessthan5Invitees Deviation Trigger"), Value = formData.EventLessThan5Days });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Less than 5 attendees excluding speaker" });
                            }
                            else if (file == "ExcludingGSTDeviationFile")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "POST-Deviation Excluding GST?"), Value = "Yes" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "POST-Deviation Excluding GST" });
                            }
                            else if (file == "Change in venue")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "POST-Deviation Change in venue trigger"), Value = "Yes" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = file });

                            }
                            else if (file == "Change in topic")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "POST-Deviation Change in topic trigger"), Value = "Yes" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = file });

                            }
                            else if (file == "Change in speaker")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "POST-Deviation Change in speaker trigger"), Value = "Yes" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = file });

                            }
                            else if (file == "Attendees not captured in photographs")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "POST-Deviation Attendees not captured trigger"), Value = "Yes" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = file });

                            }
                            else if (file == "Speaker not captured in photographs")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "POST-Deviation Speaker not captured trigger"), Value = "Yes" });//POST-Deviation Speaker not captured  trigger
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = file });

                            }
                            else
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "POST-Deviation Other Deviation Trigger"), Value = "Yes" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Others" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Other Deviation Type"), Value = file });

                            }



                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formData.SalesHead });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formData.FinanceHead });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formData.InitiatorName });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formData.InitiatorEmail });

                            var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId7, new Row[] { newRow7 });




                            var j = 1;
                            foreach (var p in formData.DeviationFiles)
                            {
                                string[] words = p.Split(':');
                                var r = words[0];
                                var q = words[1];
                                if (deviationname == r)
                                {
                                    var name = r.Split(".")[0];

                                    var filePath = SheetHelper.testingFile(q, val, name);




                                    var addedRow = addeddeviationrow[0];

                                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                            parsedSheetId7, addedRow.Id.Value, filePath, "application/msword");
                                    var attachmentintoMain = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                            parsedSheetId, addedRows[0].Id.Value, filePath, "application/msword");
                                    j++;

                                    if (System.IO.File.Exists(filePath))
                                    {
                                        SheetHelper.DeleteFile(filePath);
                                    }
                                }



                            }
                        }
                        catch (Exception ex)
                        {
                            return BadRequest(ex.Message);
                        }
                    }



                }


                var s = addedRows[0];
                long ColumnId = SheetHelper.GetColumnIdByName(sheet, "Role");
                var UpdateB = new Cell { ColumnId = ColumnId, Value = formData.Role };
                Row updateRows = new Row { Id = s.Id, Cells = new Cell[] { UpdateB } };
                var cellsToUpdate = s.Cells.FirstOrDefault(c => c.ColumnId == ColumnId);
                if (cellsToUpdate != null) { cellsToUpdate.Value = formData.Role; }

                smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId, new Row[] { updateRows });


                return Ok(new
                { Message = "Data added successfully." });













            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }




        [HttpPost("AddEventRequestExpensesData"), DisableRequestSizeLimit]
        public IActionResult AddEventRequestExpensesData(EventRequestExpenseSheet formData)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;


                long.TryParse(sheetId, out long parsedSheetId);

                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);



                var newRow = new Row();
                newRow.Cells = new List<Cell>();
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "HCP Name"), Value = formData.EventId });

                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventId/EventRequestId"), Value = formData.Expense });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventType"), Value = formData.Amount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "HCPRole"), Value = formData.AmountExcludingTax });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "MISCODE"), Value = formData.BtcorBte });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "GO/Non-GO"), Value = formData.BtcAmount });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = SheetHelper.GetColumnIdByName(sheet, "IsItincludingGST?"),
                    Value = formData.BteAmount
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = SheetHelper.GetColumnIdByName(sheet, "AgreementAmount"),
                    Value = formData.BudgetAmount
                });



                var addedRows = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId, new Row[] { newRow });


                return Ok(new
                { Message = "Data added successfully." });

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }


    }
}





