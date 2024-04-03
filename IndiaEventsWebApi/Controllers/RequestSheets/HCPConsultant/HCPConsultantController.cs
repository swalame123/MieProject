using IndiaEventsWebApi.Helper;
using IndiaEventsWebApi.Models;
using IndiaEventsWebApi.Models.EventTypeSheets;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Globalization;
using System.Text;

namespace IndiaEventsWebApi.Controllers.RequestSheets.HCPConsultant
{
    [Route("api/[controller]")]
    [ApiController]
    //[Authorize]
    public class HCPConsultantController : ControllerBase
    {

        private readonly string accessToken;
        private readonly IConfiguration configuration;

        public HCPConsultantController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

        }
        [HttpPost("HCPConsultant"), DisableRequestSizeLimit]
        public IActionResult AllObjModelsData(HCPConsultantPayload formDataList)
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

            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
            Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
            Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
            Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);


            //Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId1, null, null, null, null, null, null, null);
            //Sheet sheet2 = smartsheet.SheetResources.GetSheet(parsedSheetId2, null, null, null, null, null, null, null);
            //Sheet sheet4 = smartsheet.SheetResources.GetSheet(parsedSheetId4, null, null, null, null, null, null, null);
            //Sheet sheet6 = smartsheet.SheetResources.GetSheet(parsedSheetId6, null, null, null, null, null, null, null);
            //Sheet sheet7 = smartsheet.SheetResources.GetSheet(parsedSheetId7, null, null, null, null, null, null, null);

            StringBuilder addedBrandsData = new StringBuilder();
            StringBuilder addedHcpData = new StringBuilder();
            StringBuilder addedExpences = new StringBuilder();

            int addedSlideKitDataNo = 1;
            int addedHcpDataNo = 1;
            int addedInviteesDataNo = 1;
            int addedBrandsDataNo = 1;
            int addedExpencesNo = 1;

            var TotalHonorariumAmount = 0;
            var TotalTravelAmount = 0;
            var TotalAccomodateAmount = 0;
            var TotalHCPLcAmount = 0;
            var TotalInviteesLcAmount = 0;
            var TotalExpenseAmount = 0;
            var TotalRegstAmount = 0;


            CultureInfo hindi = new CultureInfo("hi-IN");


            var EventOpen30Days = !string.IsNullOrEmpty(formDataList.HcpConsultant.EventOpen30days) ? "Yes" : "No";
            var EventWithin7Days = !string.IsNullOrEmpty(formDataList.HcpConsultant.EventWithin7days) ? "Yes" : "No";
            var BrouchereUpload = !string.IsNullOrEmpty(formDataList.HcpConsultant.BrochureFile) ? "Yes" : "No";
            var FCPA = !string.IsNullOrEmpty(formDataList.HcpConsultant.FcpaFile) ? "Yes" : "No";


            foreach (var formdata in formDataList.ExpenseSheet)
            {
                string rowData = $"{addedExpencesNo}. {formdata.Expense} | RegstAmount: {formdata.RegstAmount}| {formdata.BTC_BTE}";
                addedExpences.AppendLine(rowData);
                addedExpencesNo++;

                var amount = SheetHelper.NumCheck(formdata.ExpenseAmount);
                var regst = SheetHelper.NumCheck(formdata.RegstAmount);
                TotalExpenseAmount = TotalExpenseAmount + amount;
                TotalRegstAmount = TotalRegstAmount + regst;
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
                var HM = SheetHelper.NumCheck(formdata.RegistrationAmount);
                var x = string.Format(hindi, "{0:#,#}", HM);
                var t = SheetHelper.NumCheck(formdata.TravelAmount) + SheetHelper.NumCheck(formdata.AccomAmount);
                var y = string.Format(hindi, "{0:#,#}", t);
                string rowData = $"{addedHcpDataNo}. {formdata.MisCode} |{formdata.HcpName} | Regst.Amt: {x} |Trav.&Acc.Amt: {y} ";



                addedHcpData.AppendLine(rowData);
                addedHcpDataNo++;
                TotalHonorariumAmount = TotalHonorariumAmount + SheetHelper.NumCheck(formdata.RegistrationAmount);
                TotalTravelAmount = TotalTravelAmount + SheetHelper.NumCheck(formdata.TravelAmount);
                TotalAccomodateAmount = TotalAccomodateAmount + SheetHelper.NumCheck(formdata.AccomAmount);
                TotalHCPLcAmount = TotalHCPLcAmount + SheetHelper.NumCheck(formdata.LcAmount);
            }
            string HCP = addedHcpData.ToString();
            var c = TotalHCPLcAmount + TotalInviteesLcAmount;
            var total = TotalHonorariumAmount + TotalTravelAmount + TotalAccomodateAmount + TotalHCPLcAmount + TotalInviteesLcAmount;
            //var BTE =SheetHelper.NumCheck(formDataList.HcpConsultant.TotalExpenseBTE);
            //var BTC =SheetHelper.NumCheck(formDataList.HcpConsultant.TotalExpenseBTC);

            //var total = BTC + BTE;

            var s = (TotalTravelAmount + TotalAccomodateAmount);

            try
            {
                var newRow = new Row();
                newRow.Cells = new List<Cell>();
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.HcpConsultant.EventTopic });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventType"), Value = formDataList.HcpConsultant.EventType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventDate"), Value = formDataList.HcpConsultant.EventDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event End Date"), Value = formDataList.HcpConsultant.EventEndDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "VenueName"), Value = formDataList.HcpConsultant.VenueName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "StartTime"), Value = formDataList.HcpConsultant.StartTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EndTime"), Value = formDataList.HcpConsultant.EndTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sponsorship Society Name"), Value = formDataList.HcpConsultant.SponsorshipSocietyName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Venue Country"), Value = formDataList.HcpConsultant.Country });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "IsAdvanceRequired"), Value = formDataList.HcpConsultant.IsAdvanceRequired });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Advance Amount"), Value = SheetHelper.NumCheck(formDataList.HcpConsultant.AdvanceAmount) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Brands"), Value = brand });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Expenses"), Value = Expense });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Panelists"), Value = HCP });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "InitiatorName"), Value = formDataList.HcpConsultant.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total HCP Registration Amount"), Value = TotalHonorariumAmount });
                //newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total HCP Registration Amount"), Value = TotalExpenseAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel Amount"), Value = TotalTravelAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel & Accommodation Amount"), Value = s });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Accomodation Amount"), Value = TotalAccomodateAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Local Conveyance"), Value = c });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense"), Value = TotalExpenseAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Budget"), Value = total });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Email"), Value = formDataList.HcpConsultant.Initiator_Email });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "RBM/BM"), Value = formDataList.HcpConsultant.RBMorBM });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Head"), Value = formDataList.HcpConsultant.Sales_Head });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Coordinator"), Value = formDataList.HcpConsultant.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Marketing Head"), Value = formDataList.HcpConsultant.Marketing_Head });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Compliance"), Value = formDataList.HcpConsultant.ComplianceEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Accounts"), Value = formDataList.HcpConsultant.FinanceAccountsEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Treasury"), Value = formDataList.HcpConsultant.Finance });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Reporting Manager"), Value = formDataList.HcpConsultant.ReportingManagerEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "1 Up Manager"), Value = formDataList.HcpConsultant.FirstLevelEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Affairs Head"), Value = formDataList.HcpConsultant.MedicalAffairsEmail });
                //newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role"), Value = formDataList.HcpConsultant.Role });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, " Total Expense BTC"), Value = SheetHelper.NumCheck(formDataList.HcpConsultant.TotalExpenseBTC) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense BTE"), Value = SheetHelper.NumCheck(formDataList.HcpConsultant.TotalExpenseBTE) });

                var addedRows = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId1, new Row[] { newRow });

                var eventIdColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                var eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                var val = eventIdCell.DisplayValue;

                if (BrouchereUpload == "Yes")
                {


                    var filename = "Brochure";
                    var filePath = SheetHelper.testingFile(formDataList.HcpConsultant.BrochureFile, val, filename);


                    var addedRow = addedRows[0];
                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                            parsedSheetId1, addedRow.Id.Value, filePath, "application/msword");


                    if (System.IO.File.Exists(filePath))
                    {
                        System.IO.File.Delete(filePath);
                    }
                }



                if (EventOpen30Days == "Yes" || EventWithin7Days == "Yes" || formDataList.HcpConsultant.AggregateDeviation == "Yes")
                {
                    List<string> DeviationNames = new List<string>();

                    DeviationNames.Add($"30DaysDeviationFile:{formDataList.HcpConsultant.EventOpen30days}");
                    DeviationNames.Add($"7DaysDeviationFile:{formDataList.HcpConsultant.EventWithin7days}");
                    DeviationNames.Add($"AgregateSpendDeviationFile:{formDataList.HcpConsultant.AggregateDeviationFiles}");
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
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.HcpConsultant.EventTopic });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formDataList.HcpConsultant.EventType });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formDataList.HcpConsultant.EventDate });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "StartTime"), Value = formDataList.HcpConsultant.StartTime });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EndTime"), Value = formDataList.HcpConsultant.EndTime });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "VenueName"), Value = formDataList.HcpConsultant.VenueName });
                                if (fn == "30DaysDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = EventOpen30Days });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Outstanding with initiator for more than 45 days" });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = SheetHelper.NumCheck(formDataList.HcpConsultant.EventOpen30dayscount) });

                                }
                                else if (fn == "7DaysDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = EventWithin7Days });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "5 days from the Event Date" });
                                }
                                else if (fn == "AgregateSpendDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HCP exceeds 5,00,000 Trigger"), Value = formDataList.HcpConsultant.AggregateDeviation });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Aggregate Limit of 5,00,000 is Exceeded" });
                                }
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.HcpConsultant.Sales_Head });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formDataList.HcpConsultant.FinanceHead });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formDataList.HcpConsultant.InitiatorName });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.HcpConsultant.Initiator_Email });

                                var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId7, new Row[] { newRow7 });

                                if (fn == "AgregateSpendDeviationFile")
                                {
                                    foreach (var file in formDataList.HcpConsultant.AggregateDeviationFiles)
                                    {
                                        var filename = fn;
                                        var filePath = SheetHelper.testingFile(file, val, filename);
                                        var addedRow = addeddeviationrow[0];
                                        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                                parsedSheetId7, addedRow.Id.Value, filePath, "application/msword");
                                        var attachmentintoMain = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                                parsedSheetId1, addedRows[0].Id.Value, filePath, "application/msword");
                                        if (System.IO.File.Exists(filePath))
                                        {
                                            SheetHelper.DeleteFile(filePath);
                                        }
                                    }
                                }
                                else
                                {
                                    var filename = fn;
                                    var filePath = SheetHelper.testingFile(bs, val, filename);

                                    var addedRow = addeddeviationrow[0];

                                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                            parsedSheetId7, addedRow.Id.Value, filePath, "application/msword");
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
                    }


                }







                foreach (var formData in formDataList.HcpList)
                {
                    var newRow1 = new Row();
                    newRow1.Cells = new List<Cell>();
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCPName"), Value = formData.HcpName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "MISCode"), Value = formData.MisCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCP Type"), Value = formData.HcpType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Travel"), Value = SheetHelper.NumCheck(formData.TravelAmount) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Topic"), Value = formDataList.HcpConsultant.EventTopic });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Type"), Value = formDataList.HcpConsultant.EventType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Date Start"), Value = formDataList.HcpConsultant.EventDate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event End Date"), Value = formDataList.HcpConsultant.EventEndDate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Venue name"), Value = formDataList.HcpConsultant.VenueName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Accomodation"), Value = SheetHelper.NumCheck(formData.AccomAmount) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "LocalConveyance"), Value = SheetHelper.NumCheck(formData.LcAmount) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Registration Amount"), Value = SheetHelper.NumCheck(formData.RegistrationAmount) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TotalSpend"), Value = SheetHelper.NumCheck(formData.BudgetAmount) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Legitimate Need"), Value = formData.Legitimate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Objective Criteria"), Value = formData.Objective });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Rationale"), Value = formData.Rationale });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "FCPA Date"), Value = formData.Fcpadate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "ExpenseType"), Value = formData.ExpenseType });



                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "EventId/EventRequestId"), Value = val });







                    var addeddatarow = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId4, new Row[] { newRow1 });

                    var columnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                    var Cell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == columnId);
                    var value = Cell.DisplayValue;
                    if (FCPA == "Yes")
                    {
                        var filename = " FCPA";
                        var filePath = SheetHelper.testingFile(formDataList.HcpConsultant.FcpaFile, value, filename);
                        var addedRow = addedRows[0];
                        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(parsedSheetId1, addedRow.Id.Value, filePath, "application/msword");
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

                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Expense"), Value = formdata.Expense });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "EventId/EventRequestID"), Value = val });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formdata.BTC_BTE });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTCAmount"), Value = SheetHelper.NumCheck(formdata.BtcAmount) });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTEAmount"), Value = SheetHelper.NumCheck(formdata.BteAmount) });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Registration Amount"), Value = SheetHelper.NumCheck(formdata.RegstAmount) });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Registration Amount Excluding Tax"), Value = formdata.RegstAmountExcludingTax });

                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount"), Value = SheetHelper.NumCheck(formdata.ExpenseAmount )});
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount Excluding Tax"), Value = formdata.ExpenseAmountExcludingTax });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Topic"), Value = formDataList.HcpConsultant.EventTopic });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Type"), Value = formDataList.HcpConsultant.EventType });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Date Start"), Value = formDataList.HcpConsultant.EventDate });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event End Date"), Value = formDataList.HcpConsultant.EventEndDate });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Venue name"), Value = formDataList.HcpConsultant.VenueName });

                    smartsheet.SheetResources.RowResources.AddRows(parsedSheetId6, new Row[] { newRow6 });
                }



                var addedrow = addedRows[0];
                long ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role");
                var UpdateB = new Cell { ColumnId = ColumnId, Value = formDataList.HcpConsultant.Role };
                Row updateRows = new Row { Id = addedrow.Id, Cells = new Cell[] { UpdateB } };
                var cellsToUpdate = addedrow.Cells.FirstOrDefault(c => c.ColumnId == ColumnId);
                if (cellsToUpdate != null) { cellsToUpdate.Value = formDataList.HcpConsultant.Role; }

                smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId1, new Row[] { updateRows });

                return Ok(new
                { Message = " Success!" });
            }
            catch (Exception ex)
            {
                return BadRequest($"Could not find {ex.Message}");
            }
        }


        [HttpGet("HcpFollowUpData")]
        public IActionResult GetEventRequestWebData()
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:HcpFollowup").Value;
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


        [HttpPost("HcpFollowup"), DisableRequestSizeLimit]
        public IActionResult HcpFollowup(List<HCPfollow_upsheet> formDataList)
        {
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            string sheetId1 = configuration.GetSection("SmartsheetSettings:HcpFollowup").Value;
            long.TryParse(sheetId1, out long parsedSheetId1);
            Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId1, null, null, null, null, null, null, null);

            foreach (var formdata in formDataList)
            {
                try
                {

                    var newRow = new Row();
                    newRow.Cells = new List<Cell>();
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "HCP Name"), Value = formdata.HCPName });

                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "MIS Code"), Value = formdata.MisCode });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "GO/N-GO"), Value = formdata.GO_NGO });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Country"), Value = formdata.Country });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "How many days since the parent event Completes"), Value = formdata.How_many_days_since_the_parent_event_completes });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Follow-up Event Date"), Value = formdata.Follow_up_Event_Date });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Date"), Value = formdata.Event_Date });

                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Follow-up Event Code"), Value = formdata.Follow_up_Event });

                    var addedRows = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId1, new Row[] { newRow });
                    var columnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                    var Cell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == columnId);
                    var value = Cell.DisplayValue;
                    if (formdata.AgreementFile != "")
                    {

                        var filename = "AgreementFile";
                        var filePath = SheetHelper.testingFile(formdata.AgreementFile, value, filename);




                        var addedRow = addedRows[0];


                        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                parsedSheetId1, addedRow.Id.Value, filePath, "application/msword");

                    }



                }
                catch (Exception ex)
                {
                    return BadRequest(ex.Message);
                }
            }
            return Ok(new { Message = " success!" });
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
        //private Row GetRowById(SmartsheetClient smartsheet, long sheetId, string val)
        //{
        //    Sheet sheet = smartsheet.SheetResources.GetSheet(sheetId, null, null, null, null, null, null, null);

        //    // Assuming you have a column named "Id"

        //    Column idColumn = sheet.Columns.FirstOrDefault(col => col.Title == "Honorarium Submitted?");

        //    if (idColumn != null)
        //    {
        //        foreach (var row in sheet.Rows)
        //        {
        //            var cell = row.Cells.FirstOrDefault(c => c.ColumnId == idColumn.Id && c.Value.ToString() == val);

        //            if (cell != null)
        //            {
        //                return row;
        //            }
        //        }
        //    }

        //    return null;
        //}


    }
}
