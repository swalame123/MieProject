using IndiaEvents.Models.Models.EventTypeSheets;
using IndiaEventsWebApi.Helper;
using IndiaEventsWebApi.Models.EventTypeSheets;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Text;
using static Org.BouncyCastle.Bcpg.Attr.ImageAttrib;

namespace IndiaEventsWebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class EventSettlementController : ControllerBase
    {


        private readonly string accessToken;
        private readonly IConfiguration configuration;
        private readonly SmartsheetClient smartsheet;

        private readonly Sheet sheet1;
        private readonly Sheet sheet2;
        private readonly Sheet sheet3;
        private readonly Sheet sheet4;


        public EventSettlementController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

            string panelSheet = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
            string InviteeSheet = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
            string ExpenseSheet = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
            string SlideKitSheet = configuration.GetSection("SmartsheetSettings:EventRequestsHcpSlideKit").Value;

            smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

            sheet1 = SheetHelper.GetSheetById(smartsheet, panelSheet);
            sheet2 = SheetHelper.GetSheetById(smartsheet, InviteeSheet);
            sheet3 = SheetHelper.GetSheetById(smartsheet, ExpenseSheet);
            sheet4 = SheetHelper.GetSheetById(smartsheet, SlideKitSheet);
        }

        //[HttpPost("AddEventSettlementData")]
        //public IActionResult AddEventSettlementData(EventSettlement formData)
        //{
        //    try
        //    {
        //        SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

        //        string sheetId = configuration.GetSection("SmartsheetSettings:EventSettlement").Value;
        //        string sheetId1 = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;
        //        string sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;
        //        Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
        //        Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
        //        Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);

        //        StringBuilder addedInviteesData = new StringBuilder();
        //        StringBuilder addedExpences = new StringBuilder();
        //        int addedInviteesDataNo = 1;
        //        int addedExpencesNo = 1;
        //        foreach (var formdata in formData.expenseSheets)
        //        {
        //            string rowData = $"{addedExpencesNo}. {formdata.Expense} | AmountExcludingTax: {formdata.AmountExcludingTax}| Amount: {formdata.Amount} | {formdata.BtcorBte}";
        //            addedExpences.AppendLine(rowData);
        //            addedExpencesNo++;
        //        }
        //        string Expense = addedExpences.ToString();
        //        foreach (var formdata in formData.Invitee)
        //        {
        //            string rowData = $"{addedInviteesDataNo}. {formdata.InviteeName} | {formdata.MISCode} | {formdata.LocalConveyance}";
        //            addedInviteesData.AppendLine(rowData);
        //            addedInviteesDataNo++;
        //        }
        //        string Invitee = addedInviteesData.ToString();

        //        var newRow = new Row();
        //        newRow.Cells = new List<Cell>();
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventId/EventRequestId"), Value = formData.EventId });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventTopic"), Value = formData.EventTopic });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventType"), Value = formData.EventType });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventDate"), Value = formData.EventDate });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "StartTime"), Value = formData.StartTime });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EndTime"), Value = formData.EndTime });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "VenueName"), Value = formData.VenueName });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "City"), Value = formData.City });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "State"), Value = formData.State });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Attended"), Value = formData.Attended });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "InviteesParticipated"), Value = Invitee });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "ExpenseDetails"), Value = Expense });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "TotalExpenseDetails"), Value = int.Parse(formData.TotalExpense) });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "AdvanceDetails"), Value = formData.Advance });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "InitiatorName"), Value = formData.InitiatorName });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Brands"), Value = formData.Brands });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Invitees"), Value = Invitee });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Panelists"), Value = formData.Panalists });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "SlideKits"), Value = formData.SlideKits });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Expenses"), Value = Expense });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Invitees"), Value = int.Parse(formData.totalInvitees )});
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Attendees"), Value = int.Parse(formData.TotalAttendees )});
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "IsAdvanceRequired"), Value = formData.IsAdvanceRequired });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "PostEventSubmitted?"), Value = formData.PostEventSubmitted });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Medical Affairs Head"), Value = formData.MedicalAffairsHead });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance Head"), Value = formData.FinanceHead });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance"), Value = formData.FinanceHead });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Initiator Email"), Value = formData.InitiatorEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "RBM/BM"), Value = formData.RBMorBM });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Sales Head"), Value = formData.SalesHead });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Sales Coordinator"), Value = formData.SalesCoordinatorEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Marketing Head"), Value = formData.MarkeringHead });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Compliance"), Value = formData.Compliance });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance Accounts"), Value = formData.FinanceAccounts });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance Treasury"), Value = formData.FinanceTreasury });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Reporting Manager"), Value = formData.ReportingManagerEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "1 Up Manager"), Value = formData.FirstLevelEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Role"), Value = formData.Role });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Travel Amount"), Value = int.Parse(formData.TotalTravelSpend) });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Accommodation Amount"), Value = int.Parse(formData.TotalAccomodationSpend) });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Expense"), Value = int.Parse(formData.TotalExpenses )});
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Travel & Accommodation Amount"), Value = int.Parse(formData.TotalTravelAndAccomodationSpend )});
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Honorarium Amount"), Value = int.Parse(formData.TotalHonorariumSpend) });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Budget"), Value = int.Parse(formData.TotalSpend )});
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Actual"), Value = int.Parse(formData.TotalActuals )});
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Advance Utilized For Event"), Value = int.Parse(formData.AdvanceUtilizedForEvents) });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Pay Back Amount To Company"), Value = int.Parse(formData.PayBackAmountToCompany) });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Additional Amount Needed To Pay For Initiator"), Value = int.Parse(formData.AdditionalAmountNeededToPayForInitiator )});
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Local Conveyance"), Value = int.Parse(formData.TotalLocalConveyance) });


        //        var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheet.Id.Value, new Row[] { newRow });

        //        var eventIdColumnId = SheetHelper.GetColumnIdByName(sheet, "EventId/EventRequestId");
        //        var eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
        //        var val = eventIdCell.DisplayValue;

        //        var x = 1;
        //        foreach (var p in formData.Files)
        //        {

        //            string[] words = p.Split(':');
        //            var r = words[0];
        //            var q = words[1];

        //            var name = r.Split(".")[0];

        //            var filePath = SheetHelper.testingFile(q, val, name);




        //            var addedRow = addedRows[0];

        //            var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
        //                    sheet.Id.Value, addedRow.Id.Value, filePath, "application/msword");
        //            x++;


        //            if (System.IO.File.Exists(filePath))
        //            {
        //                SheetHelper.DeleteFile(filePath);
        //            }
        //        }
        //        if (formData.EventOpen30Days == "Yes" || formData.EventLessThan5Days == "Yes" || formData.IsDeviationUpload == "Yes")
        //        {
        //            List<string> DeviationNames = new List<string>();
        //            foreach (var p in formData.DeviationFiles)
        //            {
        //                string[] words = p.Split(':');
        //                var r = words[0];
        //                DeviationNames.Add(r);
        //            }
        //            foreach (var deviationname in DeviationNames)
        //            {
        //                var file = deviationname.Split(".")[0];
        //                var eventId = val;

        //                try
        //                {

        //                    var newRow7 = new Row();
        //                    newRow7.Cells = new List<Cell>();

        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = val });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formData.EventTopic });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formData.EventType });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formData.EventDate });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "StartTime"), Value = formData.StartTime });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EndTime"), Value = formData.EndTime });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "VenueName"), Value = formData.VenueName });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "City"), Value = formData.City });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "State"), Value = formData.State });
        //                    if (file == "30DaysDeviationFile")
        //                    {
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "POST- Beyond45Days Deviation Date Trigger"), Value = formData.EventOpen30Days });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Outstanding with initiator for more than 45 days" });
        //                    }
        //                    else if (file == "Lessthan5InviteesDeviationFile")
        //                    {
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "POST-Lessthan5Invitees Deviation Trigger"), Value = formData.EventLessThan5Days });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Less than 5 attendees excluding speaker" });
        //                    }
        //                    else if (file == "ExcludingGSTDeviationFile")
        //                    {
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "POST-Deviation Excluding GST?"), Value = "Yes" });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "POST-Deviation Excluding GST" });
        //                    }



        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formData.Compliance });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formData.FinanceHead });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formData.InitiatorName });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formData.InitiatorEmail });

        //                    var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });




        //                    var j = 1;
        //                    foreach (var p in formData.DeviationFiles)
        //                    {
        //                        string[] words = p.Split(':');
        //                        var r = words[0];
        //                        var q = words[1];
        //                        if (deviationname == r)
        //                        {
        //                            var name = r.Split(".")[0];

        //                            var filePath = SheetHelper.testingFile(q, val, name);




        //                            var addedRow = addeddeviationrow[0];
        //                            var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");
        //                            j++;
        //                            if (System.IO.File.Exists(filePath))
        //                            {
        //                                SheetHelper.DeleteFile(filePath);
        //                            }
        //                        }
        //                    }
        //                }
        //                catch (Exception ex)
        //                {
        //                    return BadRequest(ex.Message);
        //                }
        //            }
        //        }
        //        return Ok(new { Message = "Data added successfully." });
        //    }
        //    catch (Exception ex)
        //    {
        //        return BadRequest(ex.Message);
        //    }
        //}



        [HttpPost("AddEventSettlementDataForFirst5"), DisableRequestSizeLimit]
        public IActionResult AddEventSettlementDataForFirst5(EventSettlement formData)
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







        [HttpPut("EventSettlementUpdate")]
        public IActionResult EventSettlementUpdate(UpdateAttendees formData)
        {
            try
            {
                if (formData.InviteesData.Count > 0)
                {
                    foreach (var formdata in formData.InviteesData)
                    {
                        var targetRow = sheet2.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == formdata.InviteeId));
                        if (targetRow != null)
                        {
                            Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };

                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Attended?"), Value = formdata.IsAttended });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Actual Local Conveyance Amount"), Value = formdata.ActualAmount });
                            IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet2.Id.Value, new Row[] { updateRow });
                            if (formdata.IsUploadDocument == "Yes")
                            {
                                foreach (var p in formdata.UploadDocument)
                                {
                                    string[] words = p.Split(':');
                                    var r = words[0];
                                    var q = words[1];
                                    var name = r.Split(".")[0];
                                    var filePath = SheetHelper.testingFile(q, formData.EventId, name);
                                    var addedRow = updatedRow[0];
                                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                            sheet2.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                                    if (System.IO.File.Exists(filePath))
                                    {
                                        SheetHelper.DeleteFile(filePath);
                                    }
                                }

                            }
                        }
                    }
                }
                if (formData.PanelData.Count > 0)
                {
                    foreach (var formdata in formData.PanelData)
                    {
                        var targetRow = sheet1.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == formdata.PanelId));
                        if (targetRow != null)
                        {

                            Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Attended?"), Value = formdata.IsAttended });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Actual Accomodation Amount"), Value = formdata.ActualAccomodationAmount });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Actual Travel Amount"), Value = formdata.ActualTravelAmount });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Actual Local Conveyance Amount"), Value = formdata.ActualLCAmount });

                            IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRow });

                        }
                    }
                }
                if (formData.ExpenseData.Count > 0)
                {
                    foreach (var formdata in formData.ExpenseData)
                    {
                        var targetRow = sheet3.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == formdata.ExpenseId));
                        if (targetRow != null)
                        {
                            Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Actual Amount"), Value = formdata.ActualAmount });

                            IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet3.Id.Value, new Row[] { updateRow });
                            if (formdata.IsUploadDocument == "Yes")
                            {
                                foreach (var p in formdata.UploadDocument)
                                {

                                    string[] words = p.Split(':');
                                    var r = words[0];
                                    var q = words[1];
                                    var name = r.Split(".")[0];
                                    var filePath = SheetHelper.testingFile(q, formData.EventId, name);
                                    var addedRow = updatedRow[0];
                                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                            sheet3.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                                    if (System.IO.File.Exists(filePath))
                                    {
                                        SheetHelper.DeleteFile(filePath);
                                    }
                                }

                            }
                        }
                    }
                }
                if (formData.SlideKitData.Count > 0)
                {
                    foreach (var formdata in formData.SlideKitData)
                    {
                        var targetRow = sheet4.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == formdata.SlideKitId));
                        if (targetRow != null)
                        {
                            Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Actual Amount"), Value = formdata.ProductName });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Actual Amount"), Value = formdata.IndicationsDone });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Actual Amount"), Value = formdata.BatchNumber });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Actual Amount"), Value = formdata.SubjectNameandSurName });

                            IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet4.Id.Value, new Row[] { updateRow });
                            if (formdata.IsUploadDocument == "Yes")
                            {
                                foreach (var p in formdata.UploadDocument)
                                {

                                    string[] words = p.Split(':');
                                    var r = words[0];
                                    var q = words[1];
                                    var name = r.Split(".")[0];
                                    var filePath = SheetHelper.testingFile(q, formData.EventId, name);
                                    var addedRow = updatedRow[0];
                                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                            sheet4.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                                    if (System.IO.File.Exists(filePath))
                                    {
                                        SheetHelper.DeleteFile(filePath);
                                    }
                                }

                            }
                        }
                    }
                }

                return Ok(new { Message = "Attendees Updated Successfully" });
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
            return Ok(new { Message = "" });
        }
    }
}
