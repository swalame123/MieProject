using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using IndiaEventsWebApi.Models;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using Microsoft.AspNetCore.Authorization;
using IndiaEventsWebApi.Helper;
using IndiaEventsWebApi.Models.RequestSheets;
using IndiaEvents.Models.Models.GetData;
using Serilog;
using Aspose.Pdf.Operators;
using Aspose.Pdf.Plugins;

namespace IndiaEventsWebApi.Controllers.RequestSheets
{

    [Route("api/[controller]")]
    [ApiController]
    //[Authorize]
    public class GetRequestSheetsController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;
        private readonly SmartsheetClient smartsheet;

        public GetRequestSheetsController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
            smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
        }

        [HttpGet("GetEventRequestWebData")]
        public IActionResult GetEventRequestWebData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:Class1").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("GetEventRequestProcessData")]
        public IActionResult GetEventRequestProcessData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("GetHonorariumPaymentData")]
        public IActionResult GetHonorariumPaymentData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:HonorariumPayment").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("GetEventSettlementData")]
        public IActionResult GetEventSettlementData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:EventSettlement").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("GetHCPSlideKitDetailsData")]
        public IActionResult GetHCPSlideKitDetailsData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsHcpSlideKit").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("GetExpenseData")]
        public IActionResult GetExpenseData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("GetHCPRoleDetailsData")]
        public IActionResult GetHCPRoleDetailsData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("GetInviteesData")]
        public IActionResult GetInviteesData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("GetEventRequestBrandsListData")]
        public IActionResult GetEventRequestBrandsListData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestBrandsList").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        [HttpPost("GetEventRequestsHcpRoleByIds")]
        public IActionResult GetEventRequestsHcpRoleByIds([FromBody] getIds eventIdorEventRequestIds)
        {

            string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId);
            List<Dictionary<string, object>> hcpRoleData = new List<Dictionary<string, object>>();
            List<string> columnNames = new List<string>();
            foreach (Column column in sheet1.Columns)
            {
                columnNames.Add(column.Title);
            }
            foreach (var val in eventIdorEventRequestIds.EventIds)
            {
                foreach (Row row in sheet1.Rows)
                {
                    var eventIdorEventRequestIdCell = row.Cells.FirstOrDefault(cell => cell.ColumnId == SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId"));
                    var x = eventIdorEventRequestIdCell.Value.ToString();
                    if (eventIdorEventRequestIdCell != null && x == val)
                    {
                        Dictionary<string, object> hcpRoleRowData = new Dictionary<string, object>();

                        for (int i = 0; i < row.Cells.Count && i < columnNames.Count; i++)
                        {
                            hcpRoleRowData[columnNames[i]] = row.Cells[i].Value;
                        }

                        hcpRoleData.Add(hcpRoleRowData);
                    }

                }
            }

            return Ok(hcpRoleData);

        }
        [HttpGet("GetEventRequestsHcpDetailsTotalSpendValue")]
        public IActionResult GetcgtColumnValue(string EventID)
        {
            string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;

            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);

            Column SpecialityColumn = sheet.Columns.FirstOrDefault(column => string.Equals(column.Title, "EventId/EventRequestId", StringComparison.OrdinalIgnoreCase));
            Column targetColumn = sheet.Columns.FirstOrDefault(column => string.Equals(column.Title, "TotalSpend", StringComparison.OrdinalIgnoreCase));

            if (targetColumn != null && SpecialityColumn != null)
            {
                List<Row> targetRows = sheet.Rows
                    .Where(row => row.Cells.Any(cell => cell.ColumnId == SpecialityColumn.Id && cell.Value.ToString() == EventID))
                    .ToList();

                if (targetRows.Any())
                {
                    decimal totalValue = targetRows.Sum(row => Convert.ToDecimal(row.Cells.FirstOrDefault(cell => cell.ColumnId == targetColumn.Id)?.Value ?? 0));
                    return Ok(totalValue);
                }
                else
                {
                    return NotFound("NotFound");
                }
            }
            else
            {
                return NotFound("NotFound");
            }
        }
        [HttpGet("GetEventRequestsInviteesLcAmountValue")]
        public IActionResult GetEventRequestsInviteesLcAmountValue(string EventID)
        {
            string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;

            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);

            Column SpecialityColumn = sheet.Columns.FirstOrDefault(column => string.Equals(column.Title, "EventId/EventRequestId", StringComparison.OrdinalIgnoreCase));
            Column targetColumn = sheet.Columns.FirstOrDefault(column => string.Equals(column.Title, "LcAmount", StringComparison.OrdinalIgnoreCase));

            if (targetColumn != null && SpecialityColumn != null)
            {

                List<Row> targetRows = sheet.Rows
                    .Where(row => row.Cells.Any(cell => cell.ColumnId == SpecialityColumn.Id && cell.Value.ToString() == EventID))
                    .ToList();

                if (targetRows.Any())
                {

                    decimal totalValue = targetRows.Sum(row => Convert.ToDecimal(row.Cells.FirstOrDefault(cell => cell.ColumnId == targetColumn.Id)?.Value ?? 0));

                    return Ok(totalValue);
                }
                else
                {
                    return NotFound("NotFound");
                }
            }
            else
            {
                return NotFound("NotFound");
            }
        }
        [HttpGet("GetEventRequestExpenseSheetAmountValue")]
        public IActionResult GetEventRequestExpenseSheetAmountValue(string EventID)
        {
            string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;

            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);

            Column SpecialityColumn = sheet.Columns.FirstOrDefault(column => string.Equals(column.Title, "EventId/EventRequestID", StringComparison.OrdinalIgnoreCase));
            Column targetColumn = sheet.Columns.FirstOrDefault(column => string.Equals(column.Title, "Amount", StringComparison.OrdinalIgnoreCase));

            if (targetColumn != null && SpecialityColumn != null)
            {
                List<Row> targetRows = sheet.Rows.Where(row => row.Cells?.Any(cell => cell.ColumnId == SpecialityColumn.Id && cell.Value?.ToString() == EventID) == true).ToList();

                if (targetRows.Any() == true)
                {

                    decimal totalValue = targetRows.Sum(row => Convert.ToDecimal(row.Cells?.FirstOrDefault(cell => cell.ColumnId == targetColumn.Id)?.Value ?? 0));

                    return Ok(totalValue);
                }
                else
                {
                    return NotFound("NotFound");
                }
            }
            else
            {
                return NotFound("NotFound");
            }
        }
        [HttpGet("GetProcessSheetData")]
        public IActionResult GetProcessSheetData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                List<ProcessSheetPayload> eventRequestBrandsList = sheetData
                    .Select(row => new ProcessSheetPayload
                    {

                        EventTopic = row.TryGetValue("Event Topic", out var eventTopic) ? eventTopic?.ToString() : null,
                        EventIdEventRequestId = row.TryGetValue("EventId/EventRequestId", out var eventId) ? eventId?.ToString() : null,
                        EventType = row.TryGetValue("EventType", out var eventType) ? eventType?.ToString() : null,
                        EventDate = row.TryGetValue("EventDate", out var eventDate) ? eventDate?.ToString() : null,
                        StartTime = row.TryGetValue("StartTime", out var startTime) ? startTime?.ToString() : null,
                        EndTime = row.TryGetValue("EndTime", out var endTime) ? endTime?.ToString() : null,
                        VenueName = row.TryGetValue("VenueName", out var venueName) ? venueName?.ToString() : null,
                        EventEndDate = row.TryGetValue("Event End Date", out var eventEndDate) ? eventEndDate?.ToString() : null,
                        State = row.TryGetValue("State", out var state) ? state?.ToString() : null,
                        City = row.TryGetValue("City", out var city) ? city?.ToString() : null,
                        InitiatorName = row.TryGetValue("InitiatorName", out var initiatorName) ? initiatorName?.ToString() : null,
                        InitiatorEmail = row.TryGetValue("Initiator Email", out var initiatorEmail) ? initiatorEmail?.ToString() : null,
                        CreatedDateHelper = row.TryGetValue("Created Date - Helper", out var createdDateHelper) ? createdDateHelper?.ToString() : null,
                        IsAdvanceRequired = row.TryGetValue("IsAdvanceRequired", out var isAdvanceRequired) ? isAdvanceRequired?.ToString() : null,
                        PRERBMBMApproval = row.TryGetValue("PRE-RBM/BM Approval", out var prerBmBmApproval) ? prerBmBmApproval?.ToString() : null,
                        PRERBMBMApprovalDate = row.TryGetValue("PRE-RBM/BM Approval Date", out var prerBmBmApprovalDate) ? prerBmBmApprovalDate?.ToString() : null,
                        PRESalesHeadApproval = row.TryGetValue("PRE-Sales Head Approval", out var presalesHeadApproval) ? presalesHeadApproval?.ToString() : null,
                        PRESalesHeadApprovalDate = row.TryGetValue("PRE-Sales Head Approval Date", out var presalesHeadApprovalDate) ? presalesHeadApprovalDate?.ToString() : null,
                        PREMarketingHeadApproval = row.TryGetValue("PRE-Marketing Head Approval", out var premarketingHeadApproval) ? premarketingHeadApproval?.ToString() : null,
                        PREMarketingHeadApprovalDate = row.TryGetValue("PRE-Marketing Head Approval Date", out var premarketingHeadApprovalDate) ? premarketingHeadApprovalDate?.ToString() : null,
                        agreement = row.TryGetValue("Agreement", out var agreement) ? agreement?.ToString() : null,
                        PREFinanceTreasuryApprovalDate = row.TryGetValue("PRE-Finance Treasury Approval Date", out var prefFinanceTreasuryApprovalDate) ? prefFinanceTreasuryApprovalDate?.ToString() : null,
                        PREMedicalAffairsHeadApproval = row.TryGetValue("PRE-Medical Affairs Head Approval", out var premedicalAffairsHeadApproval) ? premedicalAffairsHeadApproval?.ToString() : null,
                        PREMedicalAffairsHeadApprovalDate = row.TryGetValue("PRE-Medical Affairs Head Approval Date", out var premedicalAffairsHeadApprovalDate) ? premedicalAffairsHeadApprovalDate?.ToString() : null,
                        PRESalesCoordinatorApproval = row.TryGetValue("PRE-Sales Coordinator Approval", out var presalesCoordinatorApproval) ? presalesCoordinatorApproval?.ToString() : null,
                        PRESalesCoordinatorApprovalDate = row.TryGetValue("PRE-Sales Coordinator Approval Date", out var presalesCoordinatorApprovalDate) ? presalesCoordinatorApprovalDate?.ToString() : null,
                        PREComplianceApproval = row.TryGetValue("PRE-Compliance Approval", out var precomplianceApproval) ? precomplianceApproval?.ToString() : null,
                        PREComplianceApprovalDate = row.TryGetValue("PRE-Compliance Approval Date", out var precomplianceApprovalDate) ? precomplianceApprovalDate?.ToString() : null,
                        EventRequestStatus = row.TryGetValue("Event Request Status", out var eventRequestStatus) ? eventRequestStatus?.ToString() : null,
                        EventApprovedDate = row.TryGetValue("Event Approved Date", out var eventApprovedDate) ? eventApprovedDate?.ToString() : null,
                        HonorariumRequestStatus = row.TryGetValue("Honorarium Request Status", out var honorariumRequestStatus) ? honorariumRequestStatus?.ToString() : null,
                        HonorariumApprovedDate = row.TryGetValue("Honorarium Approved Date", out var honorariumApprovedDate) ? honorariumApprovedDate?.ToString() : null,
                        PostEventRequeststatus = row.TryGetValue("Post Event Request status", out var postEventRequeststatus) ? postEventRequeststatus?.ToString() : null,
                        HelperFinancetreasurytriggerBTE = row.TryGetValue("Helper Finance treasury trigger(BTE)", out var helperFinancetreasurytriggerBTE) ? helperFinancetreasurytriggerBTE?.ToString() : null,
                        PostEventApprovedDate = row.TryGetValue("Post Event Approved Date", out var postEventApprovedDate) ? postEventApprovedDate?.ToString() : null,
                        EventOpenSalesHeadApproval = row.TryGetValue("EventOpenSalesHeadApproval", out var eventOpenSalesHeadApproval) ? eventOpenSalesHeadApproval?.ToString() : null,
                        EventOpenSalesHeadApprovalDate = row.TryGetValue("EventOpenSalesHeadApproval Date", out var eventOpenSalesHeadApprovalDate) ? eventOpenSalesHeadApprovalDate?.ToString() : null,
                        _7daysSalesHeadApproval = row.TryGetValue("7daysSalesHeadApproval", out var _7daysSalesHeadApproval) ? _7daysSalesHeadApproval?.ToString() : null,
                        _7daysSalesHeadApprovaldate = row.TryGetValue("7daysSalesHeadApproval date", out var _7daysSalesHeadApprovaldate) ? _7daysSalesHeadApprovaldate?.ToString() : null,
                        PREFBExpenseExcludingTaxApproval = row.TryGetValue("PRE-F&B Expense Excluding Tax Approval", out var prefBExpenseExcludingTaxApproval) ? prefBExpenseExcludingTaxApproval?.ToString() : null,
                        HCPexceeds100000FHApproval = row.TryGetValue("HCP exceeds 1,00,000 FH Approval", out var hcpExceeds100000FHApproval) ? hcpExceeds100000FHApproval?.ToString() : null,
                        HCPexceeds500000TriggerFHApproval = row.TryGetValue("HCP exceeds 5,00,000 Trigger FH Approval", out var hcpExceeds500000TriggerFHApproval) ? hcpExceeds500000TriggerFHApproval?.ToString() : null,
                        HCPHonorarium600000ExceededApproval = row.TryGetValue("HCP Honorarium 6,00,000 Exceeded Approval", out var hcpHonorarium600000ExceededApproval) ? hcpHonorarium600000ExceededApproval?.ToString() : null,
                        TrainerHonorarium1200000ExceededApproval = row.TryGetValue("Trainer Honorarium 12,00,000 Exceeded Approval", out var trainerHonorarium120000) ? trainerHonorarium120000?.ToString() : null,
                        TravelAccomodation300000ExceededApproval = row.TryGetValue("Travel/Accomodation 3,00,000 Exceeded Approval", out var travelAccomodation300000ExceededApproval) ? travelAccomodation300000ExceededApproval?.ToString() : null,
                        HonorariumSubmitted = row.TryGetValue("Honorarium Submitted?", out var honorariumSubmitted) ? honorariumSubmitted?.ToString() : null,
                        AgreementDownload = row.TryGetValue("Agreement Download", out var agreementDownload) ? agreementDownload?.ToString() : null,
                        EventOpen30days = row.TryGetValue("EventOpen30days", out var eventOpen30days) ? eventOpen30days?.ToString() : null,
                        EventWithin7days = row.TryGetValue("EventWithin7days", out var eventWithin7days) ? eventWithin7days?.ToString() : null,
                        ViewEventRequest = row.TryGetValue("View Event Request", out var viewEventRequest) ? viewEventRequest?.ToString() : null,
                        ViewHonorariumRequest = row.TryGetValue("View Honorarium Request", out var viewHonorariumRequest) ? viewHonorariumRequest?.ToString() : null,
                        ReportingManager = row.TryGetValue("Reporting Manager", out var reportingManager) ? reportingManager?.ToString() : null,
                        _1UpManager = row.TryGetValue("1 Up Manager", out var _1UpManager) ? _1UpManager?.ToString() : null,
                        Brands = row.TryGetValue("Brands", out var brands) ? brands?.ToString() : null,
                        Panelists = row.TryGetValue("Panelists", out var panelists) ? panelists?.ToString() : null,
                        HCP = row.TryGetValue("HCP", out var hcp) ? hcp?.ToString() : null,
                        SlideKits = row.TryGetValue("SlideKits", out var slideKits) ? slideKits?.ToString() : null,
                        Invitees = row.TryGetValue("Invitees", out var invitees) ? invitees?.ToString() : null,
                        MIPLInvitees = row.TryGetValue("MIPL Invitees", out var miplInvitees) ? miplInvitees?.ToString() : null,
                        Expenses = row.TryGetValue("Expenses", out var expenses) ? expenses?.ToString() : null,
                        OtherExpenses = row.TryGetValue("Other Expenses", out var otherExpenses) ? otherExpenses?.ToString() : null,
                        TotalInvitees = row.TryGetValue("Total Invitees", out var totalInvitees) && double.TryParse(totalInvitees?.ToString(), out var parsedValue) ? parsedValue : 0,
                        //TotalInvitees = row.TryGetValue("Total Invitees", out var totalInvitees) ? Convert.ToInt32(totalInvitees) : 0,

                        ApprovalStatus = row.TryGetValue("Approval Status", out var approvalStatus) ? approvalStatus?.ToString() : null,
                        NextApprover = row.TryGetValue("Next Approver", out var nextApprover) ? nextApprover?.ToString() : null,
                        RBMBM = row.TryGetValue("RBM/BM", out var rbmbm) ? rbmbm?.ToString() : null,
                        SalesHead = row.TryGetValue("Sales Head", out var salesHead) ? salesHead?.ToString() : null,
                        MarketingHead = row.TryGetValue("Marketing Head", out var marketingHead) ? marketingHead?.ToString() : null,
                        Compliance = row.TryGetValue("Compliance", out var compliance) ? compliance?.ToString() : null,
                        FinanceTreasury = row.TryGetValue("Finance Treasury", out var financeTreasury) ? financeTreasury?.ToString() : null,
                        FinanceAccounts = row.TryGetValue("Finance Accounts", out var financeAccounts) ? financeAccounts?.ToString() : null,
                        MedicalAffairsHead = row.TryGetValue("Medical Affairs Head", out var medicalAffairsHead) ? medicalAffairsHead?.ToString() : null,
                        FinanceHead = row.TryGetValue("Finance Head", out var financeHead) ? financeHead?.ToString() : null,
                        SalesCoordinator = row.TryGetValue("Sales Coordinator", out var salesCoordinator) ? salesCoordinator?.ToString() : null,
                        Finance = row.TryGetValue("Finance", out var finance) ? finance?.ToString() : null,

                        TotalAttendees = row.TryGetValue("Total Attendees", out var totalAttendees) && double.TryParse(totalAttendees?.ToString(), out var parsedTotalAttendees) ? parsedTotalAttendees : 0,
                        TotalHonorariumAmount = row.TryGetValue("Total Honorarium Amount", out var totalHonorariumAmount) && double.TryParse(totalHonorariumAmount?.ToString(), out var parsedTotalHonorariumAmount) ? parsedTotalHonorariumAmount : 0,
                        TotalTravelAccommodationAmount = row.TryGetValue("Total Travel & Accommodation Amount", out var totalTravelAccommodationAmount) && double.TryParse(totalTravelAccommodationAmount?.ToString(), out var parsedTotalTravelAccommodationAmount) ? parsedTotalTravelAccommodationAmount : 0,
                        TotalTravelAmount = row.TryGetValue("Total Travel Amount", out var totalTravelAmount) && double.TryParse(totalTravelAmount?.ToString(), out var parsedTotalTravelAmount) ? parsedTotalTravelAmount : 0,
                        TotalAccommodationAmount = row.TryGetValue("Total Accommodation Amount", out var totalAccommodationAmount) && double.TryParse(totalAccommodationAmount?.ToString(), out var parsedTotalAccommodationAmount) ? parsedTotalAccommodationAmount : 0,
                        TotalLocalConveyance = row.TryGetValue("Total Local Conveyance", out var totalLocalConveyance) && double.TryParse(totalLocalConveyance?.ToString(), out var parsedTotalLocalConveyance) ? parsedTotalLocalConveyance : 0,
                        TotalExpense = row.TryGetValue("Total Expense", out var totalExpense) && double.TryParse(totalExpense?.ToString(), out var parsedTotalExpense) ? parsedTotalExpense : 0,
                        OtherExpenseAmount = row.TryGetValue("Other Expense Amount", out var otherExpenseAmount) && double.TryParse(otherExpenseAmount?.ToString(), out var parsedOtherExpenseAmount) ? parsedOtherExpenseAmount : 0,
                        TotalBudget = row.TryGetValue("Total Budget", out var totalBudget) && double.TryParse(totalBudget?.ToString(), out var parsedTotalBudget) ? parsedTotalBudget : 0,
                        TotalExpenseBTC = row.TryGetValue(" Total Expense BTC", out var totalExpenseBTC) && double.TryParse(totalExpenseBTC?.ToString(), out var parsedTotalExpenseBTC) ? parsedTotalExpenseBTC : 0,
                        TotalExpenseBTE = row.TryGetValue("Total Expense BTE", out var totalExpenseBTE) && double.TryParse(totalExpenseBTE?.ToString(), out var parsedTotalExpenseBTE) ? parsedTotalExpenseBTE : 0,
                        CostperparticipantHelper = row.TryGetValue("Cost per participant - Helper", out var costperparticipantHelper) && double.TryParse(costperparticipantHelper?.ToString(), out var parsedCostperparticipantHelper) ? parsedCostperparticipantHelper : 0,
                        AdvanceAmount = row.TryGetValue("Advance Amount", out var advanceAmount) && double.TryParse(advanceAmount?.ToString(), out var parsedAdvanceAmount) ? parsedAdvanceAmount : 0,
                        TotalHCPRegistrationSpend = row.TryGetValue("Total HCP Registration Spend", out var totalHCPRegistrationSpend) && double.TryParse(totalHCPRegistrationSpend?.ToString(), out var parsedTotalHCPRegistrationSpend) ? parsedTotalHCPRegistrationSpend : 0,
                        TotalHCPRegistrationAmount = row.TryGetValue("Total HCP Registration Amount", out var totalHCPRegistrationAmount) && double.TryParse(totalHCPRegistrationAmount?.ToString(), out var parsedTotalHCPRegistrationAmount) ? parsedTotalHCPRegistrationAmount : 0,
                        FacilitychargesExcludingTax = row.TryGetValue("Facility charges Excluding Tax", out var facilitychargesExcludingTax) && double.TryParse(facilitychargesExcludingTax?.ToString(), out var parsedFacilitychargesExcludingTax) ? parsedFacilitychargesExcludingTax : 0,
                        TotalFacilitychargesincludingTax = row.TryGetValue("Total Facility charges including Tax", out var totalFacilitychargesincludingTax) && double.TryParse(totalFacilitychargesincludingTax?.ToString(), out var parsedTotalFacilitychargesincludingTax) ? parsedTotalFacilitychargesincludingTax : 0,
                        AnesthetistExcludingTax = row.TryGetValue("Anesthetist Excluding Tax", out var anesthetistExcludingTax) && double.TryParse(anesthetistExcludingTax?.ToString(), out var parsedAnesthetistExcludingTax) ? parsedAnesthetistExcludingTax : 0,
                        AnesthetistincludingTax = row.TryGetValue("Anesthetist including Tax", out var anesthetistincludingTax) && double.TryParse(anesthetistincludingTax?.ToString(), out var parsedAnesthetistincludingTax) ? parsedAnesthetistincludingTax : 0,

                        Comments = row.TryGetValue("Comments", out var Comments) ? Comments?.ToString() : null,
                        Objective = row.TryGetValue("Objective", out var Objective) ? Objective?.ToString() : null,
                        IsMSLSelected = row.TryGetValue("Is MSL Selected?", out var IsMSLSelected) ? IsMSLSelected?.ToString() : null,
                        IsProtocolsGiven = row.TryGetValue("Is Protocols Given?", out var IsProtocolsGiven) ? IsProtocolsGiven?.ToString() : null,
                        Agreements = row.TryGetValue("Agreements?", out var Agreements) ? Agreements?.ToString() : null,
                        Complianceapproval = row.TryGetValue("Compliance approval", out var Complianceapproval) ? Complianceapproval?.ToString() : null,
                        TAApproval = row.TryGetValue("T/A Approval", out var TAApproval) ? TAApproval?.ToString() : null,
                        Trainer12LApproval = row.TryGetValue("Trainer 12L Approval", out var Trainer12LApproval) ? Trainer12LApproval?.ToString() : null,
                        HCP6LApproval = row.TryGetValue("HCP 6L Approval", out var HCP6LApproval) ? HCP6LApproval?.ToString() : null,
                        HCP5LApproval = row.TryGetValue("HCP 5L Approval", out var HCP5LApproval) ? HCP5LApproval?.ToString() : null,
                        HCP1LApproved = row.TryGetValue("HCP 1L Approved", out var HCP1LApproved) ? HCP1LApproved?.ToString() : null,
                        fbapproved = row.TryGetValue("f&b approved", out var fbapproved) ? fbapproved?.ToString() : null,
                        pre5daysapproval = row.TryGetValue("pre 5 days approval", out var pre5daysapproval) ? pre5daysapproval?.ToString() : null,
                        pre45daysApproval = row.TryGetValue("pre-45 days Approval", out var pre45daysApproval) ? pre45daysApproval?.ToString() : null,
                        BeneficiaryDetails = row.TryGetValue("Beneficiary Details", out var BeneficiaryDetails) ? BeneficiaryDetails?.ToString() : null,
                        SelectedProducts = row.TryGetValue("Selected Products", out var SelectedProducts) ? SelectedProducts?.ToString() : null,
                        AnesthetistBTCBTE = row.TryGetValue("Anesthetist BTC/BTE", out var AnesthetistBTCBTE) ? AnesthetistBTCBTE?.ToString() : null,
                        AnesthetistRequired = row.TryGetValue("Anesthetist Required?", out var AnesthetistRequired) ? AnesthetistRequired?.ToString() : null,
                        FacilityChargesBTCBTE = row.TryGetValue("Facility Charges BTC/BTE", out var FacilityChargesBTCBTE) ? FacilityChargesBTCBTE?.ToString() : null,
                        EmergencyContactNo = row.TryGetValue("Emergency Contact No", out var EmergencyContactNo) ? EmergencyContactNo?.ToString() : null,
                        EmergencySupport = row.TryGetValue("Emergency Support", out var EmergencySupport) ? EmergencySupport?.ToString() : null,
                        VenueSelectionChecklist = row.TryGetValue("Venue Selection Checklist", out var VenueSelectionChecklist) ? VenueSelectionChecklist?.ToString() : null,
                        HOTWebinarVendorName = row.TryGetValue("HOT Webinar Vendor Name", out var HOTWebinarVendorName) ? HOTWebinarVendorName?.ToString() : null,
                        HOTWebinarType = row.TryGetValue("HOT Webinar Type", out var HOTWebinarType) ? HOTWebinarType?.ToString() : null,
                        FacilityCharges = row.TryGetValue("Facility Charges", out var FacilityCharges) ? FacilityCharges?.ToString() : null,
                        ModeofTraining = row.TryGetValue("Mode of Training", out var ModeofTraining) ? ModeofTraining?.ToString() : null,
                        ProductBrand = row.TryGetValue("Product Brand", out var ProductBrand) ? ProductBrand?.ToString() : null,
                        ValidFrom = row.TryGetValue("Valid From", out var ValidFrom) ? ValidFrom?.ToString() : null,
                        ValidTo = row.TryGetValue("FinaValid Tonce", out var ValidTo) ? ValidTo?.ToString() : null,
                        MedicalUtilityDescription = row.TryGetValue("Medical Utility Description", out var MedicalUtilityDescription) ? MedicalUtilityDescription?.ToString() : null,
                        MedicalUtilityType = row.TryGetValue("Medical Utility Type", out var MedicalUtilityType) ? MedicalUtilityType?.ToString() : null,
                        VenueCountry = row.TryGetValue("Venue Country", out var VenueCountry) ? VenueCountry?.ToString() : null,
                        SponsorshipSocietyName = row.TryGetValue("Sponsorship Society Name", out var SponsorshipSocietyName) ? SponsorshipSocietyName?.ToString() : null,
                        MeetingType = row.TryGetValue("Meeting Type", out var MeetingType) ? MeetingType?.ToString() : null,
                        ClassIIIEventCode = row.TryGetValue("Class III Event Code", out var ClassIIIEventCode) ? ClassIIIEventCode?.ToString() : null,
                        HCPRole = row.TryGetValue("HCP Role", out var HCPRole) ? HCPRole?.ToString() : null,
                        ViewPostEventRequest = row.TryGetValue("View Post Event Request", out var ViewPostEventRequest) ? ViewPostEventRequest?.ToString() : null,
                        EventSettlementDeviationApprovalDate = row.TryGetValue("EventSettlement - Deviation Approval Date", out var EventSettlementDeviationApprovalDate) ? EventSettlementDeviationApprovalDate?.ToString() : null,
                        PREFinanceTreasuryApproval = row.TryGetValue("PRE-Finance Treasury Approval", out var PREFinanceTreasuryApproval) ? PREFinanceTreasuryApproval?.ToString() : null,
                        BankReferenceDate = row.TryGetValue("Bank Reference Date", out var BankReferenceDate) ? BankReferenceDate?.ToString() : null,
                        BankReferenceNumber = row.TryGetValue("Bank Reference Number", out var BankReferenceNumber) ? BankReferenceNumber?.ToString() : null,
                        AdvanceVoucherDate = row.TryGetValue("Advance Voucher Date", out var AdvanceVoucherDate) ? AdvanceVoucherDate?.ToString() : null,
                        AdvanceVoucherNumber = row.TryGetValue("Advance Voucher Number", out var AdvanceVoucherNumber) ? AdvanceVoucherNumber?.ToString() : null,

                    }).ToList();
                return Ok(eventRequestBrandsList);
            }
            catch (Exception ex)
            {
                Log.Error($"Error occurred on AllPreEventsController AttachmentFile method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(ex.Message);
            }
        }

        [HttpGet("GetHonorariumDataUsingEventId")]
        public IActionResult GetHonorariumDataUsingEventId(string eventId)
        {
            string sheetId = configuration.GetSection("SmartsheetSettings:HonorariumPayment").Value;
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
            Dictionary<string, object> ProductBrandsListrowData = new Dictionary<string, object>();
            //Row ExistingRow = sheet.Rows.FirstOrDefault(row => row.)
            Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == eventId));
            if (targetRow != null)
            {
                long Id = targetRow.Id.Value;
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);

                List<HonorariumPayload> eventRequestBrandsList = sheetData
                    .Where(row => row.TryGetValue("EventId/EventRequestId", out var eventIdValue) && eventIdValue?.ToString() == eventId)
                    .Select(row => new HonorariumPayload
                    {
                        EventIdEventRequestId = row.TryGetValue("EventId/EventRequestId", out var eventId) ? eventId?.ToString() : null,
                        EventType = row.TryGetValue("EventType", out var eventType) ? eventType?.ToString() : null,
                        EventTopic = row.TryGetValue("Event Topic", out var eventTopic) ? eventTopic?.ToString() : null,
                        EventDate = row.TryGetValue("EventDate", out var eventDate) ? eventDate?.ToString() : null,
                        StartTime = row.TryGetValue("StartTime", out var startTime) ? startTime?.ToString() : null,
                        EndTime = row.TryGetValue("EndTime", out var endTime) ? endTime?.ToString() : null,
                        EventEndDate = row.TryGetValue("Event End Date", out var eventEndDate) ? eventEndDate?.ToString() : null,
                        VenueName = row.TryGetValue("VenueName", out var venueName) ? venueName?.ToString() : null,
                        InitiatorName = row.TryGetValue("InitiatorName", out var initiatorName) ? initiatorName?.ToString() : null,
                        InitiatorEmail = row.TryGetValue("Initiator Email", out var initiatorEmail) ? initiatorEmail?.ToString() : null,
                        State = row.TryGetValue("State", out var state) ? state?.ToString() : null,
                        City = row.TryGetValue("City", out var city) ? city?.ToString() : null,
                        HonorariumSubmitted = row.TryGetValue("Honorarium Submitted?", out var prerBmBmApproval) ? prerBmBmApproval?.ToString() : null,
                        EventRequestStatus = row.TryGetValue("Event Request Status", out var prerBmBmApprovalDate) ? prerBmBmApprovalDate?.ToString() : null,
                        HonorariumRequestStatus = row.TryGetValue("Honorarium Request Status", out var honorariumRequestStatus) ? honorariumRequestStatus?.ToString() : null,
                        HonorariumApprovedDate = row.TryGetValue("Honorarium Approved Date", out var honorariumApprovedDate) ? honorariumApprovedDate?.ToString() : null,
                        Brands = row.TryGetValue("Brands", out var brands) ? brands?.ToString() : null,
                        Panelists = row.TryGetValue("Panelists", out var panelists) ? panelists?.ToString() : null,
                        SlideKits = row.TryGetValue("SlideKits", out var slideKits) ? slideKits?.ToString() : null,
                        Invitees = row.TryGetValue("Invitees", out var invitees) ? invitees?.ToString() : null,
                        Expenses = row.TryGetValue("Expenses", out var expenses) ? expenses?.ToString() : null,
                        PanelistsAgreements = row.TryGetValue("Panelists & Agreements", out var presalesHeadApproval) ? presalesHeadApproval?.ToString() : null,
                        HCPAgreements = row.TryGetValue("HCP & Agreements", out var presalesHeadApprovalDate) ? presalesHeadApprovalDate?.ToString() : null,
                        IsAdvanceRequired = row.TryGetValue("IsAdvanceRequired", out var isAdvanceRequired) ? isAdvanceRequired?.ToString() : null,
                        EventOpen30days = row.TryGetValue("EventOpen30days", out var eventOpen30days) ? eventOpen30days?.ToString() : null,
                        EventWithin7days = row.TryGetValue("EventWithin7days", out var eventWithin7days) ? eventWithin7days?.ToString() : null,
                        CreatedOn = row.TryGetValue("CreatedOn", out var premarketingHeadApproval) ? premarketingHeadApproval?.ToString() : null,
                        CreatedDateHelper = row.TryGetValue("Created Date - Helper", out var createdDateHelper) ? createdDateHelper?.ToString() : null,
                        Modified = row.TryGetValue("Modified", out var premarketingHeadApprovalDate) ? premarketingHeadApprovalDate?.ToString() : null,
                        HONRBMBMApproval = row.TryGetValue("HON-RBM/BM Approval", out var agreement) ? agreement?.ToString() : null,
                        HONRBMBMApprovalDate = row.TryGetValue("HON-RBM/BM Approval Date", out var HONRBMBMApprovalDate) ? HONRBMBMApprovalDate?.ToString() : null,
                        HONComplianceApproval = row.TryGetValue("HON-Compliance Approval", out var HONComplianceApproval) ? HONComplianceApproval?.ToString() : null,
                        HONComplianceApprovalDate = row.TryGetValue("HON-Compliance Approval Date", out var HONComplianceApprovalDate) ? HONComplianceApprovalDate?.ToString() : null,
                        HONFinanceAccountsApproval = row.TryGetValue("HON-Finance Accounts Approval", out var HONFinanceAccountsApproval) ? HONFinanceAccountsApproval?.ToString() : null,
                        FinanceRejectionComments = row.TryGetValue("Finance Rejection Comments", out var FinanceRejectionComments) ? FinanceRejectionComments?.ToString() : null,
                        HONFinanceAccountsApprovalDate = row.TryGetValue("HON-Finance Accounts Approval Date", out var HONFinanceAccountsApprovalDate) ? HONFinanceAccountsApprovalDate?.ToString() : null,
                        HONFinanceTreasuryApproval = row.TryGetValue("HON-Finance Treasury Approval", out var HONFinanceTreasuryApproval) ? HONFinanceTreasuryApproval?.ToString() : null,
                        FinanceTreasuryRejectionComments = row.TryGetValue("Finance Treasury Rejection Comments", out var FinanceTreasuryRejectionComments) ? FinanceTreasuryRejectionComments?.ToString() : null,
                        HONFinanceTreasuryApprovalDate = row.TryGetValue("HON-Finance Treasury Approval Date", out var HONFinanceTreasuryApprovalDate) ? HONFinanceTreasuryApprovalDate?.ToString() : null,
                        HONMarketingHeadApproval = row.TryGetValue("HON-Marketing Head Approval", out var HONMarketingHeadApproval) ? HONMarketingHeadApproval?.ToString() : null,
                        HONMarketingHeadApprovalDate = row.TryGetValue("HON-Marketing Head Approval Date", out var HONMarketingHeadApprovalDate) ? HONMarketingHeadApprovalDate?.ToString() : null,
                        HONSalesHeadApproval = row.TryGetValue("HON-Sales Head Approval", out var HONSalesHeadApproval) ? HONSalesHeadApproval?.ToString() : null,
                        HONSalesHeadApprovalDate = row.TryGetValue("HON-Sales Head Approval Date", out var HONSalesHeadApprovalDate) ? HONSalesHeadApprovalDate?.ToString() : null,
                        HONMedicalAffairsHeadApproval = row.TryGetValue("HON-Medical Affairs Head Approval", out var HONMedicalAffairsHeadApproval) ? HONMedicalAffairsHeadApproval?.ToString() : null,
                        HONMedicalAffairsHeadApprovalDate = row.TryGetValue("HON-Medical Affairs Head Approval Date", out var HONMedicalAffairsHeadApprovalDate) ? HONMedicalAffairsHeadApprovalDate?.ToString() : null,
                        _5workingdays = row.TryGetValue("5working days", out var prefFinanceTreasuryApprovalDate) ? prefFinanceTreasuryApprovalDate?.ToString() : null,
                        HON2WorkingdaysDeviationApproval = row.TryGetValue("HON-2Workingdays Deviation Approval", out var premedicalAffairsHeadApproval) ? premedicalAffairsHeadApproval?.ToString() : null,
                        HON2WorkingdaysDeviationApprovalDate = row.TryGetValue("HON-2Workingdays Deviation Approval Date", out var premedicalAffairsHeadApprovalDate) ? premedicalAffairsHeadApprovalDate?.ToString() : null,
                        HONLessthan5inviteesDeviationApproval = row.TryGetValue("HON-Lessthan5invitees Deviation Approval", out var presalesCoordinatorApproval) ? presalesCoordinatorApproval?.ToString() : null,
                        HONLessthan5inviteesDeviationApprovalDate = row.TryGetValue("HON-Lessthan5invitees Deviation Approval Date", out var presalesCoordinatorApprovalDate) ? presalesCoordinatorApprovalDate?.ToString() : null,
                        FinanceAccountsGivenDetails = row.TryGetValue("Finance Accounts Given Details", out var precomplianceApproval) ? precomplianceApproval?.ToString() : null,
                        FinanceTreasuryGivenDetails = row.TryGetValue("Finance Treasury Given Details", out var precomplianceApprovalDate) ? precomplianceApprovalDate?.ToString() : null,
                        JVNo = row.TryGetValue("JV No", out var eventRequestStatus) ? eventRequestStatus?.ToString() : null,
                        JVDate = row.TryGetValue("JV Date", out var eventApprovedDate) ? eventApprovedDate?.ToString() : null,
                        PVNo = row.TryGetValue("PV No", out var miplInvitees) ? miplInvitees?.ToString() : null,
                        PVDate = row.TryGetValue("PV Date", out var postEventRequeststatus) ? postEventRequeststatus?.ToString() : null,
                        ActualAmountGreaterThan50 = row.TryGetValue("ActualAmountGreaterThan50%", out var helperFinancetreasurytriggerBTE) ? helperFinancetreasurytriggerBTE?.ToString() : null,
                        ReportingManager = row.TryGetValue("Reporting Manager", out var postEventApprovedDate) ? postEventApprovedDate?.ToString() : null,
                        _1UpManager = row.TryGetValue("1 Up Manager", out var _1UpManager) ? _1UpManager?.ToString() : null,
                        SalesHeadapproval = row.TryGetValue("Sales Head approval", out var eventOpenSalesHeadApproval) ? eventOpenSalesHeadApproval?.ToString() : null,
                        ApprovalStatus = row.TryGetValue("Approval Status", out var eventOpenSalesHeadApprovalDate) ? eventOpenSalesHeadApprovalDate?.ToString() : null,
                        NextApprover = row.TryGetValue("Next Approver", out var _7daysSalesHeadApproval) ? _7daysSalesHeadApproval?.ToString() : null,
                        RBMBM = row.TryGetValue("RBM/BM", out var rbmbm) ? rbmbm?.ToString() : null,
                        SalesHead = row.TryGetValue("Sales Head", out var salesHead) ? salesHead?.ToString() : null,
                        MarketingHead = row.TryGetValue("Marketing Head", out var marketingHead) ? marketingHead?.ToString() : null,
                        Finance = row.TryGetValue("Finance", out var finance) ? finance?.ToString() : null,
                        Compliance = row.TryGetValue("Compliance", out var compliance) ? compliance?.ToString() : null,
                        FinanceTreasury = row.TryGetValue("Finance Treasury", out var financeTreasury) ? financeTreasury?.ToString() : null,
                        FinanceAccounts = row.TryGetValue("Finance Accounts", out var financeAccounts) ? financeAccounts?.ToString() : null,
                        MedicalAffairsHead = row.TryGetValue("Medical Affairs Head", out var medicalAffairsHead) ? medicalAffairsHead?.ToString() : null,
                        SalesCoordinator = row.TryGetValue("Sales Coordinator", out var salesCoordinator) ? salesCoordinator?.ToString() : null,
                        FinanceHead = row.TryGetValue("Finance Head", out var financeHead) ? financeHead?.ToString() : null,
                        CostperparticipantHelper = row.TryGetValue("Cost per participant - Helper", out var _7daysSalesHeadApprovaldate) ? _7daysSalesHeadApprovaldate?.ToString() : null,
                        BrandName = row.TryGetValue("BrandName", out var prefBExpenseExcludingTaxApproval) ? prefBExpenseExcludingTaxApproval?.ToString() : null,
                        Allocation = row.TryGetValue("% Allocation", out var hcpExceeds100000FHApproval) ? hcpExceeds100000FHApproval?.ToString() : null,
                        ProjectID = row.TryGetValue("Project ID", out var hcpExceeds500000TriggerFHApproval) ? hcpExceeds500000TriggerFHApproval?.ToString() : null,
                        Honorarium = row.TryGetValue("Honorarium", out var hcpHonorarium600000ExceededApproval) ? hcpHonorarium600000ExceededApproval?.ToString() : null,
                        TotalInvitees = row.TryGetValue("Total Invitees", out var totalInvitees) && double.TryParse(totalInvitees?.ToString(), out var parsedValue) ? parsedValue : 0,
                        TotalAttendees = row.TryGetValue("Total Attendees", out var totalAttendees) && double.TryParse(totalAttendees?.ToString(), out var parsedTotalAttendees) ? parsedTotalAttendees : 0,
                        TotalHonorariumAmount = row.TryGetValue("Total Honorarium Amount", out var totalHonorariumAmount) && double.TryParse(totalHonorariumAmount?.ToString(), out var parsedTotalHonorariumAmount) ? parsedTotalHonorariumAmount : 0,
                        TotalTravelAccommodationAmount = row.TryGetValue("Total Travel & Accommodation Amount", out var totalTravelAccommodationAmount) && double.TryParse(totalTravelAccommodationAmount?.ToString(), out var parsedTotalTravelAccommodationAmount) ? parsedTotalTravelAccommodationAmount : 0,
                        TotalTravelAmount = row.TryGetValue("Total Travel Amount", out var totalTravelAmount) && double.TryParse(totalTravelAmount?.ToString(), out var parsedTotalTravelAmount) ? parsedTotalTravelAmount : 0,
                        TotalAccommodationAmount = row.TryGetValue("Total Accommodation Amount", out var totalAccommodationAmount) && double.TryParse(totalAccommodationAmount?.ToString(), out var parsedTotalAccommodationAmount) ? parsedTotalAccommodationAmount : 0,
                        TotalLocalConveyance = row.TryGetValue("Total Local Conveyance", out var totalLocalConveyance) && double.TryParse(totalLocalConveyance?.ToString(), out var parsedTotalLocalConveyance) ? parsedTotalLocalConveyance : 0,
                        TotalExpense = row.TryGetValue("Total Expense", out var totalExpense) && double.TryParse(totalExpense?.ToString(), out var parsedTotalExpense) ? parsedTotalExpense : 0,
                        TotalBudget = row.TryGetValue("Total Budget", out var totalBudget) && double.TryParse(totalBudget?.ToString(), out var parsedTotalBudget) ? parsedTotalBudget : 0,
                        HCPRole = row.TryGetValue("HCP Role", out var HCPRole) ? HCPRole?.ToString() : null,
                        Role = row.TryGetValue("Role", out var Role) ? Role?.ToString() : null,
                        AdvanceAmount = row.TryGetValue("Advance Amount", out var advanceAmount) && double.TryParse(advanceAmount?.ToString(), out var parsedAdvanceAmount) ? parsedAdvanceAmount : 0,
                        TotalExpenseBTC = row.TryGetValue(" Total Expense BTC", out var totalExpenseBTC) && double.TryParse(totalExpenseBTC?.ToString(), out var parsedTotalExpenseBTC) ? parsedTotalExpenseBTC : 0,
                        TotalExpenseBTE = row.TryGetValue("Total Expense BTE", out var totalExpenseBTE) && double.TryParse(totalExpenseBTE?.ToString(), out var parsedTotalExpenseBTE) ? parsedTotalExpenseBTE : 0,
                        MeetingType = row.TryGetValue("Meeting Type", out var MeetingType) ? MeetingType?.ToString() : null,

                    }).ToList();

                return Ok(eventRequestBrandsList);
            }
            else
            {
                return BadRequest(new
                {
                    Message = "Row not found"
                });
            }
            return Ok();
        }
        [HttpGet("GetDataFromHonorariumDataSheet")]
        public IActionResult GetDataFromHonorariumDataSheet()
        {
            string sheetId = configuration.GetSection("SmartsheetSettings:HonorariumPayment").Value;
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
            Dictionary<string, object> ProductBrandsListrowData = new Dictionary<string, object>();
            //Row ExistingRow = sheet.Rows.FirstOrDefault(row => row.)
            // Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == eventId));
            //if (targetRow != null)
            //{
            //    long Id = targetRow.Id.Value;
            List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);

            List<HonorariumPayload> eventRequestBrandsList = sheetData
                //.Where(row => row.TryGetValue("EventId/EventRequestId", out var eventIdValue) && eventIdValue?.ToString() == eventId)
                .Select(row => new HonorariumPayload
                {
                    EventIdEventRequestId = row.TryGetValue("EventId/EventRequestId", out var eventId) ? eventId?.ToString() : null,
                    EventType = row.TryGetValue("Event Type", out var eventType) ? eventType?.ToString() : null,
                    EventTopic = row.TryGetValue("Event Topic", out var eventTopic) ? eventTopic?.ToString() : null,
                    EventDate = row.TryGetValue("Event Date", out var eventDate) ? eventDate?.ToString() : null,
                    StartTime = row.TryGetValue("Start Time", out var startTime) ? startTime?.ToString() : null,
                    EndTime = row.TryGetValue("End Time", out var endTime) ? endTime?.ToString() : null,
                    EventEndDate = row.TryGetValue("Event End Date", out var eventEndDate) ? eventEndDate?.ToString() : null,
                    VenueName = row.TryGetValue("Venue Name", out var venueName) ? venueName?.ToString() : null,
                    InitiatorName = row.TryGetValue("Initiator Name", out var initiatorName) ? initiatorName?.ToString() : null,
                    InitiatorEmail = row.TryGetValue("Initiator Email", out var initiatorEmail) ? initiatorEmail?.ToString() : null,
                    State = row.TryGetValue("State", out var state) ? state?.ToString() : null,
                    City = row.TryGetValue("City", out var city) ? city?.ToString() : null,
                    HonorariumSubmitted = row.TryGetValue("Honorarium Submitted?", out var prerBmBmApproval) ? prerBmBmApproval?.ToString() : null,
                    EventRequestStatus = row.TryGetValue("Event Request Status", out var prerBmBmApprovalDate) ? prerBmBmApprovalDate?.ToString() : null,
                    HonorariumRequestStatus = row.TryGetValue("Honorarium Request Status", out var honorariumRequestStatus) ? honorariumRequestStatus?.ToString() : null,
                    HonorariumApprovedDate = row.TryGetValue("Honorarium Approved Date", out var honorariumApprovedDate) ? honorariumApprovedDate?.ToString() : null,
                    Brands = row.TryGetValue("Brands", out var brands) ? brands?.ToString() : null,
                    Panelists = row.TryGetValue("Panelists", out var panelists) ? panelists?.ToString() : null,
                    SlideKits = row.TryGetValue("SlideKits", out var slideKits) ? slideKits?.ToString() : null,
                    Invitees = row.TryGetValue("Invitees", out var invitees) ? invitees?.ToString() : null,
                    Expenses = row.TryGetValue("Expenses", out var expenses) ? expenses?.ToString() : null,
                    PanelistsAgreements = row.TryGetValue("Panelists & Agreements", out var presalesHeadApproval) ? presalesHeadApproval?.ToString() : null,
                    HCPAgreements = row.TryGetValue("HCP & Agreements", out var presalesHeadApprovalDate) ? presalesHeadApprovalDate?.ToString() : null,
                    IsAdvanceRequired = row.TryGetValue("IsAdvanceRequired", out var isAdvanceRequired) ? isAdvanceRequired?.ToString() : null,
                    EventOpen30days = row.TryGetValue("EventOpen30days", out var eventOpen30days) ? eventOpen30days?.ToString() : null,
                    EventWithin7days = row.TryGetValue("EventWithin7days", out var eventWithin7days) ? eventWithin7days?.ToString() : null,
                    CreatedOn = row.TryGetValue("CreatedOn", out var premarketingHeadApproval) ? premarketingHeadApproval?.ToString() : null,
                    CreatedDateHelper = row.TryGetValue("Created Date - Helper", out var createdDateHelper) ? createdDateHelper?.ToString() : null,
                    Modified = row.TryGetValue("Modified", out var premarketingHeadApprovalDate) ? premarketingHeadApprovalDate?.ToString() : null,
                    HONRBMBMApproval = row.TryGetValue("HON-RBM/BM Approval", out var agreement) ? agreement?.ToString() : null,
                    HONRBMBMApprovalDate = row.TryGetValue("HON-RBM/BM Approval Date", out var HONRBMBMApprovalDate) ? HONRBMBMApprovalDate?.ToString() : null,
                    HONComplianceApproval = row.TryGetValue("HON-Compliance Approval", out var HONComplianceApproval) ? HONComplianceApproval?.ToString() : null,
                    HONComplianceApprovalDate = row.TryGetValue("HON-Compliance Approval Date", out var HONComplianceApprovalDate) ? HONComplianceApprovalDate?.ToString() : null,
                    HONFinanceAccountsApproval = row.TryGetValue("HON-Finance Accounts Approval", out var HONFinanceAccountsApproval) ? HONFinanceAccountsApproval?.ToString() : null,
                    FinanceRejectionComments = row.TryGetValue("Finance Rejection Comments", out var FinanceRejectionComments) ? FinanceRejectionComments?.ToString() : null,
                    HONFinanceAccountsApprovalDate = row.TryGetValue("HON-Finance Accounts Approval Date", out var HONFinanceAccountsApprovalDate) ? HONFinanceAccountsApprovalDate?.ToString() : null,
                    HONFinanceTreasuryApproval = row.TryGetValue("HON-Finance Treasury Approval", out var HONFinanceTreasuryApproval) ? HONFinanceTreasuryApproval?.ToString() : null,
                    FinanceTreasuryRejectionComments = row.TryGetValue("Finance Treasury Rejection Comments", out var FinanceTreasuryRejectionComments) ? FinanceTreasuryRejectionComments?.ToString() : null,
                    HONFinanceTreasuryApprovalDate = row.TryGetValue("HON-Finance Treasury Approval Date", out var HONFinanceTreasuryApprovalDate) ? HONFinanceTreasuryApprovalDate?.ToString() : null,
                    HONMarketingHeadApproval = row.TryGetValue("HON-Marketing Head Approval", out var HONMarketingHeadApproval) ? HONMarketingHeadApproval?.ToString() : null,
                    HONMarketingHeadApprovalDate = row.TryGetValue("HON-Marketing Head Approval Date", out var HONMarketingHeadApprovalDate) ? HONMarketingHeadApprovalDate?.ToString() : null,
                    HONSalesHeadApproval = row.TryGetValue("HON-Sales Head Approval", out var HONSalesHeadApproval) ? HONSalesHeadApproval?.ToString() : null,
                    HONSalesHeadApprovalDate = row.TryGetValue("HON-Sales Head Approval Date", out var HONSalesHeadApprovalDate) ? HONSalesHeadApprovalDate?.ToString() : null,
                    HONMedicalAffairsHeadApproval = row.TryGetValue("HON-Medical Affairs Head Approval", out var HONMedicalAffairsHeadApproval) ? HONMedicalAffairsHeadApproval?.ToString() : null,
                    HONMedicalAffairsHeadApprovalDate = row.TryGetValue("HON-Medical Affairs Head Approval Date", out var HONMedicalAffairsHeadApprovalDate) ? HONMedicalAffairsHeadApprovalDate?.ToString() : null,
                    _5workingdays = row.TryGetValue("5working days", out var prefFinanceTreasuryApprovalDate) ? prefFinanceTreasuryApprovalDate?.ToString() : null,
                    HON2WorkingdaysDeviationApproval = row.TryGetValue("HON-2Workingdays Deviation Approval", out var premedicalAffairsHeadApproval) ? premedicalAffairsHeadApproval?.ToString() : null,
                    HON2WorkingdaysDeviationApprovalDate = row.TryGetValue("HON-2Workingdays Deviation Approval Date", out var premedicalAffairsHeadApprovalDate) ? premedicalAffairsHeadApprovalDate?.ToString() : null,
                    HONLessthan5inviteesDeviationApproval = row.TryGetValue("HON-Lessthan5invitees Deviation Approval", out var presalesCoordinatorApproval) ? presalesCoordinatorApproval?.ToString() : null,
                    HONLessthan5inviteesDeviationApprovalDate = row.TryGetValue("HON-Lessthan5invitees Deviation Approval Date", out var presalesCoordinatorApprovalDate) ? presalesCoordinatorApprovalDate?.ToString() : null,
                    FinanceAccountsGivenDetails = row.TryGetValue("Finance Accounts Given Details", out var precomplianceApproval) ? precomplianceApproval?.ToString() : null,
                    FinanceTreasuryGivenDetails = row.TryGetValue("Finance Treasury Given Details", out var precomplianceApprovalDate) ? precomplianceApprovalDate?.ToString() : null,
                    JVNo = row.TryGetValue("JV No", out var eventRequestStatus) ? eventRequestStatus?.ToString() : null,
                    JVDate = row.TryGetValue("JV Date", out var eventApprovedDate) ? eventApprovedDate?.ToString() : null,
                    PVNo = row.TryGetValue("PV No", out var miplInvitees) ? miplInvitees?.ToString() : null,
                    PVDate = row.TryGetValue("PV Date", out var postEventRequeststatus) ? postEventRequeststatus?.ToString() : null,
                    ActualAmountGreaterThan50 = row.TryGetValue("ActualAmountGreaterThan50%", out var helperFinancetreasurytriggerBTE) ? helperFinancetreasurytriggerBTE?.ToString() : null,
                    ReportingManager = row.TryGetValue("Reporting Manager", out var postEventApprovedDate) ? postEventApprovedDate?.ToString() : null,
                    _1UpManager = row.TryGetValue("1 Up Manager", out var _1UpManager) ? _1UpManager?.ToString() : null,
                    SalesHeadapproval = row.TryGetValue("Sales Head approval", out var eventOpenSalesHeadApproval) ? eventOpenSalesHeadApproval?.ToString() : null,
                    ApprovalStatus = row.TryGetValue("Approval Status", out var eventOpenSalesHeadApprovalDate) ? eventOpenSalesHeadApprovalDate?.ToString() : null,
                    NextApprover = row.TryGetValue("Next Approver", out var _7daysSalesHeadApproval) ? _7daysSalesHeadApproval?.ToString() : null,
                    RBMBM = row.TryGetValue("RBM/BM", out var rbmbm) ? rbmbm?.ToString() : null,
                    SalesHead = row.TryGetValue("Sales Head", out var salesHead) ? salesHead?.ToString() : null,
                    MarketingHead = row.TryGetValue("Marketing Head", out var marketingHead) ? marketingHead?.ToString() : null,
                    Finance = row.TryGetValue("Finance", out var finance) ? finance?.ToString() : null,
                    Compliance = row.TryGetValue("Compliance", out var compliance) ? compliance?.ToString() : null,
                    FinanceTreasury = row.TryGetValue("Finance Treasury", out var financeTreasury) ? financeTreasury?.ToString() : null,
                    FinanceAccounts = row.TryGetValue("Finance Accounts", out var financeAccounts) ? financeAccounts?.ToString() : null,
                    MedicalAffairsHead = row.TryGetValue("Medical Affairs Head", out var medicalAffairsHead) ? medicalAffairsHead?.ToString() : null,
                    SalesCoordinator = row.TryGetValue("Sales Coordinator", out var salesCoordinator) ? salesCoordinator?.ToString() : null,
                    FinanceHead = row.TryGetValue("Finance Head", out var financeHead) ? financeHead?.ToString() : null,
                    CostperparticipantHelper = row.TryGetValue("Cost per participant - Helper", out var _7daysSalesHeadApprovaldate) ? _7daysSalesHeadApprovaldate?.ToString() : null,
                    BrandName = row.TryGetValue("BrandName", out var prefBExpenseExcludingTaxApproval) ? prefBExpenseExcludingTaxApproval?.ToString() : null,
                    Allocation = row.TryGetValue("% Allocation", out var hcpExceeds100000FHApproval) ? hcpExceeds100000FHApproval?.ToString() : null,
                    ProjectID = row.TryGetValue("Project ID", out var hcpExceeds500000TriggerFHApproval) ? hcpExceeds500000TriggerFHApproval?.ToString() : null,
                    Honorarium = row.TryGetValue("Honorarium", out var hcpHonorarium600000ExceededApproval) ? hcpHonorarium600000ExceededApproval?.ToString() : null,
                    TotalInvitees = row.TryGetValue("Total Invitees", out var totalInvitees) && double.TryParse(totalInvitees?.ToString(), out var parsedValue) ? parsedValue : 0,
                    TotalAttendees = row.TryGetValue("Total Attendees", out var totalAttendees) && double.TryParse(totalAttendees?.ToString(), out var parsedTotalAttendees) ? parsedTotalAttendees : 0,
                    TotalHonorariumAmount = row.TryGetValue("Total Honorarium Amount", out var totalHonorariumAmount) && double.TryParse(totalHonorariumAmount?.ToString(), out var parsedTotalHonorariumAmount) ? parsedTotalHonorariumAmount : 0,
                    TotalTravelAccommodationAmount = row.TryGetValue("Total Travel & Accommodation Amount", out var totalTravelAccommodationAmount) && double.TryParse(totalTravelAccommodationAmount?.ToString(), out var parsedTotalTravelAccommodationAmount) ? parsedTotalTravelAccommodationAmount : 0,
                    TotalTravelAmount = row.TryGetValue("Total Travel Amount", out var totalTravelAmount) && double.TryParse(totalTravelAmount?.ToString(), out var parsedTotalTravelAmount) ? parsedTotalTravelAmount : 0,
                    TotalAccommodationAmount = row.TryGetValue("Total Accommodation Amount", out var totalAccommodationAmount) && double.TryParse(totalAccommodationAmount?.ToString(), out var parsedTotalAccommodationAmount) ? parsedTotalAccommodationAmount : 0,
                    TotalLocalConveyance = row.TryGetValue("Total Local Conveyance", out var totalLocalConveyance) && double.TryParse(totalLocalConveyance?.ToString(), out var parsedTotalLocalConveyance) ? parsedTotalLocalConveyance : 0,
                    TotalExpense = row.TryGetValue("Total Expense", out var totalExpense) && double.TryParse(totalExpense?.ToString(), out var parsedTotalExpense) ? parsedTotalExpense : 0,
                    TotalBudget = row.TryGetValue("Total Budget", out var totalBudget) && double.TryParse(totalBudget?.ToString(), out var parsedTotalBudget) ? parsedTotalBudget : 0,
                    HCPRole = row.TryGetValue("HCP Role", out var HCPRole) ? HCPRole?.ToString() : null,
                    Role = row.TryGetValue("Role", out var Role) ? Role?.ToString() : null,
                    AdvanceAmount = row.TryGetValue("Advance Amount", out var advanceAmount) && double.TryParse(advanceAmount?.ToString(), out var parsedAdvanceAmount) ? parsedAdvanceAmount : 0,
                    TotalExpenseBTC = row.TryGetValue(" Total Expense BTC", out var totalExpenseBTC) && double.TryParse(totalExpenseBTC?.ToString(), out var parsedTotalExpenseBTC) ? parsedTotalExpenseBTC : 0,
                    TotalExpenseBTE = row.TryGetValue("Total Expense BTE", out var totalExpenseBTE) && double.TryParse(totalExpenseBTE?.ToString(), out var parsedTotalExpenseBTE) ? parsedTotalExpenseBTE : 0,
                    MeetingType = row.TryGetValue("Meeting Type", out var MeetingType) ? MeetingType?.ToString() : null,

                }).ToList();

            return Ok(eventRequestBrandsList);
            //}
            //else
            //{
            //    return BadRequest(new
            //    {
            //        Message = "Row not found"
            //    });
            //}
            //return Ok();
        }

        [HttpGet("GetDeviationSheetDataUsingEventId")]
        public IActionResult GetDeviationSheetDataUsingEventId(string eventId)
        {
            string sheetId = configuration.GetSection("SmartsheetSettings:DeviationProcess").Value;
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
            Dictionary<string, object> ProductBrandsListrowData = new Dictionary<string, object>();
            //Row ExistingRow = sheet.Rows.FirstOrDefault(row => row.)
            Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == eventId));
            if (targetRow != null)
            {
                long Id = targetRow.Id.Value;
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);

                List<DeviationDataPayload> eventRequestBrandsList = sheetData
                    .Where(row => row.TryGetValue("EventId/EventRequestId", out var eventIdValue) && eventIdValue?.ToString() == eventId)
                    .Select(row => new DeviationDataPayload
                    {
                        DeviationID = row.TryGetValue("Deviation ID", out var DeviationID) ? DeviationID?.ToString() : null,
                        EventIdEventRequestId = row.TryGetValue("EventId/EventRequestId", out var eventId) ? eventId?.ToString() : null,
                        EventTopic = row.TryGetValue("Event Topic", out var eventTopic) ? eventTopic?.ToString() : null,
                        MISCode = row.TryGetValue("MIS Code", out var MISCode) ? MISCode?.ToString() : null,
                        HCPName = row.TryGetValue("HCP Name", out var HCPName) ? HCPName?.ToString() : null,
                        HonorariumAmount = row.TryGetValue("Honorarium Amount", out var HonorariumAmount) ? HonorariumAmount?.ToString() : null,
                        TravelAccommodationAmount = row.TryGetValue("Travel & Accommodation Amount", out var TravelAccommodationAmount) ? TravelAccommodationAmount?.ToString() : null,
                        OtherExpenses = row.TryGetValue("Other Expenses", out var OtherExpenses) ? OtherExpenses?.ToString() : null,
                        DeviationType = row.TryGetValue("Deviation Type", out var DeviationType) ? DeviationType?.ToString() : null,
                        EventLevel = row.TryGetValue("Event Level", out var EventLevel) ? EventLevel?.ToString() : null,
                        OutstandingEvents = row.TryGetValue("Outstanding Events", out var OutstandingEvents) ? OutstandingEvents?.ToString() : null,
                        InitiatorName = row.TryGetValue("InitiatorName", out var InitiatorName) ? InitiatorName?.ToString() : null,
                        InitiatorEmail = row.TryGetValue("Initiator Email", out var initiatorEmail) ? initiatorEmail?.ToString() : null,
                        CreatedOn = row.TryGetValue("CreatedOn", out var premarketingHeadApproval) ? premarketingHeadApproval?.ToString() : null,
                        Modified = row.TryGetValue("Modified", out var premarketingHeadApprovalDate) ? premarketingHeadApprovalDate?.ToString() : null,
                        EventType = row.TryGetValue("EventType", out var eventType) ? eventType?.ToString() : null,
                        EventDate = row.TryGetValue("EventDate", out var eventDate) ? eventDate?.ToString() : null,
                        StartTime = row.TryGetValue("StartTime", out var startTime) ? startTime?.ToString() : null,
                        EndTime = row.TryGetValue("EndTime", out var endTime) ? endTime?.ToString() : null,
                        EventEndDate = row.TryGetValue("Event End Date", out var eventEndDate) ? eventEndDate?.ToString() : null,
                        VenueName = row.TryGetValue("VenueName", out var venueName) ? venueName?.ToString() : null,
                        State = row.TryGetValue("State", out var state) ? state?.ToString() : null,
                        City = row.TryGetValue("City", out var city) ? city?.ToString() : null,
                        IsAdvanceRequired = row.TryGetValue("IsAdvanceRequired", out var isAdvanceRequired) ? isAdvanceRequired?.ToString() : null,
                        SalesHeadapproval = row.TryGetValue("Sales Head approval", out var eventOpenSalesHeadApproval) ? eventOpenSalesHeadApproval?.ToString() : null,
                        SalesHeadapprovalDate = row.TryGetValue("Sales Head approval Date", out var SalesHeadapprovalDate) ? SalesHeadapprovalDate?.ToString() : null,
                        FinanceHeadApproval = row.TryGetValue("Finance Head Approval", out var FinanceHeadApproval) ? FinanceHeadApproval?.ToString() : null,
                        FinanceHeadApprovalDate = row.TryGetValue("Finance Head Approval Date", out var FinanceHeadApprovalDate) ? FinanceHeadApprovalDate?.ToString() : null,
                        ApprovalStatus = row.TryGetValue("Approval Status", out var eventOpenSalesHeadApprovalDate) ? eventOpenSalesHeadApprovalDate?.ToString() : null,
                        HON5WorkingdaysDeviationDateTrigger = row.TryGetValue("HON-5Workingdays Deviation Date Trigger", out var HON5WorkingdaysDeviationDateTrigger) ? HON5WorkingdaysDeviationDateTrigger?.ToString() : null,
                        HON5WorkingdaysDeviationApproval = row.TryGetValue("HON-5Workingdays Deviation Approval", out var HON5WorkingdaysDeviationApproval) ? HON5WorkingdaysDeviationApproval?.ToString() : null,
                        HON5WorkingdaysDeviationApprovalDate = row.TryGetValue("HON-5Workingdays Deviation Approval Date", out var HON5WorkingdaysDeviationApprovalDate) ? HON5WorkingdaysDeviationApprovalDate?.ToString() : null,
                        POSTBeyond45DaysDeviationDateTrigger = row.TryGetValue("POST- Beyond45Days Deviation Date Trigger", out var POSTBeyond45DaysDeviationDateTrigger) ? POSTBeyond45DaysDeviationDateTrigger?.ToString() : null,
                        POSTBeyond45DaysDeviationApproval = row.TryGetValue("POST- Beyond45Days Deviation Approval", out var POSTBeyond45DaysDeviationApproval) ? POSTBeyond45DaysDeviationApproval?.ToString() : null,
                        POSTBeyond45DaysDeviationApprovalDate = row.TryGetValue("POST- Beyond45Days Deviation Approval Date", out var POSTBeyond45DaysDeviationApprovalDate) ? POSTBeyond45DaysDeviationApprovalDate?.ToString() : null,
                        POSTLessthan5InviteesDeviationTrigger = row.TryGetValue("POST-Lessthan5Invitees Deviation Trigger", out var POSTLessthan5InviteesDeviationTrigger) ? POSTLessthan5InviteesDeviationTrigger?.ToString() : null,
                        POSTLessthan5inviteesDeviationApproval = row.TryGetValue("POST-Lessthan5invitees Deviation Approval", out var POSTLessthan5inviteesDeviationApproval) ? POSTLessthan5inviteesDeviationApproval?.ToString() : null,
                        POSTLessthan5inviteesDeviationApprovalDate = row.TryGetValue("POST-Lessthan5invitees Deviation Approval Date", out var POSTLessthan5inviteesDeviationApprovalDate) ? POSTLessthan5inviteesDeviationApprovalDate?.ToString() : null,
                        POSTDeviationCostperpaxabove1500Trigger = row.TryGetValue("POST-Deviation Costperpaxabove1500 Trigger", out var POSTDeviationCostperpaxabove1500Trigger) ? POSTDeviationCostperpaxabove1500Trigger?.ToString() : null,
                        POSTDeviationCostperpaxabove1500Approval = row.TryGetValue("POST-Deviation Costperpaxabove1500 Approval ", out var POSTDeviationCostperpaxabove1500Approval) ? POSTDeviationCostperpaxabove1500Approval?.ToString() : null,
                        POSTDeviationCostperpaxabove1500ApprovalDate = row.TryGetValue("POST-Deviation Costperpaxabove1500 Approval Date", out var POSTDeviationCostperpaxabove1500ApprovalDate) ? POSTDeviationCostperpaxabove1500ApprovalDate?.ToString() : null,
                        POSTDeviationCostperpaxabove2250Trigger = row.TryGetValue("POST-Deviation Costperpaxabove2250 Trigger", out var POSTDeviationCostperpaxabove2250Trigger) ? POSTDeviationCostperpaxabove2250Trigger?.ToString() : null,
                        POSTDeviationCostperpaxabove2250Approval = row.TryGetValue("POST-Deviation Costperpaxabove2250 Approval ", out var POSTDeviationCostperpaxabove2250Approval) ? POSTDeviationCostperpaxabove2250Approval?.ToString() : null,
                        POSTDeviationCostperpaxabove2250ApprovalDate = row.TryGetValue("POST-Deviation Costperpaxabove2250 Approval Date", out var POSTDeviationCostperpaxabove2250ApprovalDate) ? POSTDeviationCostperpaxabove2250ApprovalDate?.ToString() : null,
                        POSTDeviationChangeinvenuetrigger = row.TryGetValue("POST-Deviation Change in venue trigger", out var POSTDeviationChangeinvenuetrigger) ? POSTDeviationChangeinvenuetrigger?.ToString() : null,
                        POSTDeviationChangeinvenueApproval = row.TryGetValue("POST-Deviation Change in venue Approval", out var POSTDeviationChangeinvenueApproval) ? POSTDeviationChangeinvenueApproval?.ToString() : null,
                        POSTDeviationChangeinvenueApprovalDate = row.TryGetValue("POST-Deviation Change in venue Approval Date", out var POSTDeviationChangeinvenueApprovalDate) ? POSTDeviationChangeinvenueApprovalDate?.ToString() : null,
                        POSTDeviationChangeintopictrigger = row.TryGetValue("POST-Deviation Change in topic trigger", out var POSTDeviationChangeintopictrigger) ? POSTDeviationChangeintopictrigger?.ToString() : null,
                        POSTDeviationChangeintopicApproval = row.TryGetValue("POST-Deviation Change in topic Approval", out var POSTDeviationChangeintopicApproval) ? POSTDeviationChangeintopicApproval?.ToString() : null,
                        POSTDeviationChangeintopicApprovalDate = row.TryGetValue("POST-Deviation Change in topic Approval Date", out var POSTDeviationChangeintopicApprovalDate) ? POSTDeviationChangeintopicApprovalDate?.ToString() : null,
                        POSTDeviationChangeinspeakertrigger = row.TryGetValue("POST-Deviation Change in speaker trigger", out var POSTDeviationChangeinspeakertrigger) ? POSTDeviationChangeinspeakertrigger?.ToString() : null,
                        POSTDeviationChangeinspeakerApproval = row.TryGetValue("POST-Deviation Change in speaker Approval", out var POSTDeviationChangeinspeakerApproval) ? POSTDeviationChangeinspeakerApproval?.ToString() : null,
                        POSTDeviationChangeinspeakerApprovalDate = row.TryGetValue("POST-Deviation Change in speaker Approval Date", out var POSTDeviationChangeinspeakerApprovalDate) ? POSTDeviationChangeinspeakerApprovalDate?.ToString() : null,
                        POSTDeviationAttendeesnotcapturedtrigger = row.TryGetValue("POST-Deviation Attendees not captured trigger", out var POSTDeviationAttendeesnotcapturedtrigger) ? POSTDeviationAttendeesnotcapturedtrigger?.ToString() : null,
                        POSTDeviationAttendeesnotcapturedApproval = row.TryGetValue("POST-Deviation Attendees not captured Approval", out var POSTDeviationAttendeesnotcapturedApproval) ? POSTDeviationAttendeesnotcapturedApproval?.ToString() : null,
                        POSTDeviationAttendeesnotcapturedAppDate = row.TryGetValue("POST-Deviation Attendees not captured App Date", out var POSTDeviationAttendeesnotcapturedAppDate) ? POSTDeviationAttendeesnotcapturedAppDate?.ToString() : null,
                        POSTDeviationSpeakernotcapturedtrigger = row.TryGetValue("POST-Deviation Speaker not captured trigger", out var POSTDeviationSpeakernotcapturedtrigger) ? POSTDeviationSpeakernotcapturedtrigger?.ToString() : null,
                        POSTDeviationSpeakernotcapturedApproval = row.TryGetValue("POST-Deviation Speaker not captured  Approval", out var POSTDeviationSpeakernotcapturedApproval) ? POSTDeviationSpeakernotcapturedApproval?.ToString() : null,
                        POSTDeviationSpeakernotcapturedAppDate = row.TryGetValue("POST-Deviation Speaker not captured  App Date", out var POSTDeviationSpeakernotcapturedAppDate) ? POSTDeviationSpeakernotcapturedAppDate?.ToString() : null,
                        POSTDeviationOtherDeviationTrigger = row.TryGetValue("POST-Deviation Other Deviation Trigger", out var POSTDeviationOtherDeviationTrigger) ? POSTDeviationOtherDeviationTrigger?.ToString() : null,
                        OtherDeviationType = row.TryGetValue("Other Deviation Type", out var OtherDeviationType) ? OtherDeviationType?.ToString() : null,
                        POSTDeviationOtherDeviationApproval = row.TryGetValue("POST-Deviation Other Deviation Approval Date", out var POSTDeviationOtherDeviationApproval) ? POSTDeviationOtherDeviationApproval?.ToString() : null,
                        HCPexceeds100000Trigger = row.TryGetValue("HCP exceeds 1,00,000 Trigger", out var HCPexceeds100000Trigger) ? HCPexceeds100000Trigger?.ToString() : null,
                        HCPexceeds100000FHApproval = row.TryGetValue("HCP exceeds 1,00,000 FH Approval", out var HCPexceeds100000FHApproval) ? HCPexceeds100000FHApproval?.ToString() : null,
                        HCPexceeds100000FHApprovalDate = row.TryGetValue("HCP exceeds 1,00,000 FH Approval Date", out var HCPexceeds100000FHApprovalDate) ? HCPexceeds100000FHApprovalDate?.ToString() : null,
                        HCPexceeds500000Trigger = row.TryGetValue("HCP exceeds 5,00,000 Trigger", out var HCPexceeds500000Trigger) ? HCPexceeds500000Trigger?.ToString() : null,
                        HCPexceeds500000TriggerFHApproval = row.TryGetValue("HCP exceeds 5,00,000 Trigger FH Approval", out var HCPexceeds500000TriggerFHApproval) ? HCPexceeds500000TriggerFHApproval?.ToString() : null,
                        HCPexceeds500000TriggerApprovalDate = row.TryGetValue("HCP exceeds 5,00,000 Trigger Approval Date", out var HCPexceeds500000TriggerApprovalDate) ? HCPexceeds500000TriggerApprovalDate?.ToString() : null,
                        HCPHonorarium600000ExceededTrigger = row.TryGetValue("HCP Honorarium 6,00,000 Exceeded Trigger", out var HCPHonorarium600000ExceededTrigger) ? HCPHonorarium600000ExceededTrigger?.ToString() : null,
                        HCPHonorarium600000ExceededApproval = row.TryGetValue("HCP Honorarium 6,00,000 Exceeded Approval", out var HCPHonorarium600000ExceededApproval) ? HCPHonorarium600000ExceededApproval?.ToString() : null,
                        HCPHonorarium600000ExceededApprovalDate = row.TryGetValue("HCP Honorarium 6,00,000 Exceeded Approval Date", out var HCPHonorarium600000ExceededApprovalDate) ? HCPHonorarium600000ExceededApprovalDate?.ToString() : null,
                        TrainerHonorarium1200000ExceededTrigger = row.TryGetValue("Trainer Honorarium 12,00,000 Exceeded Trigger", out var TrainerHonorarium1200000ExceededTrigger) ? TrainerHonorarium1200000ExceededTrigger?.ToString() : null,
                        TrainerHonorarium1200000ExceededApproval = row.TryGetValue("Trainer Honorarium 12,00,000 Exceeded Approval", out var TrainerHonorarium1200000ExceededApproval) ? TrainerHonorarium1200000ExceededApproval?.ToString() : null,
                        TrainerHonorarium1200000ExceededApprovalDate = row.TryGetValue("Trainer Honorarium 12,00,000 Exceeded Approval Dat", out var TrainerHonorarium1200000ExceededApprovalDat) ? TrainerHonorarium1200000ExceededApprovalDat?.ToString() : null,
                        TravelAccomodation300000ExceededTrigger = row.TryGetValue("Travel/Accomodation 3,00,000 Exceeded Trigger", out var TravelAccomodation300000ExceededTrigger) ? TravelAccomodation300000ExceededTrigger?.ToString() : null,
                        TravelAccomodation300000ExceededApproval = row.TryGetValue("Travel/Accomodation 3,00,000 Exceeded Approval", out var TravelAccomodation300000ExceededApproval) ? TravelAccomodation300000ExceededApproval?.ToString() : null,
                        TravelAccomodation300000ExceededApprovalDate = row.TryGetValue("Travel/Accomodation 3,00,000 Exceeded Approval Dat", out var TravelAccomodation300000ExceededApprovalDat) ? TravelAccomodation300000ExceededApprovalDat?.ToString() : null,
                        BrochureRequestLetterUploadedin2daysTrigger = row.TryGetValue("Brochure/Request Letter Uploaded in 2 days Trigger", out var BrochureRequestLetterUploadedin2daysTrigger) ? BrochureRequestLetterUploadedin2daysTrigger?.ToString() : null,
                        BrochureRequestUploadedin2daysFHApproval = row.TryGetValue("Brochure/Request Uploaded in 2 days FH Approval", out var BrochureRequestUploadedin2daysFHApproval) ? BrochureRequestUploadedin2daysFHApproval?.ToString() : null,
                        BrochureRequestin2daysApprovalDate = row.TryGetValue("Brochure/Request in 2 days  Approval Date", out var BrochureRequestin2daysApprovalDate) ? BrochureRequestin2daysApprovalDate?.ToString() : null,
                        CostperparticipantINR2250Trigger = row.TryGetValue("Cost per participant  INR 2250 Trigger ", out var CostperparticipantINR2250Trigger) ? CostperparticipantINR2250Trigger?.ToString() : null,
                        CostperparticipantINR2250FHApproval = row.TryGetValue("Cost per participant INR 2250 FH Approval", out var CostperparticipantINR2250FHApproval) ? CostperparticipantINR2250FHApproval?.ToString() : null,
                        CostperparticipantINR2250ApprovalDate = row.TryGetValue("Cost per participant INR 2250 Approval Date", out var CostperparticipantINR2250ApprovalDate) ? CostperparticipantINR2250ApprovalDate?.ToString() : null,
                        YTDSpendonSpeakerTrigger = row.TryGetValue("YTD Spend on Speaker Trigger", out var YTDSpendonSpeakerTrigger) ? YTDSpendonSpeakerTrigger?.ToString() : null,
                        YTDSpendonSpeakerFHApproval = row.TryGetValue("YTD Spend on Speaker FH Approval", out var YTDSpendonSpeakerFHApproval) ? YTDSpendonSpeakerFHApproval?.ToString() : null,
                        YTDSpendonSpeakerApprovalDate = row.TryGetValue("YTD Spend on Speaker Approval Date", out var YTDSpendonSpeakerApprovalDate) ? YTDSpendonSpeakerApprovalDate?.ToString() : null,
                        amountiswithapplicablelimitTrigger = row.TryGetValue("amount is with applicable limit Trigger", out var amountiswithapplicablelimitTrigger) ? amountiswithapplicablelimitTrigger?.ToString() : null,
                        amountiswithapplicablelimitFHApproval = row.TryGetValue("amount is with applicable limit FH Approval", out var amountiswithapplicablelimitFHApproval) ? amountiswithapplicablelimitFHApproval?.ToString() : null,
                        amountiswithapplicablelimitFHApprovalDate = row.TryGetValue("amount is with applicable limit FH Approval Date", out var amountiswithapplicablelimitFHApprovalDate) ? amountiswithapplicablelimitFHApprovalDate?.ToString() : null,
                        EventOpen45days = row.TryGetValue("EventOpen45days", out var EventOpen45days) ? EventOpen45days?.ToString() : null,
                        EventOpenSalesHeadApproval = row.TryGetValue("EventOpenSalesHeadApproval", out var EventOpenSalesHeadApproval) ? EventOpenSalesHeadApproval?.ToString() : null,
                        EventOpenSalesHeadApprovalDate = row.TryGetValue("EventOpenSalesHeadApproval Date", out var EventOpenSalesHeadApprovalDate) ? EventOpenSalesHeadApprovalDate?.ToString() : null,
                        EventWithin5days = row.TryGetValue("EventWithin5days", out var EventWithin5days) ? EventWithin5days?.ToString() : null,
                        _5daysSalesHeadApproval = row.TryGetValue("5daysSalesHeadApproval", out var _5daysSalesHeadApproval) ? _5daysSalesHeadApproval?.ToString() : null,
                        _5daysSalesHeadApprovaldate = row.TryGetValue("5daysSalesHeadApproval date", out var _5daysSalesHeadApprovaldate) ? _5daysSalesHeadApprovaldate?.ToString() : null,
                        PREFBExpenseExcludingTax = row.TryGetValue("PRE-F&B Expense Excluding Tax", out var PREFBExpenseExcludingTax) ? PREFBExpenseExcludingTax?.ToString() : null,
                        PREFBExpenseExcludingTaxApproval = row.TryGetValue("PRE-F&B Expense Excluding Tax Approval", out var PREFBExpenseExcludingTaxApproval) ? PREFBExpenseExcludingTaxApproval?.ToString() : null,
                        PREFBExpenseExcludingTaxApprovalDate = row.TryGetValue("PRE-F&B Expense Excluding Tax Approval Date", out var PREFBExpenseExcludingTaxApprovalDate) ? PREFBExpenseExcludingTaxApprovalDate?.ToString() : null,
                        HonorariumSubmitted = row.TryGetValue("Honorarium Submitted?", out var prerBmBmApproval) ? prerBmBmApproval?.ToString() : null,
                        PostEventSubmitted = row.TryGetValue("PostEventSubmitted?", out var PostEventSubmitted) ? PostEventSubmitted?.ToString() : null,
                        PostEventBTCExpense = row.TryGetValue("PostEventBTCExpense", out var PostEventBTCExpense) ? PostEventBTCExpense?.ToString() : null,
                        PostEventBTEExpense = row.TryGetValue("PostEventBTEExpense", out var PostEventBTEExpense) ? PostEventBTEExpense?.ToString() : null,
                        ActualAmountGreaterThan50 = row.TryGetValue("ActualAmountGreaterThan50%", out var ActualAmountGreaterThan50) ? ActualAmountGreaterThan50?.ToString() : null,
                        _1UpManager = row.TryGetValue("1 Up Manager", out var _1UpManager) ? _1UpManager?.ToString() : null,
                        ReportingManager = row.TryGetValue("Reporting Manager", out var postEventApprovedDate) ? postEventApprovedDate?.ToString() : null,
                        Brands = row.TryGetValue("Brands", out var brands) ? brands?.ToString() : null,
                        Panelists = row.TryGetValue("Panelists", out var panelists) ? panelists?.ToString() : null,
                        SlideKits = row.TryGetValue("SlideKits", out var slideKits) ? slideKits?.ToString() : null,
                        Invitees = row.TryGetValue("Invitees", out var invitees) ? invitees?.ToString() : null,
                        Expenses = row.TryGetValue("Expenses", out var expenses) ? expenses?.ToString() : null,
                        TotalInvitees = row.TryGetValue("Total Invitees", out var totalInvitees) && int.TryParse(totalInvitees?.ToString(), out var parsedValue) ? parsedValue : 0,
                        TotalAttendees = row.TryGetValue("Total Attendees", out var totalAttendees) && int.TryParse(totalAttendees?.ToString(), out var parsedTotalAttendees) ? parsedTotalAttendees : 0,
                        Finance = row.TryGetValue("Finance", out var finance) ? finance?.ToString() : null,
                        RBMBM = row.TryGetValue("RBM/BM", out var rbmbm) ? rbmbm?.ToString() : null,
                        SalesHeadStatus = row.TryGetValue("Sales Head Status", out var SalesHeadStatus) ? SalesHeadStatus?.ToString() : null,
                        SalesHead = row.TryGetValue("Sales Head", out var salesHead) ? salesHead?.ToString() : null,
                        MarketingHead = row.TryGetValue("Marketing Head", out var marketingHead) ? marketingHead?.ToString() : null,
                        Compliance = row.TryGetValue("Compliance", out var compliance) ? compliance?.ToString() : null,
                        FinanceTreasury = row.TryGetValue("Finance Treasury", out var financeTreasury) ? financeTreasury?.ToString() : null,
                        FinanceAccounts = row.TryGetValue("Finance Accounts", out var financeAccounts) ? financeAccounts?.ToString() : null,
                        MedicalAffairsHead = row.TryGetValue("Medical Affairs Head", out var medicalAffairsHead) ? medicalAffairsHead?.ToString() : null,
                        FinanceHead = row.TryGetValue("Finance Head", out var financeHead) ? financeHead?.ToString() : null,
                        TotalHonorariumAmount = row.TryGetValue("Total Honorarium Amount", out var totalHonorariumAmount) && int.TryParse(totalHonorariumAmount?.ToString(), out var parsedTotalHonorariumAmount) ? parsedTotalHonorariumAmount : 0,
                        TotalTravelAccomodationSpend = row.TryGetValue("Total Travel & Accomodation Spend", out var TotalTravelAccomodationSpend) && int.TryParse(TotalTravelAccomodationSpend?.ToString(), out var parsedTotalTravelAccomodationSpend) ? parsedTotalTravelAccomodationSpend : 0,
                        TotalTravelAmount = row.TryGetValue("Total Travel Amount", out var totalTravelAmount) && int.TryParse(totalTravelAmount?.ToString(), out var parsedTotalTravelAmount) ? parsedTotalTravelAmount : 0,
                        TotalTravelAccommodationAmount = row.TryGetValue("Total Travel & Accommodation Amount", out var totalTravelAccommodationAmount) && int.TryParse(totalTravelAccommodationAmount?.ToString(), out var parsedTotalTravelAccommodationAmount) ? parsedTotalTravelAccommodationAmount : 0,
                        TotalLocalConveyance = row.TryGetValue("Total Local Conveyance", out var totalLocalConveyance) && int.TryParse(totalLocalConveyance?.ToString(), out var parsedTotalLocalConveyance) ? parsedTotalLocalConveyance : 0,
                        TotalExpenseforInvitee = row.TryGetValue("Total Expense for Invitee", out var TotalExpenseforInvitee) ? TotalExpenseforInvitee?.ToString() : null,
                        TotalBudget = row.TryGetValue("Total Budget", out var totalBudget) && int.TryParse(totalBudget?.ToString(), out var parsedTotalBudget) ? parsedTotalBudget : 0,
                        CostperparticipantHelper = row.TryGetValue("Cost per participant - Helper", out var _7daysSalesHeadApprovaldate) ? _7daysSalesHeadApprovaldate?.ToString() : null,
                        POSTDeviationExcludingGST = row.TryGetValue("POST-Deviation Excluding GST?", out var POSTDeviationExcludingGST) ? POSTDeviationExcludingGST?.ToString() : null,
                        FinanceHeadapproval2 = row.TryGetValue("Finance Head approval2", out var FinanceHeadapproval2) ? FinanceHeadapproval2?.ToString() : null,
                        FinanceHeadapproval3 = row.TryGetValue("Finance Head approval3", out var FinanceHeadapproval3) ? FinanceHeadapproval3?.ToString() : null,


                    }).ToList();

                return Ok(eventRequestBrandsList);
            }
            else
            {
                return BadRequest(new
                {
                    Message = "Row not found"
                });
            }
            return Ok();
        }
        [HttpGet("GetDataFromDeviationSheet")]
        public IActionResult GetDataFromDeviationSheet()
        {
            string sheetId = configuration.GetSection("SmartsheetSettings:DeviationProcess").Value;
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
            Dictionary<string, object> ProductBrandsListrowData = new Dictionary<string, object>();
            //Row ExistingRow = sheet.Rows.FirstOrDefault(row => row.)
            //Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == eventId));
            //if (targetRow != null)
            //{
            //    long Id = targetRow.Id.Value;
            List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);

            List<DeviationDataPayload> eventRequestBrandsList = sheetData
                //.Where(row => row.TryGetValue("EventId/EventRequestId", out var eventIdValue) && eventIdValue?.ToString() == eventId)
                .Select(row => new DeviationDataPayload
                {
                    DeviationID = row.TryGetValue("Deviation ID", out var DeviationID) ? DeviationID?.ToString() : null,
                    EventIdEventRequestId = row.TryGetValue("EventId/EventRequestId", out var eventId) ? eventId?.ToString() : null,
                    EventTopic = row.TryGetValue("Event Topic", out var eventTopic) ? eventTopic?.ToString() : null,
                    MISCode = row.TryGetValue("MIS Code", out var MISCode) ? MISCode?.ToString() : null,
                    HCPName = row.TryGetValue("HCP Name", out var HCPName) ? HCPName?.ToString() : null,
                    HonorariumAmount = row.TryGetValue("Honorarium Amount", out var HonorariumAmount) ? HonorariumAmount?.ToString() : null,
                    TravelAccommodationAmount = row.TryGetValue("Travel & Accommodation Amount", out var TravelAccommodationAmount) ? TravelAccommodationAmount?.ToString() : null,
                    OtherExpenses = row.TryGetValue("Other Expenses", out var OtherExpenses) ? OtherExpenses?.ToString() : null,
                    DeviationType = row.TryGetValue("Deviation Type", out var DeviationType) ? DeviationType?.ToString() : null,
                    EventLevel = row.TryGetValue("Event Level", out var EventLevel) ? EventLevel?.ToString() : null,
                    OutstandingEvents = row.TryGetValue("Outstanding Events", out var OutstandingEvents) ? OutstandingEvents?.ToString() : null,
                    InitiatorName = row.TryGetValue("InitiatorName", out var InitiatorName) ? InitiatorName?.ToString() : null,
                    InitiatorEmail = row.TryGetValue("Initiator Email", out var initiatorEmail) ? initiatorEmail?.ToString() : null,
                    CreatedOn = row.TryGetValue("CreatedOn", out var premarketingHeadApproval) ? premarketingHeadApproval?.ToString() : null,
                    Modified = row.TryGetValue("Modified", out var premarketingHeadApprovalDate) ? premarketingHeadApprovalDate?.ToString() : null,
                    EventType = row.TryGetValue("EventType", out var eventType) ? eventType?.ToString() : null,
                    EventDate = row.TryGetValue("EventDate", out var eventDate) ? eventDate?.ToString() : null,
                    StartTime = row.TryGetValue("StartTime", out var startTime) ? startTime?.ToString() : null,
                    EndTime = row.TryGetValue("EndTime", out var endTime) ? endTime?.ToString() : null,
                    EventEndDate = row.TryGetValue("Event End Date", out var eventEndDate) ? eventEndDate?.ToString() : null,
                    VenueName = row.TryGetValue("VenueName", out var venueName) ? venueName?.ToString() : null,
                    State = row.TryGetValue("State", out var state) ? state?.ToString() : null,
                    City = row.TryGetValue("City", out var city) ? city?.ToString() : null,
                    IsAdvanceRequired = row.TryGetValue("IsAdvanceRequired", out var isAdvanceRequired) ? isAdvanceRequired?.ToString() : null,
                    SalesHeadapproval = row.TryGetValue("Sales Head approval", out var eventOpenSalesHeadApproval) ? eventOpenSalesHeadApproval?.ToString() : null,
                    SalesHeadapprovalDate = row.TryGetValue("Sales Head approval Date", out var SalesHeadapprovalDate) ? SalesHeadapprovalDate?.ToString() : null,
                    FinanceHeadApproval = row.TryGetValue("Finance Head Approval", out var FinanceHeadApproval) ? FinanceHeadApproval?.ToString() : null,
                    FinanceHeadApprovalDate = row.TryGetValue("Finance Head Approval Date", out var FinanceHeadApprovalDate) ? FinanceHeadApprovalDate?.ToString() : null,
                    ApprovalStatus = row.TryGetValue("Approval Status", out var eventOpenSalesHeadApprovalDate) ? eventOpenSalesHeadApprovalDate?.ToString() : null,
                    HON5WorkingdaysDeviationDateTrigger = row.TryGetValue("HON-5Workingdays Deviation Date Trigger", out var HON5WorkingdaysDeviationDateTrigger) ? HON5WorkingdaysDeviationDateTrigger?.ToString() : null,
                    HON5WorkingdaysDeviationApproval = row.TryGetValue("HON-5Workingdays Deviation Approval", out var HON5WorkingdaysDeviationApproval) ? HON5WorkingdaysDeviationApproval?.ToString() : null,
                    HON5WorkingdaysDeviationApprovalDate = row.TryGetValue("HON-5Workingdays Deviation Approval Date", out var HON5WorkingdaysDeviationApprovalDate) ? HON5WorkingdaysDeviationApprovalDate?.ToString() : null,
                    POSTBeyond45DaysDeviationDateTrigger = row.TryGetValue("POST- Beyond45Days Deviation Date Trigger", out var POSTBeyond45DaysDeviationDateTrigger) ? POSTBeyond45DaysDeviationDateTrigger?.ToString() : null,
                    POSTBeyond45DaysDeviationApproval = row.TryGetValue("POST- Beyond45Days Deviation Approval", out var POSTBeyond45DaysDeviationApproval) ? POSTBeyond45DaysDeviationApproval?.ToString() : null,
                    POSTBeyond45DaysDeviationApprovalDate = row.TryGetValue("POST- Beyond45Days Deviation Approval Date", out var POSTBeyond45DaysDeviationApprovalDate) ? POSTBeyond45DaysDeviationApprovalDate?.ToString() : null,
                    POSTLessthan5InviteesDeviationTrigger = row.TryGetValue("POST-Lessthan5Invitees Deviation Trigger", out var POSTLessthan5InviteesDeviationTrigger) ? POSTLessthan5InviteesDeviationTrigger?.ToString() : null,
                    POSTLessthan5inviteesDeviationApproval = row.TryGetValue("POST-Lessthan5invitees Deviation Approval", out var POSTLessthan5inviteesDeviationApproval) ? POSTLessthan5inviteesDeviationApproval?.ToString() : null,
                    POSTLessthan5inviteesDeviationApprovalDate = row.TryGetValue("POST-Lessthan5invitees Deviation Approval Date", out var POSTLessthan5inviteesDeviationApprovalDate) ? POSTLessthan5inviteesDeviationApprovalDate?.ToString() : null,
                    POSTDeviationCostperpaxabove1500Trigger = row.TryGetValue("POST-Deviation Costperpaxabove1500 Trigger", out var POSTDeviationCostperpaxabove1500Trigger) ? POSTDeviationCostperpaxabove1500Trigger?.ToString() : null,
                    POSTDeviationCostperpaxabove1500Approval = row.TryGetValue("POST-Deviation Costperpaxabove1500 Approval ", out var POSTDeviationCostperpaxabove1500Approval) ? POSTDeviationCostperpaxabove1500Approval?.ToString() : null,
                    POSTDeviationCostperpaxabove1500ApprovalDate = row.TryGetValue("POST-Deviation Costperpaxabove1500 Approval Date", out var POSTDeviationCostperpaxabove1500ApprovalDate) ? POSTDeviationCostperpaxabove1500ApprovalDate?.ToString() : null,
                    POSTDeviationCostperpaxabove2250Trigger = row.TryGetValue("POST-Deviation Costperpaxabove2250 Trigger", out var POSTDeviationCostperpaxabove2250Trigger) ? POSTDeviationCostperpaxabove2250Trigger?.ToString() : null,
                    POSTDeviationCostperpaxabove2250Approval = row.TryGetValue("POST-Deviation Costperpaxabove2250 Approval ", out var POSTDeviationCostperpaxabove2250Approval) ? POSTDeviationCostperpaxabove2250Approval?.ToString() : null,
                    POSTDeviationCostperpaxabove2250ApprovalDate = row.TryGetValue("POST-Deviation Costperpaxabove2250 Approval Date", out var POSTDeviationCostperpaxabove2250ApprovalDate) ? POSTDeviationCostperpaxabove2250ApprovalDate?.ToString() : null,
                    POSTDeviationChangeinvenuetrigger = row.TryGetValue("POST-Deviation Change in venue trigger", out var POSTDeviationChangeinvenuetrigger) ? POSTDeviationChangeinvenuetrigger?.ToString() : null,
                    POSTDeviationChangeinvenueApproval = row.TryGetValue("POST-Deviation Change in venue Approval", out var POSTDeviationChangeinvenueApproval) ? POSTDeviationChangeinvenueApproval?.ToString() : null,
                    POSTDeviationChangeinvenueApprovalDate = row.TryGetValue("POST-Deviation Change in venue Approval Date", out var POSTDeviationChangeinvenueApprovalDate) ? POSTDeviationChangeinvenueApprovalDate?.ToString() : null,
                    POSTDeviationChangeintopictrigger = row.TryGetValue("POST-Deviation Change in topic trigger", out var POSTDeviationChangeintopictrigger) ? POSTDeviationChangeintopictrigger?.ToString() : null,
                    POSTDeviationChangeintopicApproval = row.TryGetValue("POST-Deviation Change in topic Approval", out var POSTDeviationChangeintopicApproval) ? POSTDeviationChangeintopicApproval?.ToString() : null,
                    POSTDeviationChangeintopicApprovalDate = row.TryGetValue("POST-Deviation Change in topic Approval Date", out var POSTDeviationChangeintopicApprovalDate) ? POSTDeviationChangeintopicApprovalDate?.ToString() : null,
                    POSTDeviationChangeinspeakertrigger = row.TryGetValue("POST-Deviation Change in speaker trigger", out var POSTDeviationChangeinspeakertrigger) ? POSTDeviationChangeinspeakertrigger?.ToString() : null,
                    POSTDeviationChangeinspeakerApproval = row.TryGetValue("POST-Deviation Change in speaker Approval", out var POSTDeviationChangeinspeakerApproval) ? POSTDeviationChangeinspeakerApproval?.ToString() : null,
                    POSTDeviationChangeinspeakerApprovalDate = row.TryGetValue("POST-Deviation Change in speaker Approval Date", out var POSTDeviationChangeinspeakerApprovalDate) ? POSTDeviationChangeinspeakerApprovalDate?.ToString() : null,
                    POSTDeviationAttendeesnotcapturedtrigger = row.TryGetValue("POST-Deviation Attendees not captured trigger", out var POSTDeviationAttendeesnotcapturedtrigger) ? POSTDeviationAttendeesnotcapturedtrigger?.ToString() : null,
                    POSTDeviationAttendeesnotcapturedApproval = row.TryGetValue("POST-Deviation Attendees not captured Approval", out var POSTDeviationAttendeesnotcapturedApproval) ? POSTDeviationAttendeesnotcapturedApproval?.ToString() : null,
                    POSTDeviationAttendeesnotcapturedAppDate = row.TryGetValue("POST-Deviation Attendees not captured App Date", out var POSTDeviationAttendeesnotcapturedAppDate) ? POSTDeviationAttendeesnotcapturedAppDate?.ToString() : null,
                    POSTDeviationSpeakernotcapturedtrigger = row.TryGetValue("POST-Deviation Speaker not captured trigger", out var POSTDeviationSpeakernotcapturedtrigger) ? POSTDeviationSpeakernotcapturedtrigger?.ToString() : null,
                    POSTDeviationSpeakernotcapturedApproval = row.TryGetValue("POST-Deviation Speaker not captured  Approval", out var POSTDeviationSpeakernotcapturedApproval) ? POSTDeviationSpeakernotcapturedApproval?.ToString() : null,
                    POSTDeviationSpeakernotcapturedAppDate = row.TryGetValue("POST-Deviation Speaker not captured  App Date", out var POSTDeviationSpeakernotcapturedAppDate) ? POSTDeviationSpeakernotcapturedAppDate?.ToString() : null,
                    POSTDeviationOtherDeviationTrigger = row.TryGetValue("POST-Deviation Other Deviation Trigger", out var POSTDeviationOtherDeviationTrigger) ? POSTDeviationOtherDeviationTrigger?.ToString() : null,
                    OtherDeviationType = row.TryGetValue("Other Deviation Type", out var OtherDeviationType) ? OtherDeviationType?.ToString() : null,
                    POSTDeviationOtherDeviationApproval = row.TryGetValue("POST-Deviation Other Deviation Approval Date", out var POSTDeviationOtherDeviationApproval) ? POSTDeviationOtherDeviationApproval?.ToString() : null,
                    HCPexceeds100000Trigger = row.TryGetValue("HCP exceeds 1,00,000 Trigger", out var HCPexceeds100000Trigger) ? HCPexceeds100000Trigger?.ToString() : null,
                    HCPexceeds100000FHApproval = row.TryGetValue("HCP exceeds 1,00,000 FH Approval", out var HCPexceeds100000FHApproval) ? HCPexceeds100000FHApproval?.ToString() : null,
                    HCPexceeds100000FHApprovalDate = row.TryGetValue("HCP exceeds 1,00,000 FH Approval Date", out var HCPexceeds100000FHApprovalDate) ? HCPexceeds100000FHApprovalDate?.ToString() : null,
                    HCPexceeds500000Trigger = row.TryGetValue("HCP exceeds 5,00,000 Trigger", out var HCPexceeds500000Trigger) ? HCPexceeds500000Trigger?.ToString() : null,
                    HCPexceeds500000TriggerFHApproval = row.TryGetValue("HCP exceeds 5,00,000 Trigger FH Approval", out var HCPexceeds500000TriggerFHApproval) ? HCPexceeds500000TriggerFHApproval?.ToString() : null,
                    HCPexceeds500000TriggerApprovalDate = row.TryGetValue("HCP exceeds 5,00,000 Trigger Approval Date", out var HCPexceeds500000TriggerApprovalDate) ? HCPexceeds500000TriggerApprovalDate?.ToString() : null,
                    HCPHonorarium600000ExceededTrigger = row.TryGetValue("HCP Honorarium 6,00,000 Exceeded Trigger", out var HCPHonorarium600000ExceededTrigger) ? HCPHonorarium600000ExceededTrigger?.ToString() : null,
                    HCPHonorarium600000ExceededApproval = row.TryGetValue("HCP Honorarium 6,00,000 Exceeded Approval", out var HCPHonorarium600000ExceededApproval) ? HCPHonorarium600000ExceededApproval?.ToString() : null,
                    HCPHonorarium600000ExceededApprovalDate = row.TryGetValue("HCP Honorarium 6,00,000 Exceeded Approval Date", out var HCPHonorarium600000ExceededApprovalDate) ? HCPHonorarium600000ExceededApprovalDate?.ToString() : null,
                    TrainerHonorarium1200000ExceededTrigger = row.TryGetValue("Trainer Honorarium 12,00,000 Exceeded Trigger", out var TrainerHonorarium1200000ExceededTrigger) ? TrainerHonorarium1200000ExceededTrigger?.ToString() : null,
                    TrainerHonorarium1200000ExceededApproval = row.TryGetValue("Trainer Honorarium 12,00,000 Exceeded Approval", out var TrainerHonorarium1200000ExceededApproval) ? TrainerHonorarium1200000ExceededApproval?.ToString() : null,
                    TrainerHonorarium1200000ExceededApprovalDate = row.TryGetValue("Trainer Honorarium 12,00,000 Exceeded Approval Dat", out var TrainerHonorarium1200000ExceededApprovalDat) ? TrainerHonorarium1200000ExceededApprovalDat?.ToString() : null,
                    TravelAccomodation300000ExceededTrigger = row.TryGetValue("Travel/Accomodation 3,00,000 Exceeded Trigger", out var TravelAccomodation300000ExceededTrigger) ? TravelAccomodation300000ExceededTrigger?.ToString() : null,
                    TravelAccomodation300000ExceededApproval = row.TryGetValue("Travel/Accomodation 3,00,000 Exceeded Approval", out var TravelAccomodation300000ExceededApproval) ? TravelAccomodation300000ExceededApproval?.ToString() : null,
                    TravelAccomodation300000ExceededApprovalDate = row.TryGetValue("Travel/Accomodation 3,00,000 Exceeded Approval Dat", out var TravelAccomodation300000ExceededApprovalDat) ? TravelAccomodation300000ExceededApprovalDat?.ToString() : null,
                    BrochureRequestLetterUploadedin2daysTrigger = row.TryGetValue("Brochure/Request Letter Uploaded in 2 days Trigger", out var BrochureRequestLetterUploadedin2daysTrigger) ? BrochureRequestLetterUploadedin2daysTrigger?.ToString() : null,
                    BrochureRequestUploadedin2daysFHApproval = row.TryGetValue("Brochure/Request Uploaded in 2 days FH Approval", out var BrochureRequestUploadedin2daysFHApproval) ? BrochureRequestUploadedin2daysFHApproval?.ToString() : null,
                    BrochureRequestin2daysApprovalDate = row.TryGetValue("Brochure/Request in 2 days  Approval Date", out var BrochureRequestin2daysApprovalDate) ? BrochureRequestin2daysApprovalDate?.ToString() : null,
                    CostperparticipantINR2250Trigger = row.TryGetValue("Cost per participant  INR 2250 Trigger ", out var CostperparticipantINR2250Trigger) ? CostperparticipantINR2250Trigger?.ToString() : null,
                    CostperparticipantINR2250FHApproval = row.TryGetValue("Cost per participant INR 2250 FH Approval", out var CostperparticipantINR2250FHApproval) ? CostperparticipantINR2250FHApproval?.ToString() : null,
                    CostperparticipantINR2250ApprovalDate = row.TryGetValue("Cost per participant INR 2250 Approval Date", out var CostperparticipantINR2250ApprovalDate) ? CostperparticipantINR2250ApprovalDate?.ToString() : null,
                    YTDSpendonSpeakerTrigger = row.TryGetValue("YTD Spend on Speaker Trigger", out var YTDSpendonSpeakerTrigger) ? YTDSpendonSpeakerTrigger?.ToString() : null,
                    YTDSpendonSpeakerFHApproval = row.TryGetValue("YTD Spend on Speaker FH Approval", out var YTDSpendonSpeakerFHApproval) ? YTDSpendonSpeakerFHApproval?.ToString() : null,
                    YTDSpendonSpeakerApprovalDate = row.TryGetValue("YTD Spend on Speaker Approval Date", out var YTDSpendonSpeakerApprovalDate) ? YTDSpendonSpeakerApprovalDate?.ToString() : null,
                    amountiswithapplicablelimitTrigger = row.TryGetValue("amount is with applicable limit Trigger", out var amountiswithapplicablelimitTrigger) ? amountiswithapplicablelimitTrigger?.ToString() : null,
                    amountiswithapplicablelimitFHApproval = row.TryGetValue("amount is with applicable limit FH Approval", out var amountiswithapplicablelimitFHApproval) ? amountiswithapplicablelimitFHApproval?.ToString() : null,
                    amountiswithapplicablelimitFHApprovalDate = row.TryGetValue("amount is with applicable limit FH Approval Date", out var amountiswithapplicablelimitFHApprovalDate) ? amountiswithapplicablelimitFHApprovalDate?.ToString() : null,
                    EventOpen45days = row.TryGetValue("EventOpen45days", out var EventOpen45days) ? EventOpen45days?.ToString() : null,
                    EventOpenSalesHeadApproval = row.TryGetValue("EventOpenSalesHeadApproval", out var EventOpenSalesHeadApproval) ? EventOpenSalesHeadApproval?.ToString() : null,
                    EventOpenSalesHeadApprovalDate = row.TryGetValue("EventOpenSalesHeadApproval Date", out var EventOpenSalesHeadApprovalDate) ? EventOpenSalesHeadApprovalDate?.ToString() : null,
                    EventWithin5days = row.TryGetValue("EventWithin5days", out var EventWithin5days) ? EventWithin5days?.ToString() : null,
                    _5daysSalesHeadApproval = row.TryGetValue("5daysSalesHeadApproval", out var _5daysSalesHeadApproval) ? _5daysSalesHeadApproval?.ToString() : null,
                    _5daysSalesHeadApprovaldate = row.TryGetValue("5daysSalesHeadApproval date", out var _5daysSalesHeadApprovaldate) ? _5daysSalesHeadApprovaldate?.ToString() : null,
                    PREFBExpenseExcludingTax = row.TryGetValue("PRE-F&B Expense Excluding Tax", out var PREFBExpenseExcludingTax) ? PREFBExpenseExcludingTax?.ToString() : null,
                    PREFBExpenseExcludingTaxApproval = row.TryGetValue("PRE-F&B Expense Excluding Tax Approval", out var PREFBExpenseExcludingTaxApproval) ? PREFBExpenseExcludingTaxApproval?.ToString() : null,
                    PREFBExpenseExcludingTaxApprovalDate = row.TryGetValue("PRE-F&B Expense Excluding Tax Approval Date", out var PREFBExpenseExcludingTaxApprovalDate) ? PREFBExpenseExcludingTaxApprovalDate?.ToString() : null,
                    HonorariumSubmitted = row.TryGetValue("Honorarium Submitted?", out var prerBmBmApproval) ? prerBmBmApproval?.ToString() : null,
                    PostEventSubmitted = row.TryGetValue("PostEventSubmitted?", out var PostEventSubmitted) ? PostEventSubmitted?.ToString() : null,
                    PostEventBTCExpense = row.TryGetValue("PostEventBTCExpense", out var PostEventBTCExpense) ? PostEventBTCExpense?.ToString() : null,
                    PostEventBTEExpense = row.TryGetValue("PostEventBTEExpense", out var PostEventBTEExpense) ? PostEventBTEExpense?.ToString() : null,
                    ActualAmountGreaterThan50 = row.TryGetValue("ActualAmountGreaterThan50%", out var ActualAmountGreaterThan50) ? ActualAmountGreaterThan50?.ToString() : null,
                    _1UpManager = row.TryGetValue("1 Up Manager", out var _1UpManager) ? _1UpManager?.ToString() : null,
                    ReportingManager = row.TryGetValue("Reporting Manager", out var postEventApprovedDate) ? postEventApprovedDate?.ToString() : null,
                    Brands = row.TryGetValue("Brands", out var brands) ? brands?.ToString() : null,
                    Panelists = row.TryGetValue("Panelists", out var panelists) ? panelists?.ToString() : null,
                    SlideKits = row.TryGetValue("SlideKits", out var slideKits) ? slideKits?.ToString() : null,
                    Invitees = row.TryGetValue("Invitees", out var invitees) ? invitees?.ToString() : null,
                    Expenses = row.TryGetValue("Expenses", out var expenses) ? expenses?.ToString() : null,
                    TotalInvitees = row.TryGetValue("Total Invitees", out var totalInvitees) && int.TryParse(totalInvitees?.ToString(), out var parsedValue) ? parsedValue : 0,
                    TotalAttendees = row.TryGetValue("Total Attendees", out var totalAttendees) && int.TryParse(totalAttendees?.ToString(), out var parsedTotalAttendees) ? parsedTotalAttendees : 0,
                    Finance = row.TryGetValue("Finance", out var finance) ? finance?.ToString() : null,
                    RBMBM = row.TryGetValue("RBM/BM", out var rbmbm) ? rbmbm?.ToString() : null,
                    SalesHeadStatus = row.TryGetValue("Sales Head Status", out var SalesHeadStatus) ? SalesHeadStatus?.ToString() : null,
                    SalesHead = row.TryGetValue("Sales Head", out var salesHead) ? salesHead?.ToString() : null,
                    MarketingHead = row.TryGetValue("Marketing Head", out var marketingHead) ? marketingHead?.ToString() : null,
                    Compliance = row.TryGetValue("Compliance", out var compliance) ? compliance?.ToString() : null,
                    FinanceTreasury = row.TryGetValue("Finance Treasury", out var financeTreasury) ? financeTreasury?.ToString() : null,
                    FinanceAccounts = row.TryGetValue("Finance Accounts", out var financeAccounts) ? financeAccounts?.ToString() : null,
                    MedicalAffairsHead = row.TryGetValue("Medical Affairs Head", out var medicalAffairsHead) ? medicalAffairsHead?.ToString() : null,
                    FinanceHead = row.TryGetValue("Finance Head", out var financeHead) ? financeHead?.ToString() : null,
                    TotalHonorariumAmount = row.TryGetValue("Total Honorarium Amount", out var totalHonorariumAmount) && int.TryParse(totalHonorariumAmount?.ToString(), out var parsedTotalHonorariumAmount) ? parsedTotalHonorariumAmount : 0,
                    TotalTravelAccomodationSpend = row.TryGetValue("Total Travel & Accomodation Spend", out var TotalTravelAccomodationSpend) && int.TryParse(TotalTravelAccomodationSpend?.ToString(), out var parsedTotalTravelAccomodationSpend) ? parsedTotalTravelAccomodationSpend : 0,
                    TotalTravelAmount = row.TryGetValue("Total Travel Amount", out var totalTravelAmount) && int.TryParse(totalTravelAmount?.ToString(), out var parsedTotalTravelAmount) ? parsedTotalTravelAmount : 0,
                    TotalTravelAccommodationAmount = row.TryGetValue("Total Travel & Accommodation Amount", out var totalTravelAccommodationAmount) && int.TryParse(totalTravelAccommodationAmount?.ToString(), out var parsedTotalTravelAccommodationAmount) ? parsedTotalTravelAccommodationAmount : 0,
                    TotalLocalConveyance = row.TryGetValue("Total Local Conveyance", out var totalLocalConveyance) && int.TryParse(totalLocalConveyance?.ToString(), out var parsedTotalLocalConveyance) ? parsedTotalLocalConveyance : 0,
                    TotalExpenseforInvitee = row.TryGetValue("Total Expense for Invitee", out var TotalExpenseforInvitee) ? TotalExpenseforInvitee?.ToString() : null,
                    TotalBudget = row.TryGetValue("Total Budget", out var totalBudget) && int.TryParse(totalBudget?.ToString(), out var parsedTotalBudget) ? parsedTotalBudget : 0,
                    CostperparticipantHelper = row.TryGetValue("Cost per participant - Helper", out var _7daysSalesHeadApprovaldate) ? _7daysSalesHeadApprovaldate?.ToString() : null,
                    POSTDeviationExcludingGST = row.TryGetValue("POST-Deviation Excluding GST?", out var POSTDeviationExcludingGST) ? POSTDeviationExcludingGST?.ToString() : null,
                    FinanceHeadapproval2 = row.TryGetValue("Finance Head approval2", out var FinanceHeadapproval2) ? FinanceHeadapproval2?.ToString() : null,
                    FinanceHeadapproval3 = row.TryGetValue("Finance Head approval3", out var FinanceHeadapproval3) ? FinanceHeadapproval3?.ToString() : null,


                }).ToList();

            return Ok(eventRequestBrandsList);
            //}
            //else
            //{
            //    return BadRequest(new
            //    {
            //        Message = "Row not found"
            //    });
            //}
            //return Ok();
        }

        [HttpGet("GetEventSettlementDataUsingEventId")]
        public IActionResult GetEventSettlementDataUsingEventId(string eventId)
        {
            string sheetId = configuration.GetSection("SmartsheetSettings:EventSettlement").Value;
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
            Dictionary<string, object> ProductBrandsListrowData = new Dictionary<string, object>();
            //Row ExistingRow = sheet.Rows.FirstOrDefault(row => row.)
            Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == eventId));
            if (targetRow != null)
            {
                long Id = targetRow.Id.Value;
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);

                List<EventSettlementPayload> eventRequestBrandsList = sheetData
                    .Where(row => row.TryGetValue("EventId/EventRequestId", out var eventIdValue) && eventIdValue?.ToString() == eventId)
                    .Select(row => new EventSettlementPayload
                    {
                        EventType = row.TryGetValue("EventType", out var eventType) ? eventType?.ToString() : null,
                        EventIdEventRequestId = row.TryGetValue("EventId/EventRequestId", out var eventId) ? eventId?.ToString() : null,
                        EventTopic = row.TryGetValue("Event Topic", out var eventTopic) ? eventTopic?.ToString() : null,
                        EventDate = row.TryGetValue("EventDate", out var eventDate) ? eventDate?.ToString() : null,
                        EventEndDate = row.TryGetValue("Event End Date", out var eventEndDate) ? eventEndDate?.ToString() : null,
                        StartTime = row.TryGetValue("StartTime", out var startTime) ? startTime?.ToString() : null,
                        EndTime = row.TryGetValue("EndTime", out var endTime) ? endTime?.ToString() : null,
                        VenueName = row.TryGetValue("VenueName", out var venueName) ? venueName?.ToString() : null,
                        State = row.TryGetValue("State", out var state) ? state?.ToString() : null,
                        City = row.TryGetValue("City", out var city) ? city?.ToString() : null,
                        Attended = row.TryGetValue("Attended", out var totalInvitees) && double.TryParse(totalInvitees?.ToString(), out var parsedValue) ? parsedValue : 0,
                        InviteesParticipated = row.TryGetValue("InviteesParticipated", out var honorariumApprovedDate) ? honorariumApprovedDate?.ToString() : null,
                        ExpenseDetails = row.TryGetValue("ExpenseDetails", out var ExpenseDetails) ? ExpenseDetails?.ToString() : null,
                        TotalExpenseDetails = row.TryGetValue("TotalExpenseDetails", out var TotalExpenseDetails) ? TotalExpenseDetails?.ToString() : null,
                        AdvanceDetails = row.TryGetValue("AdvanceDetails", out var AdvanceDetails) ? AdvanceDetails?.ToString() : null,
                        AdvanceUtilizedForEvent = row.TryGetValue("Advance Utilized For Event", out var AdvanceUtilizedForEvent) ? AdvanceUtilizedForEvent?.ToString() : null,
                        PayBackAmountToCompany = row.TryGetValue("Pay Back Amount To Company", out var PayBackAmountToCompany) ? PayBackAmountToCompany?.ToString() : null,
                        AdditionalAmountNeededToPayForInitiator = row.TryGetValue("Additional Amount Needed To Pay For Initiator", out var AdditionalAmountNeededToPayForInitiator) ? AdditionalAmountNeededToPayForInitiator?.ToString() : null,
                        CreatedOn = row.TryGetValue("CreatedOn", out var premarketingHeadApproval) ? premarketingHeadApproval?.ToString() : null,
                        CreatedDateHelper = row.TryGetValue("Created Date - Helper", out var createdDateHelper) ? createdDateHelper?.ToString() : null,
                        Modified = row.TryGetValue("Modified", out var premarketingHeadApprovalDate) ? premarketingHeadApprovalDate?.ToString() : null,
                        InitiatorName = row.TryGetValue("InitiatorName", out var initiatorName) ? initiatorName?.ToString() : null,
                        InitiatorEmail = row.TryGetValue("Initiator Email", out var initiatorEmail) ? initiatorEmail?.ToString() : null,
                        IsAdvanceRequired = row.TryGetValue("IsAdvanceRequired", out var isAdvanceRequired) ? isAdvanceRequired?.ToString() : null,
                        EventRequestStatus = row.TryGetValue("Event Request Status", out var prerBmBmApprovalDate) ? prerBmBmApprovalDate?.ToString() : null,
                        HonorariumRequestStatus = row.TryGetValue("Honorarium Request Status", out var honorariumRequestStatus) ? honorariumRequestStatus?.ToString() : null,
                        PostEventRequeststatus = row.TryGetValue("Post Event Request status", out var PostEventRequeststatus) ? PostEventRequeststatus?.ToString() : null,
                        PostEventApprovedDate = row.TryGetValue("Post Event Approved Date", out var PostEventApprovedDate) ? PostEventApprovedDate?.ToString() : null,
                        EventSettlementRBMBMApproval = row.TryGetValue("EventSettlement-RBM/BM Approval", out var agreement) ? agreement?.ToString() : null,
                        EventSettlementRBMBMApprovalDate = row.TryGetValue("EventSettlement-RBM/BM Approval Date", out var HONRBMBMApprovalDate) ? HONRBMBMApprovalDate?.ToString() : null,
                        EventSettlementSalesHeadApproval = row.TryGetValue("EventSettlement-SalesHead Approval", out var HONSalesHeadApproval) ? HONSalesHeadApproval?.ToString() : null,
                        EventSettlementSalesHeadApprovalDate = row.TryGetValue("EventSettlementSalesHeadApprovalDate", out var HONSalesHeadApprovalDate) ? HONSalesHeadApprovalDate?.ToString() : null,
                        EventSettlementComplianceApproval = row.TryGetValue("EventSettlement-Compliance Approval", out var EventSettlementComplianceApproval) ? EventSettlementComplianceApproval?.ToString() : null,
                        EventSettlementComplianceApprovalDate = row.TryGetValue("EventSettlement-Compliance Approval Date", out var EventSettlementComplianceApprovalDate) ? EventSettlementComplianceApprovalDate?.ToString() : null,
                        EventSettlementFinanceAccountApproval = row.TryGetValue("EventSettlement-Finance Account Approval", out var EventSettlementFinanceAccountApproval) ? EventSettlementFinanceAccountApproval?.ToString() : null,
                        EventSettlementFinanceAccountComments = row.TryGetValue("EventSettlement-Finance Account Comments", out var EventSettlementFinanceAccountComments) ? EventSettlementFinanceAccountComments?.ToString() : null,
                        EventSettlementFinanceAccountApprovalDate = row.TryGetValue("EventSettlement-Finance Account Approval Date", out var EventSettlementFinanceAccountApprovalDate) ? EventSettlementFinanceAccountApprovalDate?.ToString() : null,
                        EventSettlementFinanceTreasuryApproval = row.TryGetValue("EventSettlement-Finance Treasury Approval Date", out var EventSettlementFinanceTreasuryApproval) ? EventSettlementFinanceTreasuryApproval?.ToString() : null,
                        EventSettlementMarketingHeadApproval = row.TryGetValue("EventSettlement-Marketing Head Approval Date", out var EventSettlementMarketingHeadApproval) ? EventSettlementMarketingHeadApproval?.ToString() : null,
                        EventSettlementMedicalAffairsHeadApproval = row.TryGetValue("EventSettlement-Medical Affairs Head Approval", out var EventSettlementMedicalAffairsHeadApproval) ? EventSettlementMedicalAffairsHeadApproval?.ToString() : null,
                        EventSettlementMedicalAffairsHeadApprovalDate = row.TryGetValue("EventSettlement-Medical Affairs Head Approval Date", out var EventSettlementMedicalAffairsHeadApprovalDate) ? EventSettlementMedicalAffairsHeadApprovalDate?.ToString() : null,
                        DeviationStatus = row.TryGetValue("Deviation Status", out var DeviationStatus) ? DeviationStatus?.ToString() : null,
                        IsAllDeviationsApproved = row.TryGetValue("Is All Deviations Approved?", out var IsAllDeviationsApproved) ? IsAllDeviationsApproved?.ToString() : null,
                        POSTBeyond30DaysDeviationApproval = row.TryGetValue("POST- Beyond30Days Deviation Approval", out var POSTBeyond30DaysDeviationApproval) ? POSTBeyond30DaysDeviationApproval?.ToString() : null,
                        post45daysapproved = row.TryGetValue("post 45 days approved", out var post45daysapproved) ? post45daysapproved?.ToString() : null,
                        POSTBeyond30DaysDeviationApprovalDate = row.TryGetValue("POST- Beyond30Days Deviation Approval Date", out var POSTBeyond30DaysDeviationApprovalDate) ? POSTBeyond30DaysDeviationApprovalDate?.ToString() : null,
                        POSTLessThan5InviteesDeviationApproval = row.TryGetValue("POST-LessThan5Invitees Deviation Approval", out var POSTLessThan5InviteesDeviationApproval) ? POSTLessThan5InviteesDeviationApproval?.ToString() : null,
                        Post5InviteesApproved = row.TryGetValue("Post <5 Invitees Approved", out var Post5InviteesApproved) ? Post5InviteesApproved?.ToString() : null,
                        POSTDeviationCostperpaxabove1500Approval = row.TryGetValue("POST-Deviation Costperpaxabove1500 Approval ", out var POSTDeviationCostperpaxabove1500Approval) ? POSTDeviationCostperpaxabove1500Approval?.ToString() : null,
                        PostCostperPaxApproved = row.TryGetValue("Post CostperPax Approved", out var PostCostperPaxApproved) ? PostCostperPaxApproved?.ToString() : null,
                        POSTDeviationCostperpaxabove1500ApprovalDate = row.TryGetValue("POST-Deviation Costperpaxabove1500 Approval Date", out var POSTDeviationCostperpaxabove1500ApprovalDate) ? POSTDeviationCostperpaxabove1500ApprovalDate?.ToString() : null,
                        POSTDeviationChangeinvenueApproval = row.TryGetValue("POST-Deviation Change in venue Approval", out var POSTDeviationChangeinvenueApproval) ? POSTDeviationChangeinvenueApproval?.ToString() : null,
                        PostChangeInVenueApproved = row.TryGetValue("Post ChangeInVenue Approved", out var PostChangeInVenueApproved) ? PostChangeInVenueApproved?.ToString() : null,
                        POSTDeviationChangeintopicApproval = row.TryGetValue("POST-Deviation Change in topic Approval", out var POSTDeviationChangeintopicApproval) ? POSTDeviationChangeintopicApproval?.ToString() : null,
                        PostChangeInTopicApproved = row.TryGetValue("Post ChangeInTopic Approved", out var PostChangeInTopicApproved) ? PostChangeInTopicApproved?.ToString() : null,
                        POSTDeviationChangeinspeakerApproval = row.TryGetValue("POST-Deviation Change in speaker Approval", out var POSTDeviationChangeinspeakerApproval) ? POSTDeviationChangeinspeakerApproval?.ToString() : null,
                        PostChangeInSpeakerApproved = row.TryGetValue("Post ChangeInSpeaker Approved", out var PostChangeInSpeakerApproved) ? PostChangeInSpeakerApproved?.ToString() : null,
                        POSTDeviationAttendeesnotcapturedApproval = row.TryGetValue("POST-Deviation Attendees not captured Approval", out var POSTDeviationAttendeesnotcapturedApproval) ? POSTDeviationAttendeesnotcapturedApproval?.ToString() : null,
                        PostAttendeesNotCapturedApproved = row.TryGetValue("Post AttendeesNotCaptured Approved", out var PostAttendeesNotCapturedApproved) ? PostAttendeesNotCapturedApproved?.ToString() : null,
                        POSTDeviationSpeakernotcapturedApproval = row.TryGetValue("POST-Deviation Speaker not captured  Approval", out var POSTDeviationSpeakernotcapturedApproval) ? POSTDeviationSpeakernotcapturedApproval?.ToString() : null,
                        PostSpeakerNotCapturedApproved = row.TryGetValue("Post SpeakerNotCaptured Approved", out var PostSpeakerNotCapturedApproved) ? PostSpeakerNotCapturedApproved?.ToString() : null,
                        POSTDeviationOtherDeviationApproval = row.TryGetValue("POST-Deviation Other Deviation Approval", out var POSTDeviationOtherDeviationApproval) ? POSTDeviationOtherDeviationApproval?.ToString() : null,
                        PostOtherDeviationApproved = row.TryGetValue("Post OtherDeviation Approved", out var PostOtherDeviationApproved) ? PostOtherDeviationApproved?.ToString() : null,
                        EventSettlementDeviationDate = row.TryGetValue("EventSettlement - Deviation Date", out var EventSettlementDeviationDate) ? EventSettlementDeviationDate?.ToString() : null,
                        EventSettlementDeviationApproval = row.TryGetValue("EventSettlement - Deviation Approval", out var EventSettlementDeviationApproval) ? EventSettlementDeviationApproval?.ToString() : null,
                        EventSettlementDeviationApprovalDate = row.TryGetValue("EventSettlement - Deviation Approval Date", out var EventSettlementDeviationApprovalDate) ? EventSettlementDeviationApprovalDate?.ToString() : null,
                        FinanceDeviationPending = row.TryGetValue("Finance Deviation Pending", out var FinanceDeviationPending) ? FinanceDeviationPending?.ToString() : null,
                        SalesDeviationPending = row.TryGetValue("Sales Deviation Pending", out var SalesDeviationPending) ? SalesDeviationPending?.ToString() : null,
                        HonorariumSubmitted = row.TryGetValue("Honorarium Submitted?", out var HonorariumSubmitted) ? HonorariumSubmitted?.ToString() : null,
                        Honorariumamount = row.TryGetValue("Honorariumamount", out var Honorariumamount) ? Honorariumamount?.ToString() : null,
                        IsItincludingGST = row.TryGetValue("IsItincludingGST?", out var IsItincludingGST) ? IsItincludingGST?.ToString() : null,
                        AgreementAmount = row.TryGetValue("EvvalDate", out var AgreementAmount) ? AgreementAmount?.ToString() : null,
                        JVNo = row.TryGetValue("JV No", out var eventRequestStatus) ? eventRequestStatus?.ToString() : null,
                        JVDate = row.TryGetValue("JV Date", out var eventApprovedDate) ? eventApprovedDate?.ToString() : null,
                        PVNo = row.TryGetValue("PV No", out var miplInvitees) ? miplInvitees?.ToString() : null,
                        PVDate = row.TryGetValue("PV Date", out var postEventRequeststatus) ? postEventRequeststatus?.ToString() : null,
                        PostEventSubmitted = row.TryGetValue("PostEventSubmitted?", out var PostEventSubmitted) ? PostEventSubmitted?.ToString() : null,
                        PostEventBTCExpense = row.TryGetValue("PostEventBTCExpense", out var PostEventBTCExpense) ? PostEventBTCExpense?.ToString() : null,
                        PostEventBTEExpense = row.TryGetValue("PostEventBTEExpense", out var PostEventBTEExpense) ? PostEventBTEExpense?.ToString() : null,
                        ActualAmountGreaterThan50Per = row.TryGetValue("ActualAmountGreaterThan50%", out var ActualAmountGreaterThan50) ? ActualAmountGreaterThan50?.ToString() : null,
                        ReportingManager = row.TryGetValue("Reporting Manager", out var ReportingManager) ? ReportingManager?.ToString() : null,
                        _1UpManager = row.TryGetValue("1 Up Manager", out var _1UpManager) ? _1UpManager?.ToString() : null,
                        Brands = row.TryGetValue("Brands", out var brands) ? brands?.ToString() : null,
                        Panelists = row.TryGetValue("Panelists", out var panelists) ? panelists?.ToString() : null,
                        HCP = row.TryGetValue("HCP", out var HCP) ? HCP?.ToString() : null,
                        SlideKits = row.TryGetValue("SlideKits", out var slideKits) ? slideKits?.ToString() : null,
                        Invitees = row.TryGetValue("Invitees", out var invitees) ? invitees?.ToString() : null,
                        Expenses = row.TryGetValue("Expenses", out var expenses) ? expenses?.ToString() : null,
                        IndicationsDone = row.TryGetValue("Indications Done", out var IndicationsDone) ? IndicationsDone?.ToString() : null,
                        TotalInvitees = row.TryGetValue("Total Invitees", out var ctotalInvitees) && double.TryParse(ctotalInvitees?.ToString(), out var vtotalInvitees) ? vtotalInvitees : 0,
                        TotalAttendees = row.TryGetValue("Total Attendees", out var totalAttendees) && double.TryParse(totalAttendees?.ToString(), out var parsedTotalAttendees) ? parsedTotalAttendees : 0,
                        ApprovalStatus = row.TryGetValue("Approval Status", out var eventOpenSalesHeadApprovalDate) ? eventOpenSalesHeadApprovalDate?.ToString() : null,
                        NextApprover = row.TryGetValue("Next Approver", out var _7daysSalesHeadApproval) ? _7daysSalesHeadApproval?.ToString() : null,
                        RBMBM = row.TryGetValue("RBM/BM", out var rbmbm) ? rbmbm?.ToString() : null,
                        SalesHead = row.TryGetValue("Sales Head", out var salesHead) ? salesHead?.ToString() : null,
                        MarketingHead = row.TryGetValue("Marketing Head", out var marketingHead) ? marketingHead?.ToString() : null,
                        Finance = row.TryGetValue("Finance", out var finance) ? finance?.ToString() : null,
                        Compliance = row.TryGetValue("Compliance", out var compliance) ? compliance?.ToString() : null,
                        FinanceTreasury = row.TryGetValue("Finance Treasury", out var financeTreasury) ? financeTreasury?.ToString() : null,
                        FinanceAccounts = row.TryGetValue("Finance Accounts", out var financeAccounts) ? financeAccounts?.ToString() : null,
                        MedicalAffairsHead = row.TryGetValue("Medical Affairs Head", out var medicalAffairsHead) ? medicalAffairsHead?.ToString() : null,
                        SalesCoordinator = row.TryGetValue("Sales Coordinator", out var salesCoordinator) ? salesCoordinator?.ToString() : null,
                        FinanceHead = row.TryGetValue("Finance Head", out var financeHead) ? financeHead?.ToString() : null,
                        TotalHonorariumAmount = row.TryGetValue("Total Honorarium Amount", out var totalHonorariumAmount) && double.TryParse(totalHonorariumAmount?.ToString(), out var parsedTotalHonorariumAmount) ? parsedTotalHonorariumAmount : 0,
                        TotalTravelAccommodationAmount = row.TryGetValue("Total Travel & Accommodation Amount", out var totalTravelAccommodationAmount) && double.TryParse(totalTravelAccommodationAmount?.ToString(), out var parsedTotalTravelAccommodationAmount) ? parsedTotalTravelAccommodationAmount : 0,
                        TotalTravelAmount = row.TryGetValue("Total Travel Amount", out var totalTravelAmount) && double.TryParse(totalTravelAmount?.ToString(), out var parsedTotalTravelAmount) ? parsedTotalTravelAmount : 0,
                        TotalAccommodationAmount = row.TryGetValue("Total Accommodation Amount", out var totalAccommodationAmount) && double.TryParse(totalAccommodationAmount?.ToString(), out var parsedTotalAccommodationAmount) ? parsedTotalAccommodationAmount : 0,
                        TotalLocalConveyance = row.TryGetValue("Total Local Conveyance", out var totalLocalConveyance) && double.TryParse(totalLocalConveyance?.ToString(), out var parsedTotalLocalConveyance) ? parsedTotalLocalConveyance : 0,
                        TotalExpense = row.TryGetValue("Total Expense", out var totalExpense) && double.TryParse(totalExpense?.ToString(), out var parsedTotalExpense) ? parsedTotalExpense : 0,
                        TotalBudget = row.TryGetValue("Total Budget", out var totalBudget) && double.TryParse(totalBudget?.ToString(), out var parsedTotalBudget) ? parsedTotalBudget : 0,
                        TotalActual = row.TryGetValue("Total Budget", out var TotalActual) && double.TryParse(TotalActual?.ToString(), out var parsedTotalActual) ? parsedTotalActual : 0,
                        _50Helper = row.TryGetValue("50% Helper", out var _50Helper) ? _50Helper?.ToString() : null,
                        CostperparticipantHelper = row.TryGetValue("Cost per participant - Helper", out var _7daysSalesHeadApprovaldate) ? _7daysSalesHeadApprovaldate?.ToString() : null,
                        BTCexceeds50ofbudget = row.TryGetValue("BTC exceeds 50% of budget", out var BTCexceeds50ofbudget) ? BTCexceeds50ofbudget?.ToString() : null,
                        FinanceAccountsGivenDetails = row.TryGetValue("Finance Accounts Given Details", out var FinanceAccountsGivenDetails) ? FinanceAccountsGivenDetails?.ToString() : null,
                        FinanceTreasuryGivenDetails = row.TryGetValue("Finance Treasury Given Details", out var FinanceTreasuryGivenDetails) ? FinanceTreasuryGivenDetails?.ToString() : null,
                        Role = row.TryGetValue("Role", out var Role) ? Role?.ToString() : null,
                        AdvanceAmount = row.TryGetValue("Advance Amount", out var advanceAmount) && double.TryParse(advanceAmount?.ToString(), out var parsedAdvanceAmount) ? parsedAdvanceAmount : 0,
                        TotalExpenseBTC = row.TryGetValue(" Total Expense BTC", out var totalExpenseBTC) && double.TryParse(totalExpenseBTC?.ToString(), out var parsedTotalExpenseBTC) ? parsedTotalExpenseBTC : 0,
                        TotalExpenseBTE = row.TryGetValue("Total Expense BTE", out var totalExpenseBTE) && double.TryParse(totalExpenseBTE?.ToString(), out var parsedTotalExpenseBTE) ? parsedTotalExpenseBTE : 0,
                        HelperFinancetreasurytriggerBTE = row.TryGetValue("Helper Finance treasury trigger(BTE)", out var HelperFinancetreasurytriggerBTE) ? HelperFinancetreasurytriggerBTE?.ToString() : null,
                        ClassIIIEventCode = row.TryGetValue("Class III Event Code", out var ClassIIIEventCode) ? ClassIIIEventCode?.ToString() : null,
                        MeetingType = row.TryGetValue("Meeting Type", out var MeetingType) ? MeetingType?.ToString() : null,
                        SponsorshipSocietyName = row.TryGetValue("Sponsorship Society Name", out var SponsorshipSocietyName) ? SponsorshipSocietyName?.ToString() : null,
                        VenueCountry = row.TryGetValue("Venue Country", out var VenueCountry) ? VenueCountry?.ToString() : null,
                        TotalHCPRegistrationAmount = row.TryGetValue("Total HCP Registration Amount", out var TotalHCPRegistrationAmount) ? TotalHCPRegistrationAmount?.ToString() : null,
                        MedicalUtilityType = row.TryGetValue("Medical Utility Type", out var MedicalUtilityType) ? MedicalUtilityType?.ToString() : null,
                        MedicalUtilityDescription = row.TryGetValue("Medical Utility Description", out var MedicalUtilityDescription) ? MedicalUtilityDescription?.ToString() : null,
                        ValidTo = row.TryGetValue("Valid To", out var ValidTo) ? ValidTo?.ToString() : null,
                        ValidFrom = row.TryGetValue("Valid From", out var ValidFrom) ? ValidFrom?.ToString() : null,

                    }).ToList();

                return Ok(eventRequestBrandsList);
            }
            else
            {
                return BadRequest(new
                {
                    Message = "Row not found"
                });
            }
            return Ok();
        }
        [HttpGet("GetDataFromEventSettlementSheet")]
        public IActionResult GetDataFromEventSettlementSheet()
        {
            string sheetId = configuration.GetSection("SmartsheetSettings:EventSettlement").Value;
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
            Dictionary<string, object> ProductBrandsListrowData = new Dictionary<string, object>();
            //Row ExistingRow = sheet.Rows.FirstOrDefault(row => row.)
            //Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == eventId));
            // if (targetRow != null)
            //  {
            //long Id = targetRow.Id.Value;
            List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);

            List<EventSettlementPayload> eventRequestBrandsList = sheetData
                // .Where(row => row.TryGetValue("EventId/EventRequestId", out var eventIdValue) && eventIdValue?.ToString() == eventId)
                .Select(row => new EventSettlementPayload
                {
                    EventType = row.TryGetValue("EventType", out var eventType) ? eventType?.ToString() : null,
                    EventIdEventRequestId = row.TryGetValue("EventId/EventRequestId", out var eventId) ? eventId?.ToString() : null,
                    EventTopic = row.TryGetValue("EventTopic", out var eventTopic) ? eventTopic?.ToString() : null,
                    EventDate = row.TryGetValue("EventDate", out var eventDate) ? eventDate?.ToString() : null,
                    EventEndDate = row.TryGetValue("Event End Date", out var eventEndDate) ? eventEndDate?.ToString() : null,
                    StartTime = row.TryGetValue("StartTime", out var startTime) ? startTime?.ToString() : null,
                    EndTime = row.TryGetValue("EndTime", out var endTime) ? endTime?.ToString() : null,
                    VenueName = row.TryGetValue("VenueName", out var venueName) ? venueName?.ToString() : null,
                    State = row.TryGetValue("State", out var state) ? state?.ToString() : null,
                    City = row.TryGetValue("City", out var city) ? city?.ToString() : null,
                    Attended = row.TryGetValue("Attended", out var totalInvitees) && double.TryParse(totalInvitees?.ToString(), out var parsedValue) ? parsedValue : 0,
                    InviteesParticipated = row.TryGetValue("InviteesParticipated", out var honorariumApprovedDate) ? honorariumApprovedDate?.ToString() : null,
                    ExpenseDetails = row.TryGetValue("ExpenseDetails", out var ExpenseDetails) ? ExpenseDetails?.ToString() : null,
                    TotalExpenseDetails = row.TryGetValue("TotalExpenseDetails", out var TotalExpenseDetails) ? TotalExpenseDetails?.ToString() : null,
                    AdvanceDetails = row.TryGetValue("AdvanceDetails", out var AdvanceDetails) ? AdvanceDetails?.ToString() : null,
                    AdvanceUtilizedForEvent = row.TryGetValue("Advance Utilized For Event", out var AdvanceUtilizedForEvent) ? AdvanceUtilizedForEvent?.ToString() : null,
                    PayBackAmountToCompany = row.TryGetValue("Pay Back Amount To Company", out var PayBackAmountToCompany) ? PayBackAmountToCompany?.ToString() : null,
                    AdditionalAmountNeededToPayForInitiator = row.TryGetValue("Additional Amount Needed To Pay For Initiator", out var AdditionalAmountNeededToPayForInitiator) ? AdditionalAmountNeededToPayForInitiator?.ToString() : null,
                    CreatedOn = row.TryGetValue("CreatedOn", out var premarketingHeadApproval) ? premarketingHeadApproval?.ToString() : null,
                    CreatedDateHelper = row.TryGetValue("Created Date - Helper", out var createdDateHelper) ? createdDateHelper?.ToString() : null,
                    Modified = row.TryGetValue("Modified", out var premarketingHeadApprovalDate) ? premarketingHeadApprovalDate?.ToString() : null,
                    InitiatorName = row.TryGetValue("InitiatorName", out var initiatorName) ? initiatorName?.ToString() : null,
                    InitiatorEmail = row.TryGetValue("Initiator Email", out var initiatorEmail) ? initiatorEmail?.ToString() : null,
                    IsAdvanceRequired = row.TryGetValue("IsAdvanceRequired", out var isAdvanceRequired) ? isAdvanceRequired?.ToString() : null,
                    EventRequestStatus = row.TryGetValue("Event Request Status", out var prerBmBmApprovalDate) ? prerBmBmApprovalDate?.ToString() : null,
                    HonorariumRequestStatus = row.TryGetValue("Honorarium Request Status", out var honorariumRequestStatus) ? honorariumRequestStatus?.ToString() : null,
                    PostEventRequeststatus = row.TryGetValue("Post Event Request status", out var PostEventRequeststatus) ? PostEventRequeststatus?.ToString() : null,
                    PostEventApprovedDate = row.TryGetValue("Post Event Approved Date", out var PostEventApprovedDate) ? PostEventApprovedDate?.ToString() : null,
                    EventSettlementRBMBMApproval = row.TryGetValue("EventSettlement-RBM/BM Approval", out var agreement) ? agreement?.ToString() : null,
                    EventSettlementRBMBMApprovalDate = row.TryGetValue("EventSettlement-RBM/BM Approval Date", out var HONRBMBMApprovalDate) ? HONRBMBMApprovalDate?.ToString() : null,
                    EventSettlementSalesHeadApproval = row.TryGetValue("EventSettlement-SalesHead Approval", out var HONSalesHeadApproval) ? HONSalesHeadApproval?.ToString() : null,
                    EventSettlementSalesHeadApprovalDate = row.TryGetValue("EventSettlementSalesHeadApprovalDate", out var HONSalesHeadApprovalDate) ? HONSalesHeadApprovalDate?.ToString() : null,
                    EventSettlementComplianceApproval = row.TryGetValue("EventSettlement-Compliance Approval", out var EventSettlementComplianceApproval) ? EventSettlementComplianceApproval?.ToString() : null,
                    EventSettlementComplianceApprovalDate = row.TryGetValue("EventSettlement-Compliance Approval Date", out var EventSettlementComplianceApprovalDate) ? EventSettlementComplianceApprovalDate?.ToString() : null,
                    EventSettlementFinanceAccountApproval = row.TryGetValue("EventSettlement-Finance Account Approval", out var EventSettlementFinanceAccountApproval) ? EventSettlementFinanceAccountApproval?.ToString() : null,
                    EventSettlementFinanceAccountComments = row.TryGetValue("EventSettlement-Finance Account Comments", out var EventSettlementFinanceAccountComments) ? EventSettlementFinanceAccountComments?.ToString() : null,
                    EventSettlementFinanceAccountApprovalDate = row.TryGetValue("EventSettlement-Finance Account Approval Date", out var EventSettlementFinanceAccountApprovalDate) ? EventSettlementFinanceAccountApprovalDate?.ToString() : null,
                    EventSettlementFinanceTreasuryApproval = row.TryGetValue("EventSettlement-Finance Treasury Approval Date", out var EventSettlementFinanceTreasuryApproval) ? EventSettlementFinanceTreasuryApproval?.ToString() : null,
                    EventSettlementMarketingHeadApproval = row.TryGetValue("EventSettlement-Marketing Head Approval Date", out var EventSettlementMarketingHeadApproval) ? EventSettlementMarketingHeadApproval?.ToString() : null,
                    EventSettlementMedicalAffairsHeadApproval = row.TryGetValue("EventSettlement-Medical Affairs Head Approval", out var EventSettlementMedicalAffairsHeadApproval) ? EventSettlementMedicalAffairsHeadApproval?.ToString() : null,
                    EventSettlementMedicalAffairsHeadApprovalDate = row.TryGetValue("EventSettlement-Medical Affairs Head Approval Date", out var EventSettlementMedicalAffairsHeadApprovalDate) ? EventSettlementMedicalAffairsHeadApprovalDate?.ToString() : null,
                    DeviationStatus = row.TryGetValue("Deviation Status", out var DeviationStatus) ? DeviationStatus?.ToString() : null,
                    IsAllDeviationsApproved = row.TryGetValue("Is All Deviations Approved?", out var IsAllDeviationsApproved) ? IsAllDeviationsApproved?.ToString() : null,
                    POSTBeyond30DaysDeviationApproval = row.TryGetValue("POST- Beyond30Days Deviation Approval", out var POSTBeyond30DaysDeviationApproval) ? POSTBeyond30DaysDeviationApproval?.ToString() : null,
                    post45daysapproved = row.TryGetValue("post 45 days approved", out var post45daysapproved) ? post45daysapproved?.ToString() : null,
                    POSTBeyond30DaysDeviationApprovalDate = row.TryGetValue("POST- Beyond30Days Deviation Approval Date", out var POSTBeyond30DaysDeviationApprovalDate) ? POSTBeyond30DaysDeviationApprovalDate?.ToString() : null,
                    POSTLessThan5InviteesDeviationApproval = row.TryGetValue("POST-LessThan5Invitees Deviation Approval", out var POSTLessThan5InviteesDeviationApproval) ? POSTLessThan5InviteesDeviationApproval?.ToString() : null,
                    Post5InviteesApproved = row.TryGetValue("Post <5 Invitees Approved", out var Post5InviteesApproved) ? Post5InviteesApproved?.ToString() : null,
                    POSTDeviationCostperpaxabove1500Approval = row.TryGetValue("POST-Deviation Costperpaxabove1500 Approval ", out var POSTDeviationCostperpaxabove1500Approval) ? POSTDeviationCostperpaxabove1500Approval?.ToString() : null,
                    PostCostperPaxApproved = row.TryGetValue("Post CostperPax Approved", out var PostCostperPaxApproved) ? PostCostperPaxApproved?.ToString() : null,
                    POSTDeviationCostperpaxabove1500ApprovalDate = row.TryGetValue("POST-Deviation Costperpaxabove1500 Approval Date", out var POSTDeviationCostperpaxabove1500ApprovalDate) ? POSTDeviationCostperpaxabove1500ApprovalDate?.ToString() : null,
                    POSTDeviationChangeinvenueApproval = row.TryGetValue("POST-Deviation Change in venue Approval", out var POSTDeviationChangeinvenueApproval) ? POSTDeviationChangeinvenueApproval?.ToString() : null,
                    PostChangeInVenueApproved = row.TryGetValue("Post ChangeInVenue Approved", out var PostChangeInVenueApproved) ? PostChangeInVenueApproved?.ToString() : null,
                    POSTDeviationChangeintopicApproval = row.TryGetValue("POST-Deviation Change in topic Approval", out var POSTDeviationChangeintopicApproval) ? POSTDeviationChangeintopicApproval?.ToString() : null,
                    PostChangeInTopicApproved = row.TryGetValue("Post ChangeInTopic Approved", out var PostChangeInTopicApproved) ? PostChangeInTopicApproved?.ToString() : null,
                    POSTDeviationChangeinspeakerApproval = row.TryGetValue("POST-Deviation Change in speaker Approval", out var POSTDeviationChangeinspeakerApproval) ? POSTDeviationChangeinspeakerApproval?.ToString() : null,
                    PostChangeInSpeakerApproved = row.TryGetValue("Post ChangeInSpeaker Approved", out var PostChangeInSpeakerApproved) ? PostChangeInSpeakerApproved?.ToString() : null,
                    POSTDeviationAttendeesnotcapturedApproval = row.TryGetValue("POST-Deviation Attendees not captured Approval", out var POSTDeviationAttendeesnotcapturedApproval) ? POSTDeviationAttendeesnotcapturedApproval?.ToString() : null,
                    PostAttendeesNotCapturedApproved = row.TryGetValue("Post AttendeesNotCaptured Approved", out var PostAttendeesNotCapturedApproved) ? PostAttendeesNotCapturedApproved?.ToString() : null,
                    POSTDeviationSpeakernotcapturedApproval = row.TryGetValue("POST-Deviation Speaker not captured  Approval", out var POSTDeviationSpeakernotcapturedApproval) ? POSTDeviationSpeakernotcapturedApproval?.ToString() : null,
                    PostSpeakerNotCapturedApproved = row.TryGetValue("Post SpeakerNotCaptured Approved", out var PostSpeakerNotCapturedApproved) ? PostSpeakerNotCapturedApproved?.ToString() : null,
                    POSTDeviationOtherDeviationApproval = row.TryGetValue("POST-Deviation Other Deviation Approval", out var POSTDeviationOtherDeviationApproval) ? POSTDeviationOtherDeviationApproval?.ToString() : null,
                    PostOtherDeviationApproved = row.TryGetValue("Post OtherDeviation Approved", out var PostOtherDeviationApproved) ? PostOtherDeviationApproved?.ToString() : null,
                    EventSettlementDeviationDate = row.TryGetValue("EventSettlement - Deviation Date", out var EventSettlementDeviationDate) ? EventSettlementDeviationDate?.ToString() : null,
                    EventSettlementDeviationApproval = row.TryGetValue("EventSettlement - Deviation Approval", out var EventSettlementDeviationApproval) ? EventSettlementDeviationApproval?.ToString() : null,
                    EventSettlementDeviationApprovalDate = row.TryGetValue("EventSettlement - Deviation Approval Date", out var EventSettlementDeviationApprovalDate) ? EventSettlementDeviationApprovalDate?.ToString() : null,
                    FinanceDeviationPending = row.TryGetValue("Finance Deviation Pending", out var FinanceDeviationPending) ? FinanceDeviationPending?.ToString() : null,
                    SalesDeviationPending = row.TryGetValue("Sales Deviation Pending", out var SalesDeviationPending) ? SalesDeviationPending?.ToString() : null,
                    HonorariumSubmitted = row.TryGetValue("Honorarium Submitted?", out var HonorariumSubmitted) ? HonorariumSubmitted?.ToString() : null,
                    Honorariumamount = row.TryGetValue("Honorariumamount", out var Honorariumamount) ? Honorariumamount?.ToString() : null,
                    IsItincludingGST = row.TryGetValue("IsItincludingGST?", out var IsItincludingGST) ? IsItincludingGST?.ToString() : null,
                    AgreementAmount = row.TryGetValue("EvvalDate", out var AgreementAmount) ? AgreementAmount?.ToString() : null,
                    JVNo = row.TryGetValue("JV No", out var eventRequestStatus) ? eventRequestStatus?.ToString() : null,
                    JVDate = row.TryGetValue("JV Date", out var eventApprovedDate) ? eventApprovedDate?.ToString() : null,
                    PVNo = row.TryGetValue("PV No", out var miplInvitees) ? miplInvitees?.ToString() : null,
                    PVDate = row.TryGetValue("PV Date", out var postEventRequeststatus) ? postEventRequeststatus?.ToString() : null,
                    PostEventSubmitted = row.TryGetValue("PostEventSubmitted?", out var PostEventSubmitted) ? PostEventSubmitted?.ToString() : null,
                    PostEventBTCExpense = row.TryGetValue("PostEventBTCExpense", out var PostEventBTCExpense) ? PostEventBTCExpense?.ToString() : null,
                    PostEventBTEExpense = row.TryGetValue("PostEventBTEExpense", out var PostEventBTEExpense) ? PostEventBTEExpense?.ToString() : null,
                    ActualAmountGreaterThan50Per = row.TryGetValue("ActualAmountGreaterThan50%", out var ActualAmountGreaterThan50) ? ActualAmountGreaterThan50?.ToString() : null,
                    ReportingManager = row.TryGetValue("Reporting Manager", out var ReportingManager) ? ReportingManager?.ToString() : null,
                    _1UpManager = row.TryGetValue("1 Up Manager", out var _1UpManager) ? _1UpManager?.ToString() : null,
                    Brands = row.TryGetValue("Brands", out var brands) ? brands?.ToString() : null,
                    Panelists = row.TryGetValue("Panelists", out var panelists) ? panelists?.ToString() : null,
                    HCP = row.TryGetValue("HCP", out var HCP) ? HCP?.ToString() : null,
                    SlideKits = row.TryGetValue("SlideKits", out var slideKits) ? slideKits?.ToString() : null,
                    Invitees = row.TryGetValue("Invitees", out var invitees) ? invitees?.ToString() : null,
                    Expenses = row.TryGetValue("Expenses", out var expenses) ? expenses?.ToString() : null,
                    IndicationsDone = row.TryGetValue("Indications Done", out var IndicationsDone) ? IndicationsDone?.ToString() : null,
                    TotalInvitees = row.TryGetValue("Total Invitees", out var ctotalInvitees) && double.TryParse(ctotalInvitees?.ToString(), out var vtotalInvitees) ? vtotalInvitees : 0,
                    TotalAttendees = row.TryGetValue("Total Attendees", out var totalAttendees) && double.TryParse(totalAttendees?.ToString(), out var parsedTotalAttendees) ? parsedTotalAttendees : 0,
                    ApprovalStatus = row.TryGetValue("Approval Status", out var eventOpenSalesHeadApprovalDate) ? eventOpenSalesHeadApprovalDate?.ToString() : null,
                    NextApprover = row.TryGetValue("Next Approver", out var _7daysSalesHeadApproval) ? _7daysSalesHeadApproval?.ToString() : null,
                    RBMBM = row.TryGetValue("RBM/BM", out var rbmbm) ? rbmbm?.ToString() : null,
                    SalesHead = row.TryGetValue("Sales Head", out var salesHead) ? salesHead?.ToString() : null,
                    MarketingHead = row.TryGetValue("Marketing Head", out var marketingHead) ? marketingHead?.ToString() : null,
                    Finance = row.TryGetValue("Finance", out var finance) ? finance?.ToString() : null,
                    Compliance = row.TryGetValue("Compliance", out var compliance) ? compliance?.ToString() : null,
                    FinanceTreasury = row.TryGetValue("Finance Treasury", out var financeTreasury) ? financeTreasury?.ToString() : null,
                    FinanceAccounts = row.TryGetValue("Finance Accounts", out var financeAccounts) ? financeAccounts?.ToString() : null,
                    MedicalAffairsHead = row.TryGetValue("Medical Affairs Head", out var medicalAffairsHead) ? medicalAffairsHead?.ToString() : null,
                    SalesCoordinator = row.TryGetValue("Sales Coordinator", out var salesCoordinator) ? salesCoordinator?.ToString() : null,
                    FinanceHead = row.TryGetValue("Finance Head", out var financeHead) ? financeHead?.ToString() : null,
                    TotalHonorariumAmount = row.TryGetValue("Total Honorarium Amount", out var totalHonorariumAmount) && double.TryParse(totalHonorariumAmount?.ToString(), out var parsedTotalHonorariumAmount) ? parsedTotalHonorariumAmount : 0,
                    TotalTravelAccommodationAmount = row.TryGetValue("Total Travel & Accommodation Amount", out var totalTravelAccommodationAmount) && double.TryParse(totalTravelAccommodationAmount?.ToString(), out var parsedTotalTravelAccommodationAmount) ? parsedTotalTravelAccommodationAmount : 0,
                    TotalTravelAmount = row.TryGetValue("Total Travel Amount", out var totalTravelAmount) && double.TryParse(totalTravelAmount?.ToString(), out var parsedTotalTravelAmount) ? parsedTotalTravelAmount : 0,
                    TotalAccommodationAmount = row.TryGetValue("Total Accommodation Amount", out var totalAccommodationAmount) && double.TryParse(totalAccommodationAmount?.ToString(), out var parsedTotalAccommodationAmount) ? parsedTotalAccommodationAmount : 0,
                    TotalLocalConveyance = row.TryGetValue("Total Local Conveyance", out var totalLocalConveyance) && double.TryParse(totalLocalConveyance?.ToString(), out var parsedTotalLocalConveyance) ? parsedTotalLocalConveyance : 0,
                    TotalExpense = row.TryGetValue("Total Expense", out var totalExpense) && double.TryParse(totalExpense?.ToString(), out var parsedTotalExpense) ? parsedTotalExpense : 0,
                    TotalBudget = row.TryGetValue("Total Budget", out var totalBudget) && double.TryParse(totalBudget?.ToString(), out var parsedTotalBudget) ? parsedTotalBudget : 0,
                    TotalActual = row.TryGetValue("Total Budget", out var TotalActual) && double.TryParse(TotalActual?.ToString(), out var parsedTotalActual) ? parsedTotalActual : 0,
                    _50Helper = row.TryGetValue("50% Helper", out var _50Helper) ? _50Helper?.ToString() : null,
                    CostperparticipantHelper = row.TryGetValue("Cost per participant - Helper", out var _7daysSalesHeadApprovaldate) ? _7daysSalesHeadApprovaldate?.ToString() : null,
                    BTCexceeds50ofbudget = row.TryGetValue("BTC exceeds 50% of budget", out var BTCexceeds50ofbudget) ? BTCexceeds50ofbudget?.ToString() : null,
                    FinanceAccountsGivenDetails = row.TryGetValue("Finance Accounts Given Details", out var FinanceAccountsGivenDetails) ? FinanceAccountsGivenDetails?.ToString() : null,
                    FinanceTreasuryGivenDetails = row.TryGetValue("Finance Treasury Given Details", out var FinanceTreasuryGivenDetails) ? FinanceTreasuryGivenDetails?.ToString() : null,
                    Role = row.TryGetValue("Role", out var Role) ? Role?.ToString() : null,
                    AdvanceAmount = row.TryGetValue("Advance Amount", out var advanceAmount) && double.TryParse(advanceAmount?.ToString(), out var parsedAdvanceAmount) ? parsedAdvanceAmount : 0,
                    TotalExpenseBTC = row.TryGetValue(" Total Expense BTC", out var totalExpenseBTC) && double.TryParse(totalExpenseBTC?.ToString(), out var parsedTotalExpenseBTC) ? parsedTotalExpenseBTC : 0,
                    TotalExpenseBTE = row.TryGetValue("Total Expense BTE", out var totalExpenseBTE) && double.TryParse(totalExpenseBTE?.ToString(), out var parsedTotalExpenseBTE) ? parsedTotalExpenseBTE : 0,
                    HelperFinancetreasurytriggerBTE = row.TryGetValue("Helper Finance treasury trigger(BTE)", out var HelperFinancetreasurytriggerBTE) ? HelperFinancetreasurytriggerBTE?.ToString() : null,
                    ClassIIIEventCode = row.TryGetValue("Class III Event Code", out var ClassIIIEventCode) ? ClassIIIEventCode?.ToString() : null,
                    MeetingType = row.TryGetValue("Meeting Type", out var MeetingType) ? MeetingType?.ToString() : null,
                    SponsorshipSocietyName = row.TryGetValue("Sponsorship Society Name", out var SponsorshipSocietyName) ? SponsorshipSocietyName?.ToString() : null,
                    VenueCountry = row.TryGetValue("Venue Country", out var VenueCountry) ? VenueCountry?.ToString() : null,
                    TotalHCPRegistrationAmount = row.TryGetValue("Total HCP Registration Amount", out var TotalHCPRegistrationAmount) ? TotalHCPRegistrationAmount?.ToString() : null,
                    MedicalUtilityType = row.TryGetValue("Medical Utility Type", out var MedicalUtilityType) ? MedicalUtilityType?.ToString() : null,
                    MedicalUtilityDescription = row.TryGetValue("Medical Utility Description", out var MedicalUtilityDescription) ? MedicalUtilityDescription?.ToString() : null,
                    ValidTo = row.TryGetValue("Valid To", out var ValidTo) ? ValidTo?.ToString() : null,
                    ValidFrom = row.TryGetValue("Valid From", out var ValidFrom) ? ValidFrom?.ToString() : null,

                }).ToList();

            return Ok(eventRequestBrandsList);
            //}
            //else
            //{
            //    return BadRequest(new
            //    {
            //        Message = "Row not found"
            //    });
            //}
            //return Ok();
        }

        [HttpGet("GetAttachmentsFromProcessSheetBasedOnEventId")]
        public IActionResult GetAttachmentsFromProcessSheetBasedOnEventId(string eventId)
        {
            string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
            Dictionary<string, object> ProductBrandsListrowData = new Dictionary<string, object>();
            Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == eventId));
            if (targetRow != null)
            {
                var attachments = GetAttachmentsForRow(smartsheet, sheet.Id.Value, targetRow.Id.Value);
                ProductBrandsListrowData["Attachments"] = attachments;
            }
            return Ok(ProductBrandsListrowData);
        }
        [HttpGet("GetAttachmentsFromDeviationSheetBasedOnEventId")]
        public IActionResult GetAttachmentsFromDeviationSheetBasedOnEventId(string eventId)
        {
            string sheetId = configuration.GetSection("SmartsheetSettings:DeviationProcess").Value;
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
            Dictionary<string, object> ProductBrandsListrowData = new Dictionary<string, object>();
            Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == eventId));
            if (targetRow != null)
            {
                var attachments = GetAttachmentsForRow(smartsheet, sheet.Id.Value, targetRow.Id.Value);
                ProductBrandsListrowData["Attachments"] = attachments;
            }
            return Ok(ProductBrandsListrowData);
        }
        [HttpGet("GetAttachmentsFromHonorariumSheetBasedOnEventId")]
        public IActionResult GetAttachmentsFromHonorariumSheetBasedOnEventId(string eventId)
        {
            string sheetId = configuration.GetSection("SmartsheetSettings:HonorariumPayment").Value;
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
            Dictionary<string, object> ProductBrandsListrowData = new Dictionary<string, object>();
            Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == eventId));
            if (targetRow != null)
            {
                var attachments = GetAttachmentsForRow(smartsheet, sheet.Id.Value, targetRow.Id.Value);
                ProductBrandsListrowData["Attachments"] = attachments;
            }
            return Ok(ProductBrandsListrowData);
        }
        [HttpGet("GetAttachmentsFromEventSettlementSheetBasedOnEventId")]
        public IActionResult GetAttachmentsFromEventSettlementSheetBasedOnEventId(string eventId)
        {
            string sheetId = configuration.GetSection("SmartsheetSettings:EventSettlement").Value;
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
            Dictionary<string, object> ProductBrandsListrowData = new Dictionary<string, object>();
            Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == eventId));
            if (targetRow != null)
            {
                var attachments = GetAttachmentsForRow(smartsheet, sheet.Id.Value, targetRow.Id.Value);
                ProductBrandsListrowData["Attachments"] = attachments;
            }
            return Ok(ProductBrandsListrowData);
        }
        [HttpGet("GetAmountsFromAdvanceProvidedColumnUsingEventId")]
        public IActionResult GetAmountsFromAdvanceProvidedColumnUsingEventId(string eventId)
        {
            string sheetId = configuration.GetSection("SmartsheetSettings:EventSettlement").Value;
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
            Column targetColumn = sheet.Columns.FirstOrDefault(column => string.Equals(column.Title, "AdvanceDetails", StringComparison.OrdinalIgnoreCase));
            if (targetColumn != null)
            {
                Row targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == eventId));
                if (targetRow != null)
                {
                    object? columnValue = targetRow.Cells.FirstOrDefault(cell => cell.ColumnId == targetColumn.Id)?.Value;
                    if (columnValue != null)
                    {
                        string? ColumnValueString = columnValue.ToString();
                        string[] lines = ColumnValueString.Split('\n');
                        string? advanceAmount = lines.FirstOrDefault(line => line.Contains("Advance Amount"))?.Split(':')[1]?.Trim();
                        string? totalBudget = lines.FirstOrDefault(line => line.Contains("Total Budget"))?.Split(':')[1]?.Trim();
                        string? actualExpense = lines.FirstOrDefault(line => line.Contains("Actual Expense"))?.Split(':')[1]?.Trim();
                        string? paybackToCompany = lines.FirstOrDefault(line => line.Contains("Payback Amount To Company"))?.Split(':')[1]?.Trim();
                        string? paybackToInitiator = lines.FirstOrDefault(line => line.Contains("Payback Amount To Initiator"))?.Split(':')[1]?.Trim();
                        var jsonResponse = new
                        {
                            AdvanceAmount = advanceAmount,
                            TotalBudget = totalBudget,
                            ActualExpense = actualExpense,
                            PaybackAmountToCompany = paybackToCompany,
                            PaybackAmountToInitiator = paybackToInitiator
                        };
                        return Ok(jsonResponse);
                    }
                }
                else
                {
                    return BadRequest(new { Message = "Target Row Not Found" });
                }
            }
            return Ok();
        }




        private List<Dictionary<string, object>> GetAttachmentsForRow(SmartsheetClient smartsheet, long sheetId, long row)
        {
            List<Dictionary<string, object>> attachmentsList = new List<Dictionary<string, object>>();
            var attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheetId, row, null);

            if (attachments.Data != null && attachments.Data.Count > 0)
            {
                foreach (var attachment in attachments.Data)
                {
                    long AID = (long)attachment.Id;
                    Attachment file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheetId, AID);
                    Dictionary<string, object> attachmentInfo = new Dictionary<string, object>
                    {
                        { "Name", file.Name },
                        { "Id", file.Id },
                        { "Url", file.Url },
                        { "base64", SheetHelper.UrlToBaseValue(file.Url) }
                    };
                    attachmentsList.Add(attachmentInfo);
                }
            }
            return attachmentsList;
        }




        public class getIds
        {
            public List<string> EventIds { get; set; }
        }
    }
}
