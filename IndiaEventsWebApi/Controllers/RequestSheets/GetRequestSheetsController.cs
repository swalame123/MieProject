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
                        TotalInvitees = row.TryGetValue("Total Invitees", out var totalInvitees) && int.TryParse(totalInvitees?.ToString(), out var parsedValue) ? parsedValue : 0,
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

                        TotalAttendees = row.TryGetValue("Total Attendees", out var totalAttendees) && int.TryParse(totalAttendees?.ToString(), out var parsedTotalAttendees) ? parsedTotalAttendees : 0,
                        TotalHonorariumAmount = row.TryGetValue("Total Honorarium Amount", out var totalHonorariumAmount) && int.TryParse(totalHonorariumAmount?.ToString(), out var parsedTotalHonorariumAmount) ? parsedTotalHonorariumAmount : 0,
                        TotalTravelAccommodationAmount = row.TryGetValue("Total Travel & Accommodation Amount", out var totalTravelAccommodationAmount) && int.TryParse(totalTravelAccommodationAmount?.ToString(), out var parsedTotalTravelAccommodationAmount) ? parsedTotalTravelAccommodationAmount : 0,
                        TotalTravelAmount = row.TryGetValue("Total Travel Amount", out var totalTravelAmount) && int.TryParse(totalTravelAmount?.ToString(), out var parsedTotalTravelAmount) ? parsedTotalTravelAmount : 0,
                        TotalAccommodationAmount = row.TryGetValue("Total Accommodation Amount", out var totalAccommodationAmount) && int.TryParse(totalAccommodationAmount?.ToString(), out var parsedTotalAccommodationAmount) ? parsedTotalAccommodationAmount : 0,
                        TotalLocalConveyance = row.TryGetValue("Total Local Conveyance", out var totalLocalConveyance) && int.TryParse(totalLocalConveyance?.ToString(), out var parsedTotalLocalConveyance) ? parsedTotalLocalConveyance : 0,
                        TotalExpense = row.TryGetValue("Total Expense", out var totalExpense) && int.TryParse(totalExpense?.ToString(), out var parsedTotalExpense) ? parsedTotalExpense : 0,
                        OtherExpenseAmount = row.TryGetValue("Other Expense Amount", out var otherExpenseAmount) && int.TryParse(otherExpenseAmount?.ToString(), out var parsedOtherExpenseAmount) ? parsedOtherExpenseAmount : 0,
                        TotalBudget = row.TryGetValue("Total Budget", out var totalBudget) && int.TryParse(totalBudget?.ToString(), out var parsedTotalBudget) ? parsedTotalBudget : 0,
                        TotalExpenseBTC = row.TryGetValue(" Total Expense BTC", out var totalExpenseBTC) && int.TryParse(totalExpenseBTC?.ToString(), out var parsedTotalExpenseBTC) ? parsedTotalExpenseBTC : 0,
                        TotalExpenseBTE = row.TryGetValue("Total Expense BTE", out var totalExpenseBTE) && int.TryParse(totalExpenseBTE?.ToString(), out var parsedTotalExpenseBTE) ? parsedTotalExpenseBTE : 0,
                        CostperparticipantHelper = row.TryGetValue("Cost per participant - Helper", out var costperparticipantHelper) && int.TryParse(costperparticipantHelper?.ToString(), out var parsedCostperparticipantHelper) ? parsedCostperparticipantHelper : 0,
                        AdvanceAmount = row.TryGetValue("Advance Amount", out var advanceAmount) && int.TryParse(advanceAmount?.ToString(), out var parsedAdvanceAmount) ? parsedAdvanceAmount : 0,
                        TotalHCPRegistrationSpend = row.TryGetValue("Total HCP Registration Spend", out var totalHCPRegistrationSpend) && int.TryParse(totalHCPRegistrationSpend?.ToString(), out var parsedTotalHCPRegistrationSpend) ? parsedTotalHCPRegistrationSpend : 0,
                        TotalHCPRegistrationAmount = row.TryGetValue("Total HCP Registration Amount", out var totalHCPRegistrationAmount) && int.TryParse(totalHCPRegistrationAmount?.ToString(), out var parsedTotalHCPRegistrationAmount) ? parsedTotalHCPRegistrationAmount : 0,
                        FacilitychargesExcludingTax = row.TryGetValue("Facility charges Excluding Tax", out var facilitychargesExcludingTax) && int.TryParse(facilitychargesExcludingTax?.ToString(), out var parsedFacilitychargesExcludingTax) ? parsedFacilitychargesExcludingTax : 0,
                        TotalFacilitychargesincludingTax = row.TryGetValue("Total Facility charges including Tax", out var totalFacilitychargesincludingTax) && int.TryParse(totalFacilitychargesincludingTax?.ToString(), out var parsedTotalFacilitychargesincludingTax) ? parsedTotalFacilitychargesincludingTax : 0,
                        AnesthetistExcludingTax = row.TryGetValue("Anesthetist Excluding Tax", out var anesthetistExcludingTax) && int.TryParse(anesthetistExcludingTax?.ToString(), out var parsedAnesthetistExcludingTax) ? parsedAnesthetistExcludingTax : 0,
                        AnesthetistincludingTax = row.TryGetValue("Anesthetist including Tax", out var anesthetistincludingTax) && int.TryParse(anesthetistincludingTax?.ToString(), out var parsedAnesthetistincludingTax) ? parsedAnesthetistincludingTax : 0,

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


        public class getIds
        {
            public List<string> EventIds { get; set; }
        }
    }
}
