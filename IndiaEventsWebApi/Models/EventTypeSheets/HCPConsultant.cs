using IndiaEventsWebApi.Models.RequestSheets;

namespace IndiaEventsWebApi.Models.EventTypeSheets
{
    public class HCPConsultant
    {
        public string? EventTopic { get; set; }
        public string? EventType { get; set; }
        public DateTime? EventDate { get; set; }
        public DateTime? EventEndDate { get; set; }
        public string? StartTime { get; set; }
        public string? EndTime { get; set; }
        public string? VenueName { get; set; }
        public string? SponsorshipSocietyName { get; set; }
        public string? Country { get; set; }
        public string? IsAdvanceRequired { get; set; }
        public string? EventOpen30days { get; set; }
        public string? EventWithin7days { get; set; }
        public string? FcpaFile { get; set; }
        public string? BrochureFile { get; set; }
        public string? AggregateDeviation { get; set; }
        public string? RBMorBM { get; set; }
        public string? Sales_Head { get; set; }
        public string? Marketing_Head { get; set; }
        //public string? MedicalAffairsEmail { get; set; }
        public string? SalesCoordinatorEmail { get; set; }

        public string? MedicalAffairsEmail { get; set; }
        public string? ReportingManagerEmail { get; set; }
        public string? FirstLevelEmail { get; set; }
        public string? ComplianceEmail { get; set; }
        //public string? FinanceTreasuryEmail { get; set; }
        public string? FinanceAccountsEmail { get; set; }
        public string? Role { get; set; }

        public string? Finance { get; set; }
        public string? InitiatorName { get; set; }
        //public string? Role { get; set; }
        public string? Initiator_Email { get; set; }
        public List<string>? AggregateDeviationFiles { get; set; }

    }

    public class HCPList
    {
        public string? HcpName { get; set; }
        public string? MisCode { get; set; }
        public string? HcpType { get; set; }
        public string? TravelAmount { get; set; }
        public string? AccomAmount { get; set; }
        public string? LcAmount { get; set; }
        public string? RegistrationAmount { get; set; }
        public string? BudgetAmount { get; set; }
        public string? Legitimate { get; set; }
        public string? Objective { get; set; }
        public string? Rationale { get; set; }
        public DateTime? Fcpadate { get; set; }
     
       

    }
    public class ExpenseList
    {
        public string? Expense { get; set; }
        public string? BTC_BTE { get; set; }
        public string? RegstAmount { get; set; }
       



    }


    public class HCPConsultantPayload
    {
        public HCPConsultant HcpConsultant { get; set; }
        public List<EventRequestBrandsList>? BrandsList { get; set; }
        public List<ExpenseList>? ExpenseSheet { get; set; }
        public List<HCPList>? HcpList { get; set; }
    }


    //public class HCPPayload
    //{
    //    //public Class1? class1 { get; set; }
    //    public HCPConsultant HcpConsultant { get; set; }
    //    public List<EventRequestBrandsList>? RequestBrandsList { get; set; }
    //    public List<EventRequestInvitees>? EventRequestInvitees { get; set; }
    //    public List<EventRequestsHcpRole>? EventRequestHcpRole { get; set; }
    //    public List<EventRequestHCPSlideKit>? EventRequestHCPSlideKits { get; set; }
    //    public List<EventRequestExpenseSheet>? EventRequestExpenseSheet { get; set; }
    //    // public IFormFile? formFile { get; set; }
    //}
    public class HCPfollow_upsheet
    {
        public string? HCPName { get; set; }
        public string? MisCode { get; set; }
        public string? GO_NGO { get; set; }
        public string? Country { get; set; }
        public string? How_many_days_since_the_parent_event_completes { get; set; }
        public string? Follow_up_Event { get; set; }
        public DateTime? Follow_up_Event_Date { get; set; }
        public DateTime? Event_Date { get; set; }
        public string? AgreementFile { get; set; }
    }
}
