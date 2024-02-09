using IndiaEventsWebApi.Models.RequestSheets;

namespace IndiaEventsWebApi.Models.EventTypeSheets
{
    public class MedicalUtilityData
    {
        public string? EventTopic { get; set; }
        public string? EventType { get; set; }
        public DateTime? EventDate { get; set; }
        public DateTime? ValidFrom { get; set; }
        public DateTime? ValidTill { get; set; }
        public string? MedicalUtilityType { get; set; }
        public string? MedicalUtilityDescription { get; set; }
        public string? Class_III_EventCode { get; set; }
        public string? IsAdvanceRequired { get; set; }
        public string? EventOpen30daysFile { get; set; }
        public string? EventWithin7daysFile { get; set; }
        // public string? FcpaFile { get; set; }
        public string? IsDeviationUpload { get; set; }
        public string? TotalBudgetAmount { get; set; }
        public string? RBMorBM { get; set; }
        public string? Sales_Head { get; set; }
        public string? Marketing_Head { get; set; }
        public string? Finance { get; set; }
        public string? InitiatorName { get; set; }
        public string? Role { get; set; }
        public string? Initiator_Email { get; set; }
        // public List<string>? AggregateDeviationFiles { get; set; }
        public string? UploadDeviationFile { get; set; }
    }

    public class HCPListData
    {
        public string? HcpName { get; set; }
        public string? MisCode { get; set; }
        public string? Speciality { get; set; }
        public string? Tier { get; set; }
        public string? HcpType { get; set; }
        public string? Rationale { get; set; }
        public string? UploadFCPA { get; set; }
        public string? UploadWrittenRequestDate { get; set; }
        public string? UploadHCPRequestDate { get; set; }
        public string? Invoice_Brouchere_Quotation { get; set; }
        public string? MedicalUtilityCostAmount { get; set; }
        public string? Legitimate { get; set; }
        public string? Objective { get; set; }
        public DateTime? Fcpadate { get; set; }
    }

    public class ExpenseListData
    {
        public string? MisCode { get;set; }

        public string? Expense { get; set; }
        public string? BTC_BTE { get; set; }
        public string? TotalExpenseAmount { get; set; }

    }


    public class MedicalUtilityPreEventPayload
    {
        public MedicalUtilityData? MedicalUtilityData { get; set; }
        public List<EventRequestBrandsList>? BrandsList { get; set; }
        public List<ExpenseListData>? ExpenseSheet { get; set; }
        public List<HCPListData>? HcpList { get; set; }
    }


}
