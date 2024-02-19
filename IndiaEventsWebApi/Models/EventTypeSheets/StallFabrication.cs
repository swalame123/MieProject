using IndiaEventsWebApi.Models.RequestSheets;

namespace IndiaEventsWebApi.Models.EventTypeSheets
{
    public class StallFabrication
    {             
        public string? EventType { get; set; }
        public string? EventName { get; set; }

        public DateTime? EventDate { get; set; }
        public string? StartTime { get; set; }
        public string? EndTime { get; set; }
        public string? Class_III_EventCode { get; set; }
        public string? RBMorBM { get; set; }
        public string? Sales_Head { get; set; }
        public string? Marketing_Head { get; set; }


        public string? MedicalAffairsEmail { get; set; }
        public string? ReportingManagerEmail { get; set; }
        public string? FirstLevelEmail { get; set; }
        public string? ComplianceEmail { get; set; }
        //public string? FinanceTreasuryEmail { get; set; }
        public string? FinanceAccountsEmail { get; set; }
        public string? SalesCoordinatorEmail { get; set; }
        public string? Role { get; set; }

        public string? Finance { get; set; }
        public string? InitiatorName { get; set; }
        public string? Initiator_Email { get; set; }       
        //public string? Role { get; set; }
        public string? TotalBudgetAmount { get; set;}
        public string? EventWithin7daysUpload { get; set; }
        public string? TableContainsDataUpload { get; set; }
        public string? EventBrouchereUpload { get; set; }
        public string? Invoice_QuotationUpload { get; set; }
        public string? IsDeviationUpload { get; set; }
        public string? IsAdvanceRequired { get; set; }

    }

    public class AllStallFabrication
    {
        public StallFabrication? StallFabrication { get; set; }
        public List<EventRequestBrandsList>? EventBrands { get; set; }
        public List<EventRequestExpenseSheet>? ExpenseSheets { get; set; }
    }

}
