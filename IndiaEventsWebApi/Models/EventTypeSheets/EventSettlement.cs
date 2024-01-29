namespace IndiaEventsWebApi.Models.EventTypeSheets
{
    public class EventSettlement
    {
        public List<Invitee> Invitee { get; set; }
        public List<ExpenseSheet> expenseSheets { get; set; }
       
        public string? EventId { get; set; }
        public string? EventTopic { get; set; }
        public string? EventType { get; set; }
        public DateTime? EventDate { get; set; }
        public string? StartTime { get; set; }
        public string? InitiatorName { get; set; }
        public string? EndTime { get; set; }
        public string? VenueName { get; set; }
        public string? State { get; set; }
        public string? City { get; set; }
        public string? Attended { get; set; }
        //public string? InviteesParticipated { get; set; }
        //public string? ExpenseParticipated { get; set; }
        public string? Brands { get; set; }
        public string? Panalists { get; set; }
        public string? SlideKits { get; set; }
        public string? TotalExpense { get; set; }
        public string? Advance { get; set; }
        public string? InitiatorEmail { get; set; }
        public string RBMorBM { get; set; }
        public string SalesHead { get; set; }
        public string MarkeringHead { get; set; }
        public string Compliance { get; set; }
        public string FinanceAccounts { get; set; }
        public string FinanceTreasury { get; set; }
        public string MedicalAffairsHead { get; set; }
        public string FinanceHead { get; set; }
        public string EventOpen30Days { get; set; }
        public string Event7Dayd { get; set; }
        public string PostEventSubmitted { get; set; }
        public string IsAdvanceRequired { get; set; }
        public string totalInvitees { get; set; }
        public string TotalAttendees { get; set; }
        //public string FinanceHead { get; set; }
        


    }
    //public class Panalists
    //{
    //    public string HcpRole { get; set; }
    //    public string HcpName { get; set; }
    //    public string HonarariumAmount { get; set; }
    //    public string TravelAmount { get; set; }
    //    public string AccomdationAmount { get; set; }

    //}
    public class Invitee
    {

        public string InviteeName { get; set; }
        public string MISCode { get; set; }
        public string LocalConveyance { get; set; }
        public string BtcorBte { get; set; }
        public string LcAmount { get; set; }
    }
    public class ExpenseSheet
    {

        public string Expense { get; set; }
        public string Amount { get; set; }
        public string AmountExcludingTax { get; set; }
        public string BtcorBte { get; set; }

    }
    //public class Branddetail
    //{
    //    public string BrandName { get; set; }
    //    public string PercentAllocation { get; set; }
    //    public string ProjectId { get; set; }
    //}
    //public class HCPSlideKits
    //{

    //    public string MIS { get; set; }
    //    public string SlideKitType { get; set; }
    //    public string SlideKitDocument { get; set; }


    //}
}
