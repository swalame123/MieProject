namespace IndiaEventsWebApi.Models.RequestSheets
{
    public class EventRequestsHcpRole
    {
        public string? EventIdorEventRequestId { get; set; }
        public string? HcpName { get; set; }
        public string? HcpRole { get; set; }
        public string? MisCode { get; set; }
        public string? SpeakerCode { get; set; }
        public string? TrainerCode { get; set; }
        public string? Speciality { get; set; }
        public string? Tier { get; set; }
        public string? GOorNGO { get; set; }
        public string? HonorariumRequired { get; set; }
        public string? HonarariumAmount { get; set; }
        public string? Travel { get; set; }
        public string? Accomdation { get; set; }
        public string? LocalConveyance { get; set; }
        public string? FinalAmount { get; set; }
        public string? Rationale { get; set; }
        public string? PresentationDuration { get; set; }
        public string? PanelSessionPreperationDuration { get; set; }
        public string? PanelDisscussionDuration { get; set; }
        public string? QASessionDuration { get; set; }
        public string? BriefingSession { get; set; }
        public string? TotalSessionHours { get; set; }
        public string? IsInclidingGst { get; set; }
        public string? AgreementAmount { get; set; }
        public string? PanCardName { get; set; }
        public string? ExpenseType { get; set; }
        public string? OthersType { get; set; }
        public string? BankAccountNumber { get; set; }
        public string? BankName { get; set; }
        public string? IFSCCode { get; set; }
        public string? Fcpadate { get; set; }

        public string? Currency { get; set; }
        public string? OtherCurrencyType { get; set; }
        public string? BeneficiaryName { get; set; }
        public string? PanNumber { get; set; }
        //Currency,Other Currency Type,Beneficiary Name,Pan Number,
        public string? LcBtcorBte { get; set; }
        public string? TravelBtcorBte { get; set; }
        public string? AccomodationBtcorBte { get; set; }

        public int? HonarariumAmountExcludingTax { get; set; }
        public int? TravelExcludingTax { get; set; }
        public int? AccomdationExcludingTax { get; set; }
        public int? LocalConveyanceExcludingTax { get; set; }
        public string? IsUpload { get; set; }
        public List<string>? FilesToUpload { get; set; }

    }
}
