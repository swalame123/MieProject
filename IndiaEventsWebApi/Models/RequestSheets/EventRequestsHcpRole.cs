namespace IndiaEventsWebApi.Models.RequestSheets
{
    public class EventRequestsHcpRole
    {
        public string EventIdorEventRequestId { get; set; }
        public string HcpRole { get; set; }
        public string MisCode { get; set; }
        public string SpeakerCode { get; set; }
        public string TrainerCode { get; set; }
        public string Speciality { get; set; }
        public string Tier { get; set; }
        public string GOorNGO { get; set; }
        public string HonorariumRequired { get; set; }
        public string HonarariumAmount { get; set; }
        public string Travel { get; set; }
        public string Accomdation { get; set; }
        public string LocalConveyance { get; set; }
        public string FinalAmount { get; set; }
        public string Rationale { get; set; }
        public string PresentationDuration { get; set; }
        public string PanelSessionPreperationDuration { get; set; }
        public string PanelDisscussionDuration { get; set; }
        public string QASessionDuration { get; set; }
        public string BriefingSession { get; set; }
        public string TotalSessionHours { get; set; }
        public string HcpName { get; set; }
        public string? IsInclidingGst { get; set; }
        public string? AgreementAmount { get; set; }

    }
}
