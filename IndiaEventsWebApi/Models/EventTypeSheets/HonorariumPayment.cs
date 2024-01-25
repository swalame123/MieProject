using IndiaEventsWebApi.Models.RequestSheets;

namespace IndiaEventsWebApi.Models.EventTypeSheets
{
    public class HonorariumPayment
    {
        public string HCPName { get; set; }
        public string EventId { get; set; }
        public string EventType { get; set; }
        public string HCPRole { get; set; }
        public string MISCode { get; set; }
        public string GONGO { get; set; }
        public string IsItincludingGST { get; set; }
        public string AgreementAmount { get; set; }
    }
    public class HonorariumPaymentList
    {
        public List<HonorariumPayment>? RequestHonorariumList { get; set; }
    }
}
