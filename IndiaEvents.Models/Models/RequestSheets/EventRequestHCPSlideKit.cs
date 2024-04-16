namespace IndiaEventsWebApi.Models.RequestSheets
{
    public class EventRequestHCPSlideKit
    {
        public string EventId { get; set; }
        public string MIS { get; set; }
        public string HcpName { get; set; }
        public string SlideKitType { get; set; }
        public string SlideKitDocument { get; set; }
        public string? IsUpload { get; set; }
        public List<string>? FilesToUpload { get; set; }


    }
    public class UpdateEventRequestHCPSlideKit
    {
        public string SlideKitId { get; set; }
        public string EventId { get; set; }
        public string MIS { get; set; }
        public string HcpName { get; set; }
        public string SlideKitType { get; set; }
        public string SlideKitDocument { get; set; }
        public string? IsUpload { get; set; }
        public List<string>? FilesToUpload { get; set; }


    }
}
