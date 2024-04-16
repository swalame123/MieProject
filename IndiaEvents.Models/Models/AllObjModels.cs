using IndiaEventsWebApi.Models.EventTypeSheets;
using IndiaEventsWebApi.Models.RequestSheets;

namespace IndiaEventsWebApi.Models
{
    public class AllObjModels
    {
        public Class1? class1 { get; set; }
        public List<EventRequestBrandsList>? RequestBrandsList { get; set; }
        public List<EventRequestInvitees>? EventRequestInvitees { get; set; }
        public List<EventRequestsHcpRole>? EventRequestHcpRole { get; set; }
        public List<EventRequestHCPSlideKit>? EventRequestHCPSlideKits { get; set; }
        public List<EventRequestExpenseSheet>? EventRequestExpenseSheet { get; set; }
        // public IFormFile? formFile { get; set; }
    }
    public class UpdateAllObjModels
    {
        public Class1? Class1 { get; set; }
        public List<UpdateEventRequestBrandsList>? RequestBrandsList { get; set; }
        public List<UpdateEventRequestInvitees>? EventRequestInvitees { get; set; }
        public List<UpdateEventRequestsHcpRole>? EventRequestHcpRole { get; set; }
        public List<UpdateEventRequestHCPSlideKit>? EventRequestHCPSlideKits { get; set; }
        public List<UpdateEventRequestExpenseSheet>? EventRequestExpenseSheet { get; set; }
        // public IFormFile? formFile { get; set; }
    }
}
