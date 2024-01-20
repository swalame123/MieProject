using MieProject.Models.EventTypeSheets;
using MieProject.Models.RequestSheets;

namespace MieProject.Models
{
    public class AllObjModels
    {
        public Class1? class1 { get; set; }
        public List<EventRequestBrandsList>? RequestBrandsList { get; set; }       
       public List<EventRequestInvitees>? EventRequestInvitees { get; set; }
        public List<EventRequestsHcpRole>? EventRequestHcpRole { get; set; }
        public List<EventRequestHCPSlideKit>? EventRequestHCPSlideKits { get; set; }
        public List<EventRequestExpenseSheet>? EventRequestExpenseSheet { get; set; }

    }
}
