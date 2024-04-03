using IndiaEventsWebApi.Models.EventTypeSheets;
using IndiaEventsWebApi.Models.RequestSheets;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IndiaEvents.Models.Models.Draft
{
    public class DraftData
    {
        public Draft? Draft { get; set; }
        public List<EventRequestBrandsList>? RequestBrandsList { get; set; }
        public List<EventRequestInvitees>? EventRequestInvitees { get; set; }
        public List<EventRequestsHcpRole>? EventRequestHcpRole { get; set; }
        public List<EventRequestHCPSlideKit>? EventRequestHCPSlideKits { get; set; }
        public List<EventRequestExpenseSheet>? EventRequestExpenseSheet { get; set; }
    }

    public class PostDraftData
    {
        public DateTime? EventDate { get; set; }
        public string? InitiatorName { get; set; }
        public string? InitiatorEmail { get; set; }
        public string? Role { get; set; }

        public string? Isfile { get; set; }
        public List<string>? UploadFiles { get; set; }
    }
    public class Draft
    {
        public string? DraftId { get; set; }
        public string? EventTopic { get; set; }
        public string? EventType { get; set; }
        public DateTime? EventDate { get; set; }
        public string? StartTime { get; set; }
        public string? EndTime { get; set; }
        public DateTime? EventEndDate { get; set; }
        public string? VenueName { get; set; }
        public string? City { get; set; }
        public string? State { get; set; }       
        public string? HCPRole { get; set; }
        public string? IsAdvanceRequired { get; set; }
       // public string? TotalInvitees { get; set; }        
        public string? AdvanceAmount { get; set; }        
        public string? ClassIIIEventCode { get; set; }
        public string? MeetingType { get; set; }
        public string? SponcershipSocietyname { get; set; }
        public string? VenueCountry { get; set; }
        public string? MedicalUtilityType { get; set; }
        public string? MedicalUtilityDescription { get; set; }
        public DateTime? ValidFrom { get; set; }
        public DateTime? ValidTo { get; set; }
        //public string? Role { get; set; }        
        //public string? InitiatorName { get; set; }
        //public string? Initiator_Email { get; set; }
        public string? IsFiles { get; set; }
        public string? IsBrands { get; set; }
        public string? IsPanelists { get; set; }
        public string? IsSlideKits { get; set; }
        public string? IsInvitees { get; set; }
        public string? IsExpense { get; set; }
        public List<string>? Files { get; set; }        
    }
}
