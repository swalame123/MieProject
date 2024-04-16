namespace IndiaEventsWebApi.Models.RequestSheets
{
    public class EventRequestInvitees
    {
        public string? EventIdOrEventRequestId { get; set; }

        public string? MISCode { get; set; }
        public string? LocalConveyance { get; set; }
        public string? BtcorBte { get; set; }
        public string? LcAmount { get; set; }
        public int? LcAmountExcludingTax { get; set; }
        public string? InviteedFrom { get; set; }
        public string? InviteeName { get; set; }
        public string? Speciality { get; set; }
        public string? HCPType { get; set; }
        public string? Designation { get; set; }
        public string? EmployeeCode { get; set; }




    } 
    public class UpdateEventRequestInvitees
    {
        public int? InviteeId { get; set; }
        public string? EventIdOrEventRequestId { get; set; }

        public string? MISCode { get; set; }
        public string? LocalConveyance { get; set; }
        public string? BtcorBte { get; set; }
        public string? LcAmount { get; set; }
        public int? LcAmountExcludingTax { get; set; }
        public string? InviteedFrom { get; set; }
        public string? InviteeName { get; set; }
        public string? Speciality { get; set; }
        public string? HCPType { get; set; }
        public string? Designation { get; set; }
        public string? EmployeeCode { get; set; }




    }

    public class AddNewInvitee
    {
        public string? EventIdOrEventRequestId { get; set; }
        public string? MISCode { get; set; }
        public string? LocalConveyance { get; set; }
        public string? BtcorBte { get; set; }
        public int? LcAmount { get; set; }
        public int? LcAmountExcludingTax { get; set; }
        public string? InviteedFrom { get; set; }
        public string? InviteeName { get; set; }
        public string? Speciality { get; set; }
        public string? HCPType { get; set; }
        public string? Designation { get; set; }
        public string? EmployeeCode { get; set; }
        public string? EventTopic { get; set; }
        public string? EventType { get; set; }
        public string? VenueName { get; set; }
        public DateTime? EventStartDate { get; set; }
        public DateTime? EventEndDate { get; set; }
    }
}
