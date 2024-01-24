namespace MieProject.Models.EventTypeSheets
{
    public class EventSettlement
    {
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
        public string? InviteesParticipated { get; set; }
        public string? ExpenseParticipated { get; set; }
        public string? TotalExpense { get; set; }
        public string? Advance { get; set; }

    }
}
