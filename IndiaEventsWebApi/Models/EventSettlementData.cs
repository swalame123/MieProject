using IndiaEventsWebApi.Models.EventTypeSheets;
using IndiaEventsWebApi.Models.RequestSheets;

namespace IndiaEventsWebApi.Models
{
    public class EventSettlementData
    {
        public EventSettlement EventSettlement { get; set; }
        public List<EventRequestInvitees> RequestInvitees { get; set; }
        public List<EventRequestExpenseSheet> ExpenseSheet { get; set; }

    }
}
