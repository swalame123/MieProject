using MieProject.Models.EventTypeSheets;
using MieProject.Models.RequestSheets;

namespace MieProject.Models
{
    public class EventSettlementData
    {
        public EventSettlement EventSettlement { get; set; }
        public List<EventRequestInvitees> RequestInvitees { get; set; }
        public List<EventRequestExpenseSheet> ExpenseSheet { get; set; }

    }
}
