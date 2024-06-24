namespace IndiaEventsWebApi.Junk.Test
{

    public class getIds
    {
        public List<string> EventIds { get; set; }
    }

    public class InvoiceIds
    {
        public List<string> ExpenseId { get; set; }
        public List<string> PanelistId { get; set; }
        
    }
    public class InvoiceIdsFromExpensePanelAndInvitee
    {
        public List<string> ExpenseId { get; set; }
        public List<string> PanelistId { get; set; }
        public List<string> InviteeId { get; set; }

    }
}
