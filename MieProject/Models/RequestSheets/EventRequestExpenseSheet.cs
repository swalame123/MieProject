namespace MieProject.Models.RequestSheets
{
    public class EventRequestExpenseSheet
    {
        public string EventId { get; set; }
        public string Expense { get; set; }
        public string Amount { get; set; }
        public string AmountExcludingTax { get; set; }
        public string BtcorBte{ get; set; }
        public string BtcAmount { get; set;}
        public string BteAmount { get; set;}
        public string BudgetAmount { get; set;}
    }
}
