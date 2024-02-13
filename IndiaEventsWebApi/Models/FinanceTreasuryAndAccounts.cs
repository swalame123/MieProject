namespace IndiaEventsWebApi.Models
{
    public class FinanceTreasuryAndAccounts
    {

    }
    public class FinanceAccounts
    {
    
        public string? Id { get; set; }
        public string? JVNumber { get; set; }
        public string? JVDate { get; set; }
    }


    public class FinanceTreasury
    {

        public string? Id { get; set; }
        public string? PVNumber { get; set; }
        public string? PVDate { get; set; }
        public string? BankReferenceNumber { get; set; }
        public string? BankReferenceDate { get; set; }
    }

}
