namespace IndiaEventsWebApi.Models
{
    public class FinanceTreasuryAndAccounts
    {

    }
    public class FinanceAccounts
    {
    
        public string? Id { get; set; }
        public string? JVNumber { get; set; }
        public DateTime? JVDate { get; set; }
    }


    public class FinanceTreasury
    {

        public string? Id { get; set; }
        public string? HCPName { get; set; }
        public string? MISCode { get; set; }

        public string? PVNumber { get; set; }
        public DateTime? PVDate { get; set; }
        public string? BankReferenceNumber { get; set; }
        public DateTime? BankReferenceDate { get; set; }
    }

}
