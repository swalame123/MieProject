namespace IndiaEventsWebApi.Models
{
    public class FinanceTreasuryAndAccounts
    {

    }
   

    public class FinanceAccounts
    {
        public string? Id { get; set; }
        public string? HCPName { get; set; }
        public string? MISCode { get; set; }
        public string? JVNumber { get; set; }
        public DateTime? JVDate { get; set; }
    }
    public class FinanceAccountsUpdate
    {
        public string? EventId { get; set; }
        public string? Status { get; set; }
        public string? Description { get; set; }
        public List<FinanceAccounts> FinanceAccounts { get; set; }
    }
    public class FinanceAccountsUpdateIn3Sheets
    {
        public string? EventId { get; set; }
        public string? Status { get; set; }
        public string? Description { get; set; }
        public List<FinanceAccounts>? PanelSheet { get; set; }
        public List<FinanceAccounts>? ExpenseSheet { get; set; }
        public List<FinanceAccounts>? InviteesSheet { get; set; }
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

    public class FinanceTreasuryUpdate
    {
        public string? EventId { get; set; }
        public string? Status { get; set; }
        public string? Description { get; set; }

        public List<FinanceTreasury> FinanceTreasury { get; set; }
    }

    public class FinanceTreasuryUpdateIn3Sheets
    {
        public string? EventId { get; set; }
        public string? Status { get; set; }
        public string? Description { get; set; }

        public List<FinanceTreasuryForPanel>? PanelSheet { get; set; }
        public List<FinanceTreasury>? ExpenseSheet { get; set; }
        public List<FinanceTreasury>? InviteesSheet { get; set; }
    }
    public class PanelDataInFinance
    {
        public PVNumberAndPVDate Travel { get; set; }
        public PVNumberAndPVDate Accomodation { get; set; }
        public PVNumberAndPVDate LocalConveyance { get; set; }

    }

    public class FinanceTreasuryForPanel
    {

        public string? Id { get; set; }
        public string? HCPName { get; set; }
        public string? MISCode { get; set; }     
        public PanelDataInFinance PanelDataInFinance { get; set; }
    }
    public class PVNumberAndPVDate
    {
        public string? PVNumber { get; set; }
        public DateTime? PVDate { get; set; }
        public string? BankReferenceNumber { get; set; }
        public DateTime? BankReferenceDate { get; set; }
    }
}
