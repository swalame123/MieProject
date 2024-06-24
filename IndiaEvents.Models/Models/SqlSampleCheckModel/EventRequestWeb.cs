using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IndiaEvents.Models.Models.SqlSampleCheckModel
{
    public class EventRequestWeb
    {
        public int Id { get; set; }

        [Column("EventId/EventRequestId")]
        public string? EventIdEventRequestId { get; set; }

        [Column("Event Topic")]
        public string? EventTopic { get; set; }

        [Column("Event Type")]
        public string? EventType { get; set; }

        [Column("Event Date")]
        public string? EventDate { get; set; }

        [Column("Start Time")]
        public string? StartTime { get; set; }

        [Column("End Time")]
        public string? EndTime { get; set; }

        [Column("End Date")]
        public DateTime EndDate { get; set; }

        [Column("Venue Name")]
        public string? VenueName { get; set; }

        [Column("Initiator Name")]
        public string? InitiatorName { get; set; }

        [Column("Initiator Email")]
        public string? InitiatorEmail { get; set; }

        [Column("City")]
        public string? City { get; set; }

        [Column("State")]
        public string? State { get; set; }

        [Column("Product Brand")]
        public string? ProductBrand { get; set; }

        [Column("Selected Products")]
        public string? SelectedProducts { get; set; }

        [Column("Beneficiary Details")]
        public string? BeneficiaryDetails { get; set; }

        [Column("Brands")]
        public string? Brands { get; set; }

        [Column("Panelists")]
        public string? Panelists { get; set; }

        [Column("SlideKits")]
        public string? SlideKits { get; set; }

        [Column("Invitees")]
        public string? Invitees { get; set; }

        [Column("MIPL Invitees")]
        public string? MIPLInvitees { get; set; }

        [Column("Expenses")]
        public string? Expenses { get; set; }

        [Column("HCP Role")]
        public string? HCPRole { get; set; }

        [Column("IsAdvanceRequired")]
        public string? IsAdvanceRequired { get; set; }

        [Column("EventOpen30days")]
        public string? EventOpen30days { get; set; }

        [Column("EventWithin7days")]
        public string? EventWithin7days { get; set; }

        [Column("Created On")]
        public DateTime CreatedOn { get; set; }

        [Column("Modified")]
        public DateTime Modified { get; set; }

        [Column("RBM/BM")]
        public string? RBMBM { get; set; }

        [Column("Sales Head")]
        public string? SalesHead { get; set; }

        [Column("Marketing Head")]
        public string? MarketingHead { get; set; }

        [Column("Finance Accounts")]
        public string? FinanceAccounts { get; set; }

        [Column("Finance Treasury")]
        public string? FinanceTreasury { get; set; }

        [Column("Compliance")]
        public string? Compliance { get; set; }

        [Column("1 Up Manager")]
        public string? _1UpManager { get; set; }

        [Column("Reporting Manager")]
        public string? ReportingManager { get; set; }

        [Column("Medical Affairs Head")]
        public string? MedicalAffairsHead { get; set; }

        [Column("Sales Coordinator")]
        public string? SalesCoordinator { get; set; }

        [Column("Marketing Coordinator")]
        public string? MarketingCoordinator { get; set; }

        [Column("Total Invitees")]
        public string? TotalInvitees { get; set; }

        [Column("Total HCP Registration Amount")]
        public string? TotalHCPRegistrationAmount { get; set; }

        [Column("Total Honorarium Amount")]
        public double TotalHonorariumAmount { get; set; }

        [Column("Total Travel Amount")]
        public int TotalTravelAmount { get; set; }

        [Column("Total Accommodation Amount")]
        public int TotalAccommodationAmount { get; set; }

        [Column("Total Travel & Accommodation Amount")]
        public int TotalTravelAccommodationAmount { get; set; }

        [Column("Total Local Conveyance")]
        public int TotalLocalConveyance { get; set; }

        [Column("Total Expense")]
        public int TotalExpense { get; set; }

        [Column("BTE Expense Details")]
        public string? BTEExpenseDetails { get; set; }

        [Column("Budget Amount")]
        public int BudgetAmount { get; set; }

        [Column("Total Expense BTC")]
        public int TotalExpenseBTC { get; set; }

        [Column("Total Expense BTE")]
        public int TotalExpenseBTE { get; set; }

        [Column("Advance Amount")]
        public int AdvanceAmount { get; set; }
        [Column("Role")]
        public string? Role { get; set; }

        [Column("Class III Event Code")]
        public string? ClassIIIEventCode { get; set; }

       

        [Column("Finance Treasury URL")]
        public string? FinanceTreasuryURL { get; set; }

        [Column("Approver Pre Event URL")]
        public string? ApproverPreEventURL { get; set; }

        [Column("Initiator URL")]
        public string? InitiatorURL { get; set; }

    }
}
