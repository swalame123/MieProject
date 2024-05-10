using IndiaEventsWebApi.Models.EventTypeSheets;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IndiaEvents.Models.Models.EventTypeSheets
{
    public class HonorariumPaymentP2
    {
        public string? EventId { get; set; }
        public string? EventType { get; set; }
        public string? EventTopic { get; set; }
        public string? HonarariumSubmitted { get; set; }
        public DateTime? EventDate { get; set; }
        public string? City { get; set; }
        public string? State { get; set; }
        public string? StartTime { get; set; }
        public string? EndTime { get; set; }
        public string? VenueName { get; set; }
        public double? TotalTravelAndAccomodationSpend { get; set; }
        public double? TotalHonorariumSpend { get; set; }
        public double? TotalSpend { get; set; }
        public double? TotalLocalConveyance { get; set; }
        public string? Brands { get; set; }
        public string? Invitees { get; set; }
        public string? Panelists { get; set; }
        public string? InitiatorName { get; set; }
        public string? InitiatorEmail { get; set; }
        public string? RBMorBM { get; set; }
        public string? Compliance { get; set; }
        public string? FinanceHead { get; set; }
        public string? MedicalAffairsEmail { get; set; }
        public string? ReportingManagerEmail { get; set; }
        public string? FirstLevelEmail { get; set; }
        public string? SalesCoordinatorEmail { get; set; }
        public string? MarketingHeadEmail { get; set; }
        public string? SalesHeadEmail { get; set; }
        public string? Role { get; set; }
        public string? FinanceAccounts { get; set; }
        public string? FinanceTreasury { get; set; }
        public string? SlideKits { get; set; }
        public string? Expenses { get; set; }
        public double? TotalTravelSpend { get; set; }
        public double? TotalAccomodationSpend { get; set; }
        public double? TotalExpenses { get; set; }
        public string? IsDeviationUpload { get; set; }



        public List<string>? Files { get; set; }
        public List<string>? DeviationFiles { get; set; }

    }
    public class HCPDetailsP2
    {
        public string? HcpName { get; set; }
        public string? HcpRole { get; set; }
        public string? MisCode { get; set; }
        public string? GOorNGO { get; set; }
        public string? IsInclidingGst { get; set; }
        public double? AgreementAmount { get; set; }
        public string? IsAnnualTrainerAgreementValid { get; set; }

    }

    public class HonorariumPaymentListPh2
    {
        public HonorariumPaymentP2? RequestHonorariumList { get; set; }
        public List<HCPDetailsP2>? HcpRoles { get; set; }
    }
    public class HonorariumUpdate
    {
        public string? PanelId { get; set; }
        public string? EventId { get; set; }
        public string? IsInclidingGst { get; set; }
        public List<string>? FilesToUpload { get; set; }
        public string? IsAnnualTrainerAgreementValid { get; set; }
    }
}
