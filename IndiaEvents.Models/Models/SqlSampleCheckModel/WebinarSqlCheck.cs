using IndiaEvents.Models.Models.RequestSheets;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IndiaEvents.Models.Models.SqlSampleCheckModel
{
    public class WebinarSqlCheck
    {
        public class WebinarPayloadForSqlCheck
        {
            public int Id { get; set; }
            public WebinarSql? Webinar { get; set; }
            public List<EventRequestBrandsListSql>? RequestBrandsList { get; set; }
            public List<EventRequestInviteesSql>? EventRequestInvitees { get; set; }
            public List<EventRequestsHcpRoleSql>? EventRequestHcpRole { get; set; }
            public List<EventRequestHCPSlideKitSql>? EventRequestHCPSlideKits { get; set; }
            public List<EventRequestExpenseSheetSql>? EventRequestExpenseSheet { get; set; }
            // public IFormFile? formFile { get; set; }
        }

        public class WebinarSql
        {
            public int Id { get; set; }
            public string? EventId { get; set; }
            public string? EventTopic { get; set; }
            public string? EventType { get; set; }
            public DateTime? EventDate { get; set; }
            public DateTime? EventEndDate { get; set; }
            public string? StartTime { get; set; }
            public string? EndTime { get; set; }
            public string? Meeting_Type { get; set; }
            public string? BTEExpenseDetails { get; set; }
            //public string? VenueName { get; set; }
            //public string? City { get; set; }
            //public string? State { get; set; }
            public string? BrandName { get; set; }
            public string? PercentAllocation { get; set; }
            public string? ProjectId { get; set; }
            public string? HCPRole { get; set; }
            public string? IsAdvanceRequired { get; set; }
            public string? AdvanceAmount { get; set; }
            public string? TotalExpenseBTC { get; set; }
            public string? TotalExpenseBTE { get; set; }
            public string? EventOpen30days { get; set; }
            public string? EventOpen30dayscount { get; set; }
            public string? EventWithin7days { get; set; }
            public string? FB_Expense_Excluding_Tax { get; set; }
            public string? RBMorBM { get; set; }
            public string? FinanceHead { get; set; }


            //public string? MedicalAffairsEmail { get; set; }
            public string? MedicalAffairsEmail { get; set; }
            public string? ReportingManagerEmail { get; set; }
            public string? FirstLevelEmail { get; set; }
            public string? ComplianceEmail { get; set; }
            //public string? FinanceTreasuryEmail { get; set; }
            public string? FinanceAccountsEmail { get; set; }
            public string? SalesCoordinatorEmail { get; set; }
            public string? MarketingCoordinatorEmail { get; set; }
            public string? Role { get; set; }


            public string? Sales_Head { get; set; }
            public string? Marketing_Head { get; set; }
            public string? Finance { get; set; }
            public string? InitiatorName { get; set; }
            public string? Initiator_Email { get; set; }
            public List<string>? Files { get; set; }
            public string? IsDeviationUpload { get; set; }
            public List<EventRequestDeviationsDataSql>? DeviationDetails { get; set; }
            //public List<string>? DeviationFiles { get; set; }
            //public string? Role { get; set; }
        }
        public class EventRequestBrandsListSql
        {
            public int Id { get; set; }
            public string? BrandName { get; set; }
            public string? PercentAllocation { get; set; }
            public string? ProjectId { get; set; }
        }
        public class EventRequestInviteesSql
        {
            public int Id { get; set; }
            public string? EventIdOrEventRequestId { get; set; }

            public string? MISCode { get; set; }
            public string? LocalConveyance { get; set; }
            public string? BtcorBte { get; set; }
            public string? LcAmount { get; set; }
            public double? LcAmountExcludingTax { get; set; }
            public string? InviteedFrom { get; set; }
            public string? InviteeName { get; set; }
            public string? Speciality { get; set; }
            public string? HCPType { get; set; }
            public string? Designation { get; set; }
            public string? EmployeeCode { get; set; }
        }

        public class EventRequestsHcpRoleSql
        {
            public int Id { get; set; }
            public string? EventIdorEventRequestId { get; set; }
            public string? HcpName { get; set; }
            public string? HcpRole { get; set; }
            public string? MisCode { get; set; }
            public string? SpeakerCode { get; set; }
            public string? TrainerCode { get; set; }
            public string? Speciality { get; set; }
            public string? Tier { get; set; }
            public string? GOorNGO { get; set; }
            public string? HonorariumRequired { get; set; }
            public string? HonarariumAmount { get; set; }
            public string? Travel { get; set; }
            public string? TravelSelection { get; set; }
            public string? Accomdation { get; set; }
            public string? LocalConveyance { get; set; }
            public string? FinalAmount { get; set; }
            public string? Rationale { get; set; }
            public string? PresentationDuration { get; set; }
            public string? PanelSessionPreperationDuration { get; set; }
            public string? PanelDisscussionDuration { get; set; }
            public string? QASessionDuration { get; set; }
            public string? BriefingSession { get; set; }
            public string? TotalSessionHours { get; set; }
            public string? IsInclidingGst { get; set; }
            public string? AgreementAmount { get; set; }
            public string? PanCardName { get; set; }
            public string? ExpenseType { get; set; }
            public string? OthersType { get; set; }
            public string? BankAccountNumber { get; set; }
            public string? BankName { get; set; }
            public string? IFSCCode { get; set; }
            public string? Fcpadate { get; set; }

            public string? Currency { get; set; }
            public string? OtherCurrencyType { get; set; }
            public string? BeneficiaryName { get; set; }
            public string? PanNumber { get; set; }
            //Currency,Other Currency Type,Beneficiary Name,Pan Number,
            public string? LcBtcorBte { get; set; }
            public string? TravelBtcorBte { get; set; }
            public string? AccomodationBtcorBte { get; set; }

            public double? HonarariumAmountExcludingTax { get; set; }
            public double? TravelExcludingTax { get; set; }
            public double? AccomdationExcludingTax { get; set; }
            public double? LocalConveyanceExcludingTax { get; set; }
            public string? IsUpload { get; set; }
            public List<string>? FilesToUpload { get; set; }

        }

        public class EventRequestHCPSlideKitSql
        {
            public int Id { get; set; }
            public string EventId { get; set; }
            public string MIS { get; set; }
            public string HcpName { get; set; }
            public string SlideKitType { get; set; }
            public string SlideKitDocument { get; set; }
            public string? IsUpload { get; set; }
            public List<string>? FilesToUpload { get; set; }


        }

        public class EventRequestExpenseSheetSql
        {
            public int Id { get; set; }
            public string? EventId { get; set; }
            public string? Expense { get; set; }
            public string? Amount { get; set; }
            public string? AmountExcludingTax { get; set; }
            public double? ExcludingTaxAmount { get; set; }
            public string? BtcorBte { get; set; }
            public string? BtcAmount { get; set; }
            public string? BteAmount { get; set; }
            public string? BudgetAmount { get; set; }
        }

        public class EventRequestDeviationsDataSql
        {
            public int Id { get; set; }
            public double HonorariumAmountExcludingTax { get; set; }
            //public int AccommodationAmountExcludingTax { get; set; }
            public double TravelorAccomodationAmountExcludingTax { get; set; }
            public double OtherExpenseAmountExcludingTax { get; set; }
            //public int FoodAndBeveragesAmountExcludingTax { get; set; }
            //public int AgregateSpendAmountExcludingTax { get; set; }
            public string? DeviationFile { get; set; }
            public string? HcpName { get; set; }
            public string? MisCode { get; set; }
        }

    }
}
