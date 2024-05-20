using IndiaEvents.Models.Models.RequestSheets;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IndiaEvents.Models.Models.EventTypeSheets
{
    public class Class2
    {
        public Class2Data? ClassII { get; set; }
        public List<BrandsList>? BrandsListData { get; set; }
        public List<PanelDetails>? PanelistData { get; set; }
        public List<Invitees>? InviteesData { get; set; }
        public List<SlideKitSelection>? SlideKitSelectionData { get; set; }
        public List<ExpenseSheetData>? ExpenseSheetData { get; set; }
    }
    public class Class2Data
    {
        public DateTime? EventDate { get; set; }
        public string? EventType { get; set; }
        public string? EventName { get; set; }
        //public DateTime? EventEndDate { get; set; }
        public string? StartTime { get; set; }
        public string? EndTime { get; set; }
        public string? Objective { get; set; }
        public string? City { get; set; }
        public string? State { get; set; }
        public string? ModeOfTraining { get; set; }
        public string? VenueName { get; set; }
        public string? VendorName { get; set; }

        public string? InitiatorName { get; set; }
        public double? AdvanceAmount { get; set; }
        public double? TotalExpenseBTC { get; set; }
        public double? TotalExpenseBTE { get; set; }
        public double? TotalHonorariumAmount { get; set; }
        public double? TotalTravelAccommodationAmount { get; set; }
        public double? TotalAccomodationAmount { get; set; }
        public double? TotalBudget { get; set; }
        public double? TotalLocalConveyance { get; set; }
        public double? TotalTravelAmount { get; set; }
        public double? TotalExpense { get; set; }
        public string? InitiatorEmail { get; set; }
        public string? RBMorBMEmail { get; set; }
        public string? SalesHeadEmail { get; set; }
        public string? FinanceHeadEmail { get; set; }
        public string? SalesCoordinatorEmail { get; set; }
        public string? MarketingCoordinatorEmail { get; set; }
        public string? MarketingHeadEmail { get; set; }
        public string? FinanceEmail { get; set; }
        public string? ComplianceEmail { get; set; }
        public string? FinanceAccountsEmail { get; set; }
        public string? ReportingManagerEmail { get; set; }
        public string? FirstLevelEmail { get; set; }
        public string? MedicalAffairsEmail { get; set; }
        public string? Role { get; set; }

        public string? AllIndiaDoctorsInvitedForTheEvent { get; set; }


        public List<string>? FilesUpload { get; set; }
        public string? IsDeviationUpload { get; set; }
        public int? EventOpen45dayscount { get; set; }
        public List<string>? DeviationFiles { get; set; }



    }

    public class PanelDetails
    {
        public string? MISCode { get; set; }
        public string? HcpRole { get; set; }
        public string? HCPRoleName { get; set; }
        public string? SpeakerCode { get; set; }
        public string? TrainerCode { get; set; }

        public string? HcpName { get; set; }
        //public string? HcpCode { get; set; }
        public string? HcpQualification { get; set; }
        public string? HcpCountry { get; set; }
        public string? Speciality { get; set; }
        public string? HcpCategory { get; set; }
        public string? HcpType { get; set; }
        public string? Rationale { get; set; }
        public DateTime? FCPAIssueDate { get; set; }

        public string? IsHonorariumApplicable { get; set; }
        public double? Presentation_Speaking_WorkshopDuration { get; set; }
        public double? DevelopmentofPresentationPanelSessionPreparation { get; set; }
        public double? PaneldiscussionSessionduration { get; set; }
        public double? QASession { get; set; }
        public double? Speaker_TrainerBriefing { get; set; }
        public double? TotalNoOfHours { get; set; }
        public double? HonorariumAmountexcludingTax { get; set; }
        public double? HonorariumAmountincludingTax { get; set; }
        public double? YTDspendIncludingCurrentEvent { get; set; }
        public string? ExpenseType { get; set; }
        public string? IsTravelBTC_BTE { get; set; }
        public string? IsAccomodationBTC_BTE { get; set; }
        public string? IsLCBTC_BTE { get; set; }
        public string? TravelSelection { get; set; }
        public double? TravelAmountExcludingTax { get; set; }
        public double? TravelAmountIncludingTax { get; set; }
        public double? AccomodationAmountExcludingTax { get; set; }
        public double? AccomodationAmountIncludingTax { get; set; }
        public double? LocalConveyanceAmountexcludingTax { get; set; }
        public double? LocalConveyanceAmountincludingTax { get; set; }
        public double? AgreementAmount { get; set; }
        public double? TravelandAccomodationspendincludingcurrentevent { get; set; }
        public double? FinalAmount { get; set; }
        public EventRequestBenificiaryDetails? BenificiaryDetailsData { get; set; }
        public string? IsFilesUpload { get; set; }
        public List<string>? Files { get; set; }


    }

    public class Invitees
    {

        public string? MISCode { get; set; }
        public string? IsLocalConveyance { get; set; }
        public string? BtcorBte { get; set; }
        public double? LcAmountIncludingTax { get; set; }
        public double? LcAmountExcludingTax { get; set; }
        public string? InviteedFrom { get; set; }
        public string? Invitee_Employee_HcpName { get; set; }
        public string? Speciality { get; set; }
        public string? HCPType { get; set; }
        public string? EmployeeCode { get; set; }
    }

    public class SlideKitSelection
    {
        public string? MisCode { get; set; }
        public string? HcpName { get; set; }
        public string? HcpType { get; set; }
        public string? SlideKitSelectionType { get; set; }
        public string? SlideKitDocument { get; set; }
        public string? IsUpload { get; set; }
        public List<string>? FilesToUpload { get; set; }

    }

    public class ExpenseSheetData
    {
        public string? ExpenseType { get; set; }
        public double? ExpenseAmountIncludingTax { get; set; }
        public double? ExpenseAmountExcludingTax { get; set; }
        public string? IsBtcorBte { get; set; }
    }

    public class BrandsList
    {
        public string? BrandName { get; set; }
        public string? PercentAllocation { get; set; }
        public string? ProjectId { get; set; }
    }
}
