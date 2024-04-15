using IndiaEvents.Models.Models.RequestSheets;
using IndiaEventsWebApi.Models.RequestSheets;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IndiaEvents.Models.Models.EventTypeSheets
{
    public class HandsOnTrainingPreEvet
    {
        public HandsOnTraining? HandsOnTraining { get; set; }
        public List<EventRequestBrandsList>? EventBrandsList { get; set; }
        //public List<ProductSelection>? ProductSelection { get; set; }
        public List<TrainerDetails>? TrainerDetails { get; set; }
        public List<SlideKitSelectionData>? SlideKitSelectionData { get; set; }
        public List<ProductSelection>? ProductSelections { get; set; }
        public List<AttenderSelection>? AttenderSelections { get; set; }
        public List<Expense>? ExpenseData { get; set; }
    }




    public class HandsOnTraining
    {
        public DateTime? EventDate { get; set; }
        public string? EventType { get; set; }
        public string? EventName { get; set; }
        public string? EventStartTime { get; set; }
        public string? EventEndTime { get; set; }
        public string? ProductBrandSelection { get; set; }
        public string? ModeOfTraining { get; set; }
        public string? WebinarType { get; set; }
        public string? VenueName { get; set; }
        public string? State { get; set; }
        public string? City { get; set; }
        public string? VendorName { get; set; }
        public string? VenueSelectionCheckList { get; set; }
        public string? IsVenueHasAnyEmergancySupport { get; set; }
        public string? EmerganctContact { get; set; }
        public string? IsVenueFacilityCharges { get; set; }
        public string? VenueFacilityChargesBtc_Bte { get; set; }
        public int? FacilityChargesExcludingTax { get; set; }
        public int? FacilityChargesIncludingTax { get; set; }
        public string? IsAnesthetistRequired { get; set; }
        public string? AnesthetistRequiredBtc_Bte { get; set; }
        public int? AnesthetistChargesExcludingTax { get; set; }
        public int? AnesthetistChargesIncludingTax { get; set; }

        public string? InitiatorName { get; set; }
        public int? AdvanceAmount { get; set; }
        public int? TotalExpenseBTC { get; set; }
        public int? TotalExpenseBTE { get; set; }
        public int? TotalHonorariumAmount { get; set; }
        public int? TotalTravelAccommodationAmount { get; set; }
        public int? TotalAccomodationAmount { get; set; }
        public int? TotalBudget { get; set; }
        public int? TotalLocalConveyance { get; set; }
        public int? TotalTravelAmount { get; set; }
        public int? TotalExpense { get; set; }
        public string? InitiatorEmail { get; set; }
        public string? RBMorBMEmail { get; set; }
        public string? SalesHeadEmail { get; set; }
        public string? SalesCoordinatorEmail { get; set; }
        public string? MarketingHeadEmail { get; set; }
        public string? FinanceEmail { get; set; }
        public string? ComplianceEmail { get; set; }
        public string? FinanceAccountsEmail { get; set; }
        public string? ReportingManagerEmail { get; set; }
        public string? FirstLevelEmail { get; set; }
        public string? MedicalAffairsEmail { get; set; }
        public string? Role { get; set; }
        public string? IsMslSelected { get; set; }
        public string? IsProtocolsForThreadsAndFillers { get; set; }
        //public string? ReportingManagerEmail { get; set; }
        //public string? Currency { get; set; }
        //public string? OtherCurrencyType { get; set; }
        //public string? BenificiaryName { get; set; }
        //public string? BAnkAccountNumber { get; set; }
        //public string? BankName { get; set; }
        //public string? NameAsPerPan { get; set; }
        //public string? PanCardNumber { get; set; }
        //public string? IfscCode { get; set; }
        //public string? EmailId { get; set; }

        public EventRequestBenificiaryDetails? VenueBenificiaryDetailsData { get; set; }
        public EventRequestBenificiaryDetails? AnaestheticBenificiaryDetailsData { get; set; }
        public List<string>? Files { get; set; }
        public string? IsDeviationUpload { get; set; }
        public int? EventOpen30dayscount { get; set; }
        public List<string>? DeviationFiles { get; set; }


    }
    public class TrainerDetails
    {
        public string? MISCode { get; set; }
        public string? HCPRole { get; set; }

        public string? TrainerName { get; set; }
        public string? TrainerCode { get; set; }
        public string? TrainerQualification { get; set; }
        public string? TrainerCountry { get; set; }
        public string? Speciality { get; set; }
        public string? TrainerCategory { get; set; }
        public string? TrainerType { get; set; }
        //public string? UploadNOC { get; set; }
        public string? Rationale { get; set; }
        public string? IsAnnualTrainerAgreementValid { get; set; }

        public DateTime? FCPAIssueDate { get; set; }
        // public string? UploadTrainerConsent { get; set; }
        public string? IsHonorariumApplicable { get; set; }
        public int? Presentation_Speaking_WorkshopDuration { get; set; }
        public int? DevelopmentofPresentationPanelSessionPreparation { get; set; }
        public int? PaneldiscussionSessionduration { get; set; }
        public int? QASession { get; set; }
        public int? Speaker_TrainerBriefing { get; set; }
        public int? TotalNoOfHours { get; set; }
        public int? HonorariumAmountexcludingTax { get; set; }
        public int? HonorariumAmountincludingTax { get; set; }
        // public string? UploadDeviation { get; set; }
        public int? YTDspendIncludingCurrentEvent { get; set; }
        public string? IsGlobalFMVCheck { get; set; }
        //public string? UploadDocument { get; set; }

        public string? ExpenseType { get; set; }
        public string? IsExpenseBTC_BTE { get; set; }
        public string? IsAccomodationBTC_BTE { get; set; }
        public string? IsLCBTC_BTE { get; set; }
        public string? TravelSelection { get; set; }
        public int? TravelAmountExcludingTax { get; set; }
        public int? TravelAmountIncludingTax { get; set; }
        public int? AccomodationAmountExcludingTax { get; set; }
        public int? AccomodationAmountIncludingTax { get; set; }
        public int? LocalConveyanceAmountexcludingTax { get; set; }
        public int? LocalConveyanceAmountincludingTax { get; set; }
        public int? AgreementAmount { get; set; }
        //public string? UploadDeviation { get; set; }
        public int? TravelandAccomodationspendincludingcurrentevent { get; set; }
        public int? FinalAmount { get; set; }

        public EventRequestBenificiaryDetails? BenificiaryDetailsData { get; set; }
        public List<string>? TrainerFiles { get; set; }


        //public string? Currency { get; set; }
        //public string? EnterCurrencyType { get; set; }
        //public string? BenificiaryName { get; set; }
        //public string? BankAccountNumber { get; set; }
        //public string? BankName { get; set; }
        //public string? NameasPerPAN { get; set; }
        //public string? PANCardNumber { get; set; }
        //public string? IFSCCode { get; set; }
        //public string? EmailID { get; set; }

    }

    public class SlideKitSelectionData
    {
        public string? MISCode { get; set; }
        public string? TrainerName { get; set; }
        public string? TrainerType { get; set; }
        public string? SlideKitSelection { get; set; }
        public string? SlideKitDocument { get; set; }
        public string? IndicatorsForThreads { get; set; }
        public string? IndicatorsForFillers { get; set; }
        public string? IsUpload { get; set; }
        public string? DocToUpload { get; set; }
    }

    public class ProductSelection
    {
        public string? ProductSelectionType { get; set; }
        public string? ProductName { get; set; }
        public int? SamplesRequires { get; set; }
    }

    public class AttenderSelection
    {
        public string InviteedFrom { get; set; }
        public string? MisCode { get; set; }
        public string? AttenderName { get; set; }
        public string? AttenderType { get; set; }
        //public int? UploadNOC { get; set; }
        public string? Experiance { get; set; }
        public string? Qualification { get; set; }
        public string? IsLocalConveyance { get; set; }
        public string? IsBtcorBte { get; set; }
        public int? LocalConveyanceAmountExcludingTax { get; set; }
        public int? LocalConveyanceAmountIncludingTax { get; set; }
       
       
        public string? EmployeeCode { get; set; }
        public string? Designation { get; set; }
        public List<string>? AttenderFiles { get; set; }
    }

    //public class MenariniEmployeeSelection
    //{
    //    public string? EmployeeCode { get; set; }
    //    public string? EmployeeName { get; set; }

    //}

    public class Expense
    {
        public string? ExpenseType { get; set; }
        public string? IsBtcorBte { get; set; }
        public int? ExpenseAmountExcludingTax { get; set; }
        public int? ExpenseAmountIncludingTax { get; set; }
    }
}
