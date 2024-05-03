using IndiaEvents.Models.Models.RequestSheets;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IndiaEvents.Models.Models.EventTypeSheets
{
    public class Class3
    {
        public DateTime? EventDate { get; set; }
        public string? EventType { get; set; }
        public string? EventName { get; set; }
        public DateTime? EventEndDate { get; set; }
        public string? MeetingType { get; set; }
        public string? Level { get; set; }
        public string? TypeOfParticipation { get; set; }        
        public string? Objective { get; set; }
        public string? VenueName { get; set; }
        public string? State { get; set; }
        public string? City { get; set; }
        public string? TargetParticipants { get; set; }

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
        public string? FinanceHeadEmail { get; set; }
        public string? SalesCoordinatorEmail { get; set; }
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
    public class Class3PanelDetails
    {
        public string? HcpRole { get; set; }
        public string? HCPRoleName { get; set; }
        public string? MISCode { get; set; }
        public string? HcpName { get; set; }

        public string? SpeakerCode { get; set; }
        public string? TrainerCode { get; set; }

        public string? HcpQualification { get; set; }
        public string? Speciality { get; set; }
        public string? HcpCategory { get; set; }
        public string? HcpType { get; set; }
        public string? HcpCountry { get; set; }
        public string? Rationale { get; set; }
        public DateTime? FCPAIssueDate { get; set; }

        public string? IsHonorariumApplicable { get; set; }
        public int? Presentation_Speaking_WorkshopDuration { get; set; }
        public int? DevelopmentofPresentationPanelSessionPreparation { get; set; }
        public int? PaneldiscussionSessionduration { get; set; }
        public int? QASession { get; set; }
        public int? Speaker_TrainerBriefing { get; set; }
        public int? TotalNoOfHours { get; set; }
        public int? HonorariumAmountexcludingTax { get; set; }
        public int? HonorariumAmountincludingTax { get; set; }
        public int? YTDspendIncludingCurrentEvent { get; set; }
        public string? ExpenseType { get; set; }
        public string? IsTravelBTC_BTE { get; set; }
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
        public int? TravelandAccomodationspendincludingcurrentevent { get; set; }
        public int? FinalAmount { get; set; }
        public EventRequestBenificiaryDetails? BenificiaryDetailsData { get; set; }
        public string? IsFilesUpload { get; set; }
        public List<string>? Files { get; set; }


    }
    public class Class3ExpenseDetails
    {
        public int CMESponsershipAmountExcludingTax { get; set; }
        public int CMESponsershipAmountIncludingTax { get; set; }
        //public int CMESponsershipAmountIncludingTax { get; set; }
    }
}
