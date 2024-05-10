using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IndiaEvents.Models.Models.RequestSheets
{
    public class EventRequestTrainerData
    {
        public string? EventId { get; set; }
        public string? MISCode { get; set; }
        public string? HCPRole { get; set; }
        public string? TrainerName { get; set; }
        public string? TrainerCode { get; set; }
        public string? TrainerQualification { get; set; }
        public string? TrainerCountry { get; set; }
        public string? Speciality { get; set; }
        public string? TrainerCategory { get; set; }
        public string? TrainerType { get; set; }
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
        public string? IsGlobalFMVCheck { get; set; }
        public string? ExpenseType { get; set; }
        public string? IsExpenseBTC_BTE { get; set; }
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
        public string? EventName { get; set; }
        public string? EventType { get; set; }
        public string? VenueName { get; set; }
        public DateTime? EventStartDate { get; set; }
        public DateTime? EventEndDate { get; set; }
        public string? EventStartTime { get; set; }
        public string? EventEndTime { get; set; }
        public string? City { get; set; }
        public string? State { get; set; }
        public string? SalesHeadEmail { get; set; }
        public string? FinanceEmail { get; set; }
        public string? InitiatorName { get; set; }
        public string? InitiatorEmail { get; set; }


        public EventRequestBenificiaryDetails? BenificiaryDetailsData { get; set; }
        public string? IsDeviationUpload
        {
            get; set;
        }
        public List<string>? DeviationFiles { get; set; }





    }


}
