using IndiaEvents.Models.Models.EventTypeSheets;
using IndiaEvents.Models.Models.RequestSheets;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IndiaEvents.Models.Models.Draft
{
    public class UpdateDataForClassI
    {

        public UpdateEventDetails EventDetails { get; set; }
        public List<UpdateBrandSelection> BrandSelection { get; set; }
        public List<UpdatePanelSelection> PanelSelection { get; set; }
        public List<UpdateSlideKitSelection> SlideKitSelection { get; set; }
        public List<UpdateInviteeSelection> InviteeSelection { get; set; }
        public List<UpdateExpenseSelection> ExpenseSelection { get; set; }
        public string IsDeviationUpload { get; set; }
        public List<EventRequestDeviationsData>? DeviationDetails { get; set; }
        //public List<string> DeviationFiles { get; set; }
    }

    public class UpdateDataForStall
    {
        public UpdateStallFabricationDetails EventDetails { get; set; }
        public List<UpdateBrandSelection> BrandSelection { get; set; }
        public List<UpdateExpenseSelection> ExpenseSelection { get; set; }
        public string IsDeviationUpload { get; set; }
        public List<EventRequestDeviationsData>? DeviationDetails { get; set; }
    }

    public class UpdateDataForHandsOnTraining
    {

        public UpdateHandsOnDetails EventDetails { get; set; }
        public List<UpdateBrandSelection> BrandSelection { get; set; }
        public List<TrainerDetails> PanelSelection { get; set; }
        public List<SlideKitSelectionData> SlideKitSelection { get; set; }
        public List<ProductSelection>? ProductSelections { get; set; }
        public List<AttenderSelection> InviteeSelection { get; set; }
        public List<Expense> ExpenseSelection { get; set; }
        public string IsDeviationUpload { get; set; }
        public List<EventRequestDeviationsData>? DeviationDetails { get; set; }
        //public List<string> DeviationFiles { get; set; }
    }

    public class UpdateBrandSelection
    {
        public string Id { get; set; }
        public string BrandName { get; set; }
        public string? PercentageAllocation { get; set; }
        public string ProjectId { get; set; }
    }

    public class UpdateEventDetails
    {
        public string Id { get; set; }
        public string EventTopic { get; set; }
        public string EventType { get; set; }
        public DateTime EventDate { get; set; }
        public string StartTime { get; set; }
        public string EndTime { get; set; }
        public string VenueName { get; set; }
        public string State { get; set; }
        public string City { get; set; }
        public string MeetingType { get; set; }
        public string Brands { get; set; }
        public string Panelists { get; set; }
        public string SlideKits { get; set; }
        public string Invitees { get; set; }
        public string MIPLInvitees { get; set; }
        public string Expenses { get; set; }
        public string ExpenseDataBTE { get; set; }
        public string Sales_Head { get; set; }
        public string FinanceHead { get; set; }
        public string SalesCoordinator { get; set; }
        public string InitiatorName { get; set; }
        public string Initiator_Email { get; set; }
        public double? TotalExpenseBTC { get; set; }
        public double? TotalExpenseBTE { get; set; }
        public double? TotalHonorariumAmount { get; set; }
        public double? TotalTravelAccommodationAmount { get; set; }
        public double? TotalAccomodationAmount { get; set; }
        public double? TotalBudget { get; set; }
        public double? TotalLocalConveyance { get; set; }
        public double? TotalTravelAmount { get; set; }
        public double? TotalExpense { get; set; }
        public double? EventOpen30dayscount { get; set; }

        public string IsFilesUpload { get; set; }
        public List<UpdateFiles> Files { get; set; }



    }

    public class UpdateHandsOnDetails
    {
        public string Id { get; set; }
        public string EventName { get; set; }
        public string EventType { get; set; }
        public DateTime EventDate { get; set; }
        public string StartTime { get; set; }
        public string EndTime { get; set; }
        public string? ProductBrandSelection { get; set; }
        public string? ModeOfTraining { get; set; }

        public string VenueName { get; set; }
        public string State { get; set; }
        public string City { get; set; }
        public string? VendorName { get; set; }
        public string? VenueSelectionCheckList { get; set; }
        public string? IsVenueHasAnyEmergancySupport { get; set; }
        public string? EmerganctContact { get; set; }
        public string? IsVenueFacilityCharges { get; set; }
        public string? VenueFacilityChargesBtc_Bte { get; set; }
        public double? FacilityChargesExcludingTax { get; set; }
        public double? FacilityChargesIncludingTax { get; set; }
        public string? IsAnesthetistRequired { get; set; }
        public string? AnesthetistRequiredBtc_Bte { get; set; }
        public double? AnesthetistChargesExcludingTax { get; set; }
        public double? AnesthetistChargesIncludingTax { get; set; }

        public string Brands { get; set; }
        public string Panelists { get; set; }
        public string SlideKits { get; set; }
        public string Invitees { get; set; }
        public string MIPLInvitees { get; set; }
        public string Expenses { get; set; }

        public string Sales_Head { get; set; }
        public string FinanceHead { get; set; }
        public string InitiatorName { get; set; }
        public string Initiator_Email { get; set; }
        public string SalesCoordinator { get; set; }
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
        public double? EventOpen30dayscount { get; set; }

        public string IsFilesUpload { get; set; }
        public List<UpdateFiles> Files { get; set; }

        public EventRequestBenificiaryDetails? VenueBenificiaryDetailsData { get; set; }
        public EventRequestBenificiaryDetails? AnaestheticBenificiaryDetailsData { get; set; }


    }

    public class UpdateStallFabricationDetails
    {
        public string Id { get; set; }
        public string EventTopic { get; set; }
        public string EventType { get; set; }
        public DateTime EventStartDate { get; set; }
        public DateTime EventEndDate { get; set; }
        public string ClassIIIEventCode { get; set; }
        public string BrandsData { get; set; }
        public string ExpenseData { get; set; }
        public string ExpenseDataBTE { get; set; }
        public string SalesCoordinator { get; set; }
        public string Sales_Head { get; set; }
        public string FinanceHead { get; set; }
        public string InitiatorName { get; set; }
        public string Initiator_Email { get; set; }
        public double? TotalExpenseBTC { get; set; }
        public double? TotalExpenseBTE { get; set; }
        public double? TotalBudget { get; set; }
        public double? TotalExpense { get; set; }
        public string? IsAdvanceRequired { get; set; }
        public double? AdvanceAmount { get; set; }
        public double? EventOpen30dayscount { get; set; }
        public string? IsFilesUpload { get; set; }
        public List<UpdateFiles>? Files { get; set; }



    }

    public class UpdateExpenseSelection
    {
        public string Id { get; set; }
        public string Expense { get; set; }
        public string ExpenseType { get; set; }
        public double ExpenseAmountIncludingTax { get; set; }
        public double ExpenseAmountExcludingTax { get; set; }
    }

    public class UpdateInviteeSelection
    {
        public string Id { get; set; }
        public string InviteeFrom { get; set; }
        public string Name { get; set; }
        public string MisCode { get; set; }
        public string EmployeeCode { get; set; }
        public string IsLocalConveyance { get; set; }
        public string LocalConveyanceType { get; set; }
        public string Speciality { get; set; }
        public string Designation { get; set; }
        public string HCPType { get; set; }
        public double LocalConveyanceAmountIncludingTax { get; set; }
        public double LocalConveyanceAmountExcludingTax { get; set; }
    }

    public class UpdatePanelSelection
    {
        public string Id { get; set; }
        public string SpeakerCode { get; set; }
        public string TrainerCode { get; set; }
        public string Speciality { get; set; }
        public string Tier { get; set; }
        public string Qualification { get; set; }
        public string Country { get; set; }
        public string Rationale { get; set; }
        public string FcpaIssueDate { get; set; }

        public double PresentationDuration { get; set; }
        public double PanelSessionPreperationDuration { get; set; }
        public double PanelDisscussionDuration { get; set; }
        public double QaSessionDuration { get; set; }
        public double BriefingSession { get; set; }
        public double TotalSessionHours { get; set; }
        public string HcpRole { get; set; }
        public string HcpName { get; set; }
        public string MisCode { get; set; }
        public string GOorNGO { get; set; }
        public string ExpenseType { get; set; }
        public string HonorariumRequired { get; set; }
        public double HonarariumAmountIncludingTax { get; set; }
        public double HonarariumAmountExcludingTax { get; set; }
        public double TravelAmountIncludingTax { get; set; }
        public double TravelExcludingTax { get; set; }
        public string TravelBtcorBte { get; set; }
        public string OthersType { get; set; }
        public double LocalConveyanceIncludingTax { get; set; }
        public double LocalConveyanceExcludingTax { get; set; }
        public double FinalAmount { get; set; }
        public double AgreementAmount { get; set; }
        public string LcBtcorBte { get; set; }
        public double AccomdationIncludingTax { get; set; }
        public double AccomdationExcludingTax { get; set; }
        public string AccomodationBtcorBte { get; set; }
        public string PanCardName { get; set; }
        public string BankAccountNumber { get; set; }
        public string IFSCCode { get; set; }
        public string BankName { get; set; }
        public string Currency { get; set; }
        public string OtherCurrencyType { get; set; }
        public string BeneficiaryName { get; set; }
        public string PanNumber { get; set; }
        public string IsGlobalFMVCheck { get; set; }
        public string SwiftCode { get; set; }
        public string IsFilesUpload { get; set; }
        public List<string> Files { get; set; }
    }

    public class UpdateSlideKitSelection
    {
        public string Id { get; set; }
        public string HcpName { get; set; }
        public string MisCode { get; set; }
        public string HcpType { get; set; }
        public string SlideKitType { get; set; }
        public string SlideKitOption { get; set; }
        public string IsFilesUpload { get; set; }
        public List<string> Files { get; set; }
    }

    public class UpdateFiles
    {
        public long? Id { get; set; }
        public string? FileBase64 { get; set; }
    }



    public class GetEventDetails
    {
        public string? EventDate { get; set; }
        public string? EventTopic { get; set; }
        public string? ClassIIIEventCode { get; set; }
        public string? StartTime { get; set; }
        public string? EndTime { get; set; }
        public string? State { get; set; }
        public string? VenueName { get; set; }
        public string? City { get; set; }
        public string? Brands { get; set; }
        public string? Panelists { get; set; }
        public string? SlideKits { get; set; }
        public string? Invitees { get; set; }
        public string? MIPLInvitees { get; set; }
        public string? Expenses { get; set; }
        public string? MeetingType { get; set; }
    }
}
