using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IndiaEvents.Models.Models.GetData
{
    public class ProcessSheetPayload
    {
        [JsonProperty("Event Topic")]
        public string EventTopic { get; set; }
        [JsonProperty("EventId/EventRequestId")]
        public string EventIdEventRequestId { get; set; }
        public string EventType { get; set; }
        public string EventDate { get; set; }
        public string StartTime { get; set; }
        public string EndTime { get; set; }
        public string VenueName { get; set; }
        [JsonProperty("Event End Date")]
        public string EventEndDate { get; set; }
        public string State { get; set; }
        public string City { get; set; }
        public string InitiatorName { get; set; }
        [JsonProperty("Initiator Email")]
        public string InitiatorEmail { get; set; }
        //public DateTime CreatedOn { get; set; }
        [JsonProperty("Created Date - Helper")]
        public string CreatedDateHelper { get; set; }
        //public DateTime Modified { get; set; }
        public string IsAdvanceRequired { get; set; }
        [JsonProperty("PRE-RBM/BM Approval")]
        public string PRERBMBMApproval { get; set; }
        [JsonProperty("PRE-RBM/BM Approval Date")]
        public string PRERBMBMApprovalDate { get; set; }
        [JsonProperty("PRE-Sales Head Approval")]
        public string PRESalesHeadApproval { get; set; }
        [JsonProperty("PRE-Sales Head Approval Date")]
        public string PRESalesHeadApprovalDate { get; set; }
        [JsonProperty("PRE-Marketing Head Approval")]
        public string PREMarketingHeadApproval { get; set; }
        [JsonProperty("PRE-Marketing Head Approval Date")]
        public string PREMarketingHeadApprovalDate { get; set; }
        public string agreement { get; set; }
        [JsonProperty("PRE-Finance Treasury Approval Date")]
        public string PREFinanceTreasuryApprovalDate { get; set; }
        [JsonProperty("PRE-Medical Affairs Head Approval")]
        public string PREMedicalAffairsHeadApproval { get; set; }
        [JsonProperty("PRE-Medical Affairs Head Approval Date")]
        public string PREMedicalAffairsHeadApprovalDate { get; set; }
        [JsonProperty("PRE-Sales Coordinator Approval")]
        public string PRESalesCoordinatorApproval { get; set; }
        [JsonProperty("PRE-Sales Coordinator Approval Date")]
        public string PRESalesCoordinatorApprovalDate { get; set; }
        [JsonProperty("PRE-Compliance Approval")]
        public string PREComplianceApproval { get; set; }
        [JsonProperty("PRE-Compliance Approval Date")]
        public string PREComplianceApprovalDate { get; set; }
        [JsonProperty("Event Request Status")]
        public string EventRequestStatus { get; set; }
        [JsonProperty("Event Approved Date")]
        public string EventApprovedDate { get; set; }
        [JsonProperty("Honorarium Request Status")]
        public string HonorariumRequestStatus { get; set; }

        [JsonProperty("Honorarium Approved Date")]
        public string HonorariumApprovedDate { get; set; }

        [JsonProperty("Post Event Request status")]
        public string PostEventRequeststatus { get; set; }

        [JsonProperty("Helper Finance treasury trigger(BTE)")]
        public string HelperFinancetreasurytriggerBTE { get; set; }

        [JsonProperty("Post Event Approved Date")]
        public object PostEventApprovedDate { get; set; }
        public string EventOpenSalesHeadApproval { get; set; }

        [JsonProperty("EventOpenSalesHeadApproval Date")]
        public string EventOpenSalesHeadApprovalDate { get; set; }

        [JsonProperty("7daysSalesHeadApproval")]
        public string _7daysSalesHeadApproval { get; set; }

        [JsonProperty("7daysSalesHeadApproval date")]
        public string _7daysSalesHeadApprovaldate { get; set; }

        [JsonProperty("PRE-F&B Expense Excluding Tax Approval")]
        public string PREFBExpenseExcludingTaxApproval { get; set; }

        [JsonProperty("HCP exceeds 1,00,000 FH Approval")]
        public string HCPexceeds100000FHApproval { get; set; }

        [JsonProperty("HCP exceeds 5,00,000 Trigger FH Approval")]
        public string HCPexceeds500000TriggerFHApproval { get; set; }

        [JsonProperty("HCP Honorarium 6,00,000 Exceeded Approval")]
        public string HCPHonorarium600000ExceededApproval { get; set; }

        [JsonProperty("Trainer Honorarium 12,00,000 Exceeded Approval")]
        public string TrainerHonorarium1200000ExceededApproval { get; set; }

        [JsonProperty("Travel/Accomodation 3,00,000 Exceeded Approval")]
        public string TravelAccomodation300000ExceededApproval { get; set; }

        [JsonProperty("Honorarium Submitted?")]
        public string HonorariumSubmitted { get; set; }

        [JsonProperty("Agreement Download")]
        public string AgreementDownload { get; set; }
        public string EventOpen30days { get; set; }
        public string EventWithin7days { get; set; }

        [JsonProperty("View Event Request")]
        public string ViewEventRequest { get; set; }

        [JsonProperty("View Honorarium Request")]
        public string ViewHonorariumRequest { get; set; }

        [JsonProperty("Reporting Manager")]
        public string ReportingManager { get; set; }

        [JsonProperty("1 Up Manager")]
        public string _1UpManager { get; set; }
        public string Brands { get; set; }
        public string Panelists { get; set; }
        public string HCP { get; set; }
        public string SlideKits { get; set; }
        public string Invitees { get; set; }

        [JsonProperty("MIPL Invitees")]
        public string MIPLInvitees { get; set; }
        public string Expenses { get; set; }

        [JsonProperty("Other Expenses")]
        public string OtherExpenses { get; set; }

        [JsonProperty("Total Invitees")]
        public int TotalInvitees { get; set; }

        [JsonProperty("Total Attendees")]
        public int TotalAttendees { get; set; }
        public string Finance { get; set; }

        [JsonProperty("Approval Status")]
        public string ApprovalStatus { get; set; }

        [JsonProperty("Next Approver")]
        public string NextApprover { get; set; }

        [JsonProperty("RBM/BM")]
        public string RBMBM { get; set; }

        [JsonProperty("Sales Head")]
        public string SalesHead { get; set; }

        [JsonProperty("Marketing Head")]
        public string MarketingHead { get; set; }
        public string Compliance { get; set; }

        [JsonProperty("Finance Treasury")]
        public string FinanceTreasury { get; set; }

        [JsonProperty("Finance Accounts")]
        public string FinanceAccounts { get; set; }

        [JsonProperty("Medical Affairs Head")]
        public string MedicalAffairsHead { get; set; }

        [JsonProperty("Finance Head")]
        public string FinanceHead { get; set; }

        [JsonProperty("Sales Coordinator")]
        public string SalesCoordinator { get; set; }

        [JsonProperty("Total Honorarium Amount")]
        public int TotalHonorariumAmount { get; set; }

        [JsonProperty("Total Travel & Accommodation Amount")]
        public int TotalTravelAccommodationAmount { get; set; }

        [JsonProperty("Total Travel Amount")]
        public int TotalTravelAmount { get; set; }

        [JsonProperty("Total Accommodation Amount")]
        public int TotalAccommodationAmount { get; set; }

        [JsonProperty("Total Accomodation Amount")]
        public int TotalAccomodationAmount { get; set; }

        [JsonProperty("Total Local Conveyance")]
        public int TotalLocalConveyance { get; set; }

        [JsonProperty("Total Expense")]
        public int TotalExpense { get; set; }

        [JsonProperty("Other Expense Amount")]
        public int OtherExpenseAmount { get; set; }

        [JsonProperty("Total Budget")]
        public int TotalBudget { get; set; }

        [JsonProperty(" Total Expense BTC")]
        public int TotalExpenseBTC { get; set; }

        [JsonProperty("Total Expense BTE")]
        public int TotalExpenseBTE { get; set; }

        [JsonProperty("Cost per participant - Helper")]
        public int CostperparticipantHelper { get; set; }

        [JsonProperty("Advance Amount")]
        public int AdvanceAmount { get; set; }

        [JsonProperty("Advance Voucher Number")]
        public string AdvanceVoucherNumber { get; set; }

        [JsonProperty("Advance Voucher Date")]
        public string AdvanceVoucherDate { get; set; }

        [JsonProperty("Bank Reference Number")]
        public string BankReferenceNumber { get; set; }

        [JsonProperty("Bank Reference Date")]
        public string BankReferenceDate { get; set; }

        [JsonProperty("PRE-Finance Treasury Approval")]
        public string PREFinanceTreasuryApproval { get; set; }

        [JsonProperty("EventSettlement - Deviation Approval Date")]
        public string EventSettlementDeviationApprovalDate { get; set; }

        [JsonProperty("View Post Event Request")]
        public string ViewPostEventRequest { get; set; }
        public string Role { get; set; }

        [JsonProperty("HCP Role")]
        public string HCPRole { get; set; }

        [JsonProperty("Class III Event Code")]
        public string ClassIIIEventCode { get; set; }

        [JsonProperty("Meeting Type")]
        public string MeetingType { get; set; }

        [JsonProperty("Sponsorship Society Name")]
        public string SponsorshipSocietyName { get; set; }

        [JsonProperty("Venue Country")]
        public string VenueCountry { get; set; }

        [JsonProperty("Total HCP Registration Spend")]
        public int TotalHCPRegistrationSpend { get; set; }

        [JsonProperty("Total HCP Registration Amount")]
        public int TotalHCPRegistrationAmount { get; set; }

        [JsonProperty("Medical Utility Type")]
        public string MedicalUtilityType { get; set; }

        [JsonProperty("Medical Utility Description")]
        public string MedicalUtilityDescription { get; set; }

        [JsonProperty("Valid To")]
        public string ValidTo { get; set; }

        [JsonProperty("Valid From")]
        public string ValidFrom { get; set; }

        [JsonProperty("Product Brand")]
        public string ProductBrand { get; set; }

        [JsonProperty("Mode of Training")]
        public string ModeofTraining { get; set; }

        [JsonProperty("Emergency Support ")]
        public string EmergencySupport { get; set; }

        [JsonProperty("Facility Charges")]
        public string FacilityCharges { get; set; }

        [JsonProperty("Facility charges Excluding Tax")]
        public int FacilitychargesExcludingTax { get; set; }

        [JsonProperty("Total Facility charges including Tax")]
        public int TotalFacilitychargesincludingTax { get; set; }

        [JsonProperty("HOT Webinar Type")]
        public string HOTWebinarType { get; set; }

        [JsonProperty("HOT Webinar Vendor Name")]
        public string HOTWebinarVendorName { get; set; }

        [JsonProperty("Venue Selection Checklist")]
        public string VenueSelectionChecklist { get; set; }

        [JsonProperty("Emergency Contact No")]
        public string EmergencyContactNo { get; set; }

        [JsonProperty("Facility Charges BTC/BTE")]
        public string FacilityChargesBTCBTE { get; set; }

        [JsonProperty("Anesthetist Required?")]
        public string AnesthetistRequired { get; set; }

        [JsonProperty("Anesthetist BTC/BTE")]
        public string AnesthetistBTCBTE { get; set; }

        [JsonProperty("Anesthetist Excluding Tax")]
        public int AnesthetistExcludingTax { get; set; }

        [JsonProperty("Anesthetist including Tax")]
        public int AnesthetistincludingTax { get; set; }

        [JsonProperty("Selected Products")]
        public string SelectedProducts { get; set; }

        [JsonProperty("Beneficiary Details")]
        public string BeneficiaryDetails { get; set; }

        [JsonProperty("pre-45 days Approval")]
        public string pre45daysApproval { get; set; }

        [JsonProperty("pre 5 days approval")]
        public string pre5daysapproval { get; set; }

        [JsonProperty("f&b approved")]
        public string fbapproved { get; set; }

        [JsonProperty("HCP 1L Approved")]
        public string HCP1LApproved { get; set; }

        [JsonProperty("HCP 5L Approval")]
        public string HCP5LApproval { get; set; }

        [JsonProperty("HCP 6L Approval")]
        public string HCP6LApproval { get; set; }

        [JsonProperty("Trainer 12L Approval")]
        public string Trainer12LApproval { get; set; }

        [JsonProperty("T/A Approval")]
        public string TAApproval { get; set; }

        [JsonProperty("Compliance approval")]
        public string Complianceapproval { get; set; }

        [JsonProperty("Agreements?")]
        public string Agreements { get; set; }

        [JsonProperty("Is Protocols Given?")]
        public string IsProtocolsGiven { get; set; }

        [JsonProperty("Is MSL Selected?")]
        public string IsMSLSelected { get; set; }
        public string Objective { get; set; }
        public string Comments { get; set; }
    }

    public class ApprovedSpeakersGetPayload
    {
        public string SpeakerId { get; set; }
        public string SpeakerName { get; set; }

        [JsonProperty("Speaker Code")]
        public string SpeakerCode { get; set; }

        [JsonProperty("Speaker Category")]
        public string SpeakerCategory { get; set; }
        public string Speciality { get; set; }

        [JsonProperty("Speaker Type")]
        public string SpeakerType { get; set; }
        public string MisCode { get; set; }

        [JsonProperty("Aggregate Honorarium Limit")]
        public int AggregateHonorariumLimit { get; set; }

        [JsonProperty("Aggregate Accommodataion Limit")]
        public int AggregateAccommodataionLimit { get; set; }

        [JsonProperty("Aggregate Honorarium Spent")]
        public int AggregateHonorariumSpent { get; set; }

        [JsonProperty("Aggregate Accommodation Spent")]
        public int AggregateAccommodationSpent { get; set; }

        [JsonProperty("Aggregate Spent on Medical Utility")]
        public int AggregateSpentonMedicalUtility { get; set; }

        [JsonProperty("Aggregate Spent as HCP Consultant")]
        public int AggregateSpentasHCPConsultant { get; set; }

        [JsonProperty("FCPA Sign Off Date")]
        public string FCPASignOffDate { get; set; }

        [JsonProperty("FCPA Expiry Date")]
        public string FCPAExpiryDate { get; set; }

        [JsonProperty("FCPA Valid?")]
        public string FCPAValid { get; set; }

        [JsonProperty("TRC Date")]
        public object TRCDate { get; set; }

        [JsonProperty("TRC Valid?")]
        public object TRCValid { get; set; }

        [JsonProperty("Total Spend on Speaker")]
        public object TotalSpendonSpeaker { get; set; }

        [JsonProperty("Aggregate Exhausted?")]
        public object AggregateExhausted { get; set; }
        public string Division { get; set; }
        public string Qualification { get; set; }
        public string Address { get; set; }
        public string State { get; set; }
        public string City { get; set; }
        public string Country { get; set; }

        [JsonProperty("Contact Number")]
        public string ContactNumber { get; set; }

        [JsonProperty("ABM Name (Requestor)")]
        public object ABMNameRequestor { get; set; }
        public object SpeakerCategoryId { get; set; }

        [JsonProperty("Speaker Criteria")]
        public string SpeakerCriteria { get; set; }

        [JsonProperty("Is Active")]
        public string IsActive { get; set; }
        public object CVdocument { get; set; }
        public string Created { get; set; }

        [JsonProperty("Created Date - Helper")]
        public string CreatedDateHelper { get; set; }

        [JsonProperty("Sales Alert Trigger")]
        public object SalesAlertTrigger { get; set; }

        [JsonProperty("Medical Affairs Alert Trigger")]
        public object MedicalAffairsAlertTrigger { get; set; }

        [JsonProperty("Sales Head Approval")]
        public string SalesHeadApproval { get; set; }

        [JsonProperty("Sales Head Approval Date")]
        public object SalesHeadApprovalDate { get; set; }

        [JsonProperty("Medical Affairs Head Approval")]
        public string MedicalAffairsHeadApproval { get; set; }

        [JsonProperty("Medical Affairs Head Approval Date")]
        public object MedicalAffairsHeadApprovalDate { get; set; }

        [JsonProperty("Speaker Criteria Details")]
        public object SpeakerCriteriaDetails { get; set; }

        [JsonProperty("Sales Head")]
        public string SalesHead { get; set; }

        [JsonProperty("Medical Affairs Head")]
        public string MedicalAffairsHead { get; set; }

        [JsonProperty("Initiator Name")]
        public string InitiatorName { get; set; }

        [JsonProperty("Initiator Email")]
        public string InitiatorEmail { get; set; }

        [JsonProperty("NA ID")]
        public object NAID { get; set; }
    }

    public class ApprovedTrainersGetPayload
    {
        public string TrainerId { get; set; }
        public string TrainerName { get; set; }
        public string TrainerCode { get; set; }
        public string TrnrCode { get; set; }
        public string TierType { get; set; }
        public string Speciality { get; set; }
        public string Is_NONGO { get; set; }
        public string MisCode { get; set; }

        [JsonProperty("NA ID")]
        public object NAID { get; set; }

        [JsonProperty("Aggregate spend on Honorarium - Trainer")]
        public int AggregatespendonHonorariumTrainer { get; set; }

        [JsonProperty("Aggregate Limit on Accomodation")]
        public int AggregateLimitonAccomodation { get; set; }

        [JsonProperty("Aggregate Honorarium Spent ")]
        public int AggregateHonorariumSpent { get; set; }

        [JsonProperty("Aggregate spend on Accomodation")]
        public int AggregatespendonAccomodation { get; set; }

        [JsonProperty("Aggregate Spent on Medical Utility")]
        public int AggregateSpentonMedicalUtility { get; set; }

        [JsonProperty("Aggregate Spent as HCP Consultant")]
        public int AggregateSpentasHCPConsultant { get; set; }
        public string Qualification { get; set; }
        public object TrainerCategoryId { get; set; }
        public object CityId { get; set; }
        public object StateId { get; set; }
        public string Country { get; set; }
        public string IsActive { get; set; }

        [JsonProperty("FCPA Sign Off Date")]
        public string FCPASignOffDate { get; set; }

        [JsonProperty("FCPA Expiry Date")]
        public string FCPAExpiryDate { get; set; }

        [JsonProperty("FCPA Valid?")]
        public string FCPAValid { get; set; }
        public object CV_Document { get; set; }

        [JsonProperty("Aggregate Exhausted?")]
        public object AggregateExhausted { get; set; }
        public string Created { get; set; }

        [JsonProperty("Created Date - Helper")]
        public string CreatedDateHelper { get; set; }

        [JsonProperty("Sales Alert Trigger")]
        public object SalesAlertTrigger { get; set; }

        [JsonProperty("Medical Affairs Head Alert Trigger")]
        public object MedicalAffairsHeadAlertTrigger { get; set; }

        [JsonProperty("Sales Head Approval")]
        public string SalesHeadApproval { get; set; }

        [JsonProperty("Sales Head Approval Date")]
        public object SalesHeadApprovalDate { get; set; }

        [JsonProperty("Medical Affairs Head Approval")]
        public object MedicalAffairsHeadApproval { get; set; }

        [JsonProperty("Medical Affairs Head Approval Date")]
        public object MedicalAffairsHeadApprovalDate { get; set; }

        [JsonProperty("Sales Head")]
        public string SalesHead { get; set; }

        [JsonProperty("Medical Affairs Head")]
        public string MedicalAffairsHead { get; set; }

        [JsonProperty("Initiator Name")]
        public string InitiatorName { get; set; }

        [JsonProperty("Initiator Email")]
        public string InitiatorEmail { get; set; }

        [JsonProperty("TrainerTybe - Shortcode")]
        public object TrainerTybeShortcode { get; set; }

        [JsonProperty("Trainer Code")]
        public object TrainerCodeNew { get; set; }

        [JsonProperty("Trainer Brand")]
        public string TrainerBrand { get; set; }

        [JsonProperty("Trainer Type")]
        public string TrainerType { get; set; }
        public string Division { get; set; }
        public string Address { get; set; }
        public object City { get; set; }
        public object State { get; set; }

        [JsonProperty("Contact Number")]
        public string ContactNumber { get; set; }

        [JsonProperty("Trained by")]
        public string Trainedby { get; set; }

        [JsonProperty("Trainer CV")]
        public string TrainerCV { get; set; }

        [JsonProperty("Trainer certificate")]
        public string Trainercertificate { get; set; }

        [JsonProperty("Trained on")]
        public string Trainedon { get; set; }

        [JsonProperty("Trainer Category")]
        public string TrainerCategory { get; set; }

        [JsonProperty("Trainer Criteria")]
        public string TrainerCriteria { get; set; }

        [JsonProperty("Trainer Criteria Details")]
        public string TrainerCriteriaDetails { get; set; }

        [JsonProperty("Medical Affairs Alert Trigger")]
        public object MedicalAffairsAlertTrigger { get; set; }
    }

    public class ApprovedVendorsGetPayload
    {
        public string VendorId { get; set; }
        public string VendorAccount { get; set; }
        public int MisCode { get; set; }

        [JsonProperty("Requestor Name")]
        public string RequestorName { get; set; }

        [JsonProperty("Bank Name")]
        public string BankName { get; set; }
        public string BeneficiaryName { get; set; }
        public string PanCardName { get; set; }
        public string PanNumber { get; set; }
        public string BankAccountNumber { get; set; }
        public string IfscCode { get; set; }

        [JsonProperty("Swift Code")]
        public object SwiftCode { get; set; }

        [JsonProperty("IBN Number")]
        public object IBNNumber { get; set; }

        [JsonProperty("Email ")]
        public string Email { get; set; }

        [JsonProperty("Pancard Document")]
        public object PancardDocument { get; set; }

        [JsonProperty("Cheque Document")]
        public object ChequeDocument { get; set; }

        [JsonProperty("Tax Residence Certificate")]
        public object TaxResidenceCertificate { get; set; }

        [JsonProperty("Tax Residence Certificate Date")]
        public string TaxResidenceCertificateDate { get; set; }

        [JsonProperty("IsActive?")]
        public string IsActive { get; set; }

        [JsonProperty("Finance Checker-1")]
        public string FinanceChecker1 { get; set; }

        [JsonProperty("Finance Checker Approval")]
        public string FinanceCheckerApproval { get; set; }

        [JsonProperty("Finance Checker Approval Date")]
        public string FinanceCheckerApprovalDate { get; set; }
        public string Requestor { get; set; }

        [JsonProperty("Initiator Name")]
        public string InitiatorName { get; set; }

        [JsonProperty("Initiator Email")]
        public string InitiatorEmail { get; set; }

        [JsonProperty("Finance Checker ")]
        public object FinanceChecker { get; set; }

        [JsonProperty("Finance Checker  approval")]
        public object FinanceCheckerapproval { get; set; }

        [JsonProperty("Finance Checker")]
        public object FinanceChecker_ { get; set; }
    }
}
