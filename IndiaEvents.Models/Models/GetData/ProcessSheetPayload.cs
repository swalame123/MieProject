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
        public bool EventCancelled { get; set; }
        public string ReasonForCancellation { get; set; }
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
        public double TotalInvitees { get; set; }

        [JsonProperty("Total Attendees")]
        public double TotalAttendees { get; set; }
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
        public double TotalHonorariumAmount { get; set; }

        [JsonProperty("Total Travel & Accommodation Amount")]
        public double TotalTravelAccommodationAmount { get; set; }

        [JsonProperty("Total Travel Amount")]
        public double TotalTravelAmount { get; set; }

        [JsonProperty("Total Accommodation Amount")]
        public double TotalAccommodationAmount { get; set; }

        [JsonProperty("Total Accomodation Amount")]
        public double TotalAccomodationAmount { get; set; }

        [JsonProperty("Total Local Conveyance")]
        public double TotalLocalConveyance { get; set; }

        [JsonProperty("Total Expense")]
        public double TotalExpense { get; set; }

        [JsonProperty("Other Expense Amount")]
        public double OtherExpenseAmount { get; set; }

        [JsonProperty("Total Budget")]
        public double TotalBudget { get; set; }

        [JsonProperty(" Total Expense BTC")]
        public double TotalExpenseBTC { get; set; }

        [JsonProperty("Total Expense BTE")]
        public double TotalExpenseBTE { get; set; }

        [JsonProperty("Cost per participant - Helper")]
        public double CostperparticipantHelper { get; set; }

        [JsonProperty("Advance Amount")]
        public double AdvanceAmount { get; set; }

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
        public double TotalHCPRegistrationSpend { get; set; }

        [JsonProperty("Total HCP Registration Amount")]
        public double TotalHCPRegistrationAmount { get; set; }

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
        public double FacilitychargesExcludingTax { get; set; }

        [JsonProperty("Total Facility charges including Tax")]
        public double TotalFacilitychargesincludingTax { get; set; }

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
        public double AnesthetistExcludingTax { get; set; }

        [JsonProperty("Anesthetist including Tax")]
        public double AnesthetistincludingTax { get; set; }

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

    public class HonorariumPayload
    {
        [JsonProperty("EventId/EventRequestId")]
        public string EventIdEventRequestId { get; set; }

        [JsonProperty("Event Type")]
        public string EventType { get; set; }

        [JsonProperty("Event Topic")]
        public string EventTopic { get; set; }

        [JsonProperty("Event Date")]
        public string EventDate { get; set; }

        [JsonProperty("Start Time")]
        public string StartTime { get; set; }

        [JsonProperty("End Time")]
        public string EndTime { get; set; }

        [JsonProperty("Event End Date")]
        public object EventEndDate { get; set; }

        [JsonProperty("Venue Name")]
        public string VenueName { get; set; }

        [JsonProperty("Initiator Name")]
        public string InitiatorName { get; set; }

        [JsonProperty("Initiator Email")]
        public string InitiatorEmail { get; set; }
        public string City { get; set; }
        public string State { get; set; }

        [JsonProperty("Honorarium Submitted?")]
        public string HonorariumSubmitted { get; set; }

        [JsonProperty("Event Request Status")]
        public object EventRequestStatus { get; set; }

        [JsonProperty("Honorarium Request Status")]
        public string HonorariumRequestStatus { get; set; }

        [JsonProperty("Honorarium Approved Date")]
        public string HonorariumApprovedDate { get; set; }
        public string Brands { get; set; }
        public string Panelists { get; set; }
        public string SlideKits { get; set; }
        public string Invitees { get; set; }
        public string Expenses { get; set; }

        [JsonProperty("Panelists & Agreements")]
        public string PanelistsAgreements { get; set; }

        [JsonProperty("HCP & Agreements")]
        public string HCPAgreements { get; set; }
        public object IsAdvanceRequired { get; set; }
        public object EventOpen30days { get; set; }
        public object EventWithin7days { get; set; }
        public string CreatedOn { get; set; }

        [JsonProperty("Created Date - Helper")]
        public string CreatedDateHelper { get; set; }
        public string Modified { get; set; }

        [JsonProperty("HON-RBM/BM Approval")]
        public string HONRBMBMApproval { get; set; }

        [JsonProperty("HON-RBM/BM Approval Date")]
        public string HONRBMBMApprovalDate { get; set; }

        [JsonProperty("HON-Compliance Approval")]
        public string HONComplianceApproval { get; set; }

        [JsonProperty("HON-Compliance Approval Date")]
        public string HONComplianceApprovalDate { get; set; }

        [JsonProperty("HON-Finance Accounts Approval")]
        public string HONFinanceAccountsApproval { get; set; }

        [JsonProperty("Finance Rejection Comments")]
        public object FinanceRejectionComments { get; set; }

        [JsonProperty("HON-Finance Accounts Approval Date")]
        public string HONFinanceAccountsApprovalDate { get; set; }

        [JsonProperty("HON-Finance Treasury Approval")]
        public string HONFinanceTreasuryApproval { get; set; }

        [JsonProperty("Finance Treasury Rejection Comments")]
        public object FinanceTreasuryRejectionComments { get; set; }

        [JsonProperty("HON-Finance Treasury Approval Date")]
        public string HONFinanceTreasuryApprovalDate { get; set; }

        [JsonProperty("HON-Marketing Head Approval")]
        public object HONMarketingHeadApproval { get; set; }

        [JsonProperty("HON-Marketing Head Approval Date")]
        public object HONMarketingHeadApprovalDate { get; set; }

        [JsonProperty("HON-Sales Head Approval")]
        public object HONSalesHeadApproval { get; set; }

        [JsonProperty("HON-Sales Head Approval Date")]
        public object HONSalesHeadApprovalDate { get; set; }

        [JsonProperty("HON-Medical Affairs Head Approval")]
        public object HONMedicalAffairsHeadApproval { get; set; }

        [JsonProperty("HON-Medical Affairs Head Approval Date")]
        public object HONMedicalAffairsHeadApprovalDate { get; set; }

        [JsonProperty("5working days")]
        public object _5workingdays { get; set; }

        [JsonProperty("HON-2Workingdays Deviation Approval")]
        public object HON2WorkingdaysDeviationApproval { get; set; }

        [JsonProperty("HON-2Workingdays Deviation Approval Date")]
        public object HON2WorkingdaysDeviationApprovalDate { get; set; }

        [JsonProperty("HON-Lessthan5invitees Deviation Approval")]
        public object HONLessthan5inviteesDeviationApproval { get; set; }

        [JsonProperty("HON-Lessthan5invitees Deviation Approval Date")]
        public object HONLessthan5inviteesDeviationApprovalDate { get; set; }

        [JsonProperty("Finance Accounts Given Details")]
        public string FinanceAccountsGivenDetails { get; set; }

        [JsonProperty("Finance Treasury Given Details")]
        public string FinanceTreasuryGivenDetails { get; set; }

        [JsonProperty("JV No")]
        public object JVNo { get; set; }

        [JsonProperty("JV Date")]
        public object JVDate { get; set; }

        [JsonProperty("PV No")]
        public object PVNo { get; set; }

        [JsonProperty("PV Date")]
        public object PVDate { get; set; }

        [JsonProperty("ActualAmountGreaterThan50%")]
        public object ActualAmountGreaterThan50 { get; set; }

        [JsonProperty("Reporting Manager")]
        public object ReportingManager { get; set; }

        [JsonProperty("1 Up Manager")]
        public object _1UpManager { get; set; }

        [JsonProperty("Sales Head approval")]
        public object SalesHeadapproval { get; set; }

        [JsonProperty("Approval Status")]
        public object ApprovalStatus { get; set; }

        [JsonProperty("Next Approver")]
        public string NextApprover { get; set; }

        [JsonProperty("RBM/BM")]
        public string RBMBM { get; set; }

        [JsonProperty("Sales Head")]
        public object SalesHead { get; set; }

        [JsonProperty("Marketing Head")]
        public object MarketingHead { get; set; }
        public object Finance { get; set; }
        public string Compliance { get; set; }

        [JsonProperty("Finance Treasury")]
        public string FinanceTreasury { get; set; }

        [JsonProperty("Finance Accounts")]
        public string FinanceAccounts { get; set; }

        [JsonProperty("Medical Affairs Head")]
        public object MedicalAffairsHead { get; set; }

        [JsonProperty("Sales Coordinator")]
        public object SalesCoordinator { get; set; }

        [JsonProperty("Finance Head")]
        public object FinanceHead { get; set; }

        [JsonProperty("Cost per participant - Helper")]
        public object CostperparticipantHelper { get; set; }
        public object BrandName { get; set; }

        [JsonProperty("% Allocation")]
        public object Allocation { get; set; }

        [JsonProperty("Project ID")]
        public object ProjectID { get; set; }
        public object Honorarium { get; set; }

        [JsonProperty("Total Invitees")]
        public object TotalInvitees { get; set; }

        [JsonProperty("Total Attendees")]
        public object TotalAttendees { get; set; }

        [JsonProperty("Total Honorarium Amount")]
        public double TotalHonorariumAmount { get; set; }

        [JsonProperty("Total Travel & Accommodation Amount")]
        public double TotalTravelAccommodationAmount { get; set; }

        [JsonProperty("Total Travel Amount")]
        public double TotalTravelAmount { get; set; }

        [JsonProperty("Total Accommodation Amount")]
        public double TotalAccommodationAmount { get; set; }

        [JsonProperty("Total Local Conveyance")]
        public double TotalLocalConveyance { get; set; }

        [JsonProperty("Total Expense")]
        public double TotalExpense { get; set; }

        [JsonProperty("Total Budget")]
        public double TotalBudget { get; set; }

        [JsonProperty("HCP Role")]
        public object HCPRole { get; set; }
        public object Role { get; set; }

        [JsonProperty("Advance Amount")]
        public object AdvanceAmount { get; set; }

        [JsonProperty(" Total Expense BTC")]
        public object TotalExpenseBTC { get; set; }

        [JsonProperty("Total Expense BTE")]
        public object TotalExpenseBTE { get; set; }

        [JsonProperty("Meeting Type")]
        public object MeetingType { get; set; }
        public object HELP { get; set; }
        public object Attachments { get; set; }
    }

    public class EventSettlementPayload
    {
        public string EventTopic { get; set; }
        [JsonProperty("EventId/EventRequestId")]
        public string EventIdEventRequestId { get; set; }
        public string EventType { get; set; }
        public string EventDate { get; set; }
        [JsonProperty("Event End Date")]
        public object EventEndDate { get; set; }
        public string StartTime { get; set; }
        public string EndTime { get; set; }
        public string VenueName { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public double Attended { get; set; }
        public string InviteesParticipated { get; set; }
        public string ExpenseDetails { get; set; }
        public string TotalExpenseDetails { get; set; }
        public string AdvanceDetails { get; set; }
        [JsonProperty("Advance Utilized For Event")]
        public object AdvanceUtilizedForEvent { get; set; }
        [JsonProperty("Pay Back Amount To Company")]
        public object PayBackAmountToCompany { get; set; }
        [JsonProperty("Additional Amount Needed To Pay For Initiator")]
        public object AdditionalAmountNeededToPayForInitiator { get; set; }
        public string CreatedOn { get; set; }
        [JsonProperty("Created Date - Helper")]
        public string CreatedDateHelper { get; set; }
        public string Modified { get; set; }
        public string InitiatorName { get; set; }
        [JsonProperty("Initiator Email")]
        public string InitiatorEmail { get; set; }
        public string IsAdvanceRequired { get; set; }
        [JsonProperty("Event Request Status")]
        public object EventRequestStatus { get; set; }
        [JsonProperty("Honorarium Request Status")]
        public object HonorariumRequestStatus { get; set; }
        [JsonProperty("Post Event Request status")]
        public string PostEventRequeststatus { get; set; }
        [JsonProperty("Post Event Approved Date")]
        public string PostEventApprovedDate { get; set; }
        [JsonProperty("EventSettlement-RBM/BM Approval")]
        public string EventSettlementRBMBMApproval { get; set; }
        [JsonProperty("EventSettlement-RBM/BM Approval Date")]
        public string EventSettlementRBMBMApprovalDate { get; set; }
        [JsonProperty("EventSettlement-SalesHead Approval")]
        public object EventSettlementSalesHeadApproval { get; set; }
        [JsonProperty("EventSettlement-Sales Head Approval Date")]
        public object EventSettlementSalesHeadApprovalDate { get; set; }
        [JsonProperty("EventSettlement-Compliance Approval")]
        public string EventSettlementComplianceApproval { get; set; }
        [JsonProperty("EventSettlement-Compliance Approval Date")]
        public string EventSettlementComplianceApprovalDate { get; set; }

        [JsonProperty("EventSettlement-Finance Account Approval")]

        public object EventSettlementFinanceAccountApproval { get; set; }
        [JsonProperty("EventSettlement-Finance Account Comments")]
        public object EventSettlementFinanceAccountComments { get; set; }
        [JsonProperty("EventSettlement-Finance Account Approval Date")]
        public object EventSettlementFinanceAccountApprovalDate { get; set; }
        [JsonProperty("EventSettlement-Finance Treasury Approval")]
        public string EventSettlementFinanceTreasuryApproval { get; set; }
        [JsonProperty("EventSettlement-Finance Treasury Comments")]
        public string EventSettlementFinanceTreasuryComments { get; set; }
        [JsonProperty("EventSettlement-Finance Treasury Approval Date")]
        public string EventSettlementFinanceTreasuryApprovalDate { get; set; }
        [JsonProperty("EventSettlement-Marketing Head Approval")]
        public object EventSettlementMarketingHeadApproval { get; set; }
        [JsonProperty("EventSettlement-Marketing Head Approval Date")]
        public object EventSettlementMarketingHeadApprovalDate { get; set; }
        [JsonProperty("EventSettlement-Medical Affairs Head Approval")]
        public object EventSettlementMedicalAffairsHeadApproval { get; set; }
        [JsonProperty("EventSettlement-Medical Affairs Head Approval Date")]
        public object EventSettlementMedicalAffairsHeadApprovalDate { get; set; }
        [JsonProperty("Deviation Status")]
        public string DeviationStatus { get; set; }
        [JsonProperty("Is All Deviations Approved?")]
        public object IsAllDeviationsApproved { get; set; }

        [JsonProperty("POST- Beyond30Days Deviation Approval")]
        public object POSTBeyond30DaysDeviationApproval { get; set; }

        [JsonProperty("post 45 days approved")]
        public object post45daysapproved { get; set; }

        [JsonProperty("POST- Beyond30Days Deviation Approval Date")]
        public object POSTBeyond30DaysDeviationApprovalDate { get; set; }

        [JsonProperty("POST-LessThan5Invitees Deviation Approval")]
        public object POSTLessThan5InviteesDeviationApproval { get; set; }

        [JsonProperty("Post <5 Invitees Approved")]
        public object Post5InviteesApproved { get; set; }

        [JsonProperty("POST-Deviation Costperpaxabove1500 Approval ")]
        public object POSTDeviationCostperpaxabove1500Approval { get; set; }

        [JsonProperty("Post CostperPax Approved")]
        public object PostCostperPaxApproved { get; set; }

        [JsonProperty("POST-Deviation Costperpaxabove1500 Approval Date")]
        public object POSTDeviationCostperpaxabove1500ApprovalDate { get; set; }

        [JsonProperty("POST-Deviation Change in venue Approval")]
        public object POSTDeviationChangeinvenueApproval { get; set; }

        [JsonProperty("Post ChangeInVenue Approved")]
        public object PostChangeInVenueApproved { get; set; }

        [JsonProperty("POST-Deviation Change in topic Approval")]
        public object POSTDeviationChangeintopicApproval { get; set; }

        [JsonProperty("Post ChangeInTopic Approved")]
        public object PostChangeInTopicApproved { get; set; }

        [JsonProperty("POST-Deviation Change in speaker Approval")]
        public object POSTDeviationChangeinspeakerApproval { get; set; }

        [JsonProperty("Post ChangeInSpeaker Approved")]
        public object PostChangeInSpeakerApproved { get; set; }

        [JsonProperty("POST-Deviation Attendees not captured Approval")]
        public object POSTDeviationAttendeesnotcapturedApproval { get; set; }

        [JsonProperty("Post AttendeesNotCaptured Approved")]
        public object PostAttendeesNotCapturedApproved { get; set; }

        [JsonProperty("POST-Deviation Speaker not captured  Approval")]
        public object POSTDeviationSpeakernotcapturedApproval { get; set; }

        [JsonProperty("Post SpeakerNotCaptured Approved")]
        public object PostSpeakerNotCapturedApproved { get; set; }

        [JsonProperty("POST-Deviation Other Deviation Approval")]
        public object POSTDeviationOtherDeviationApproval { get; set; }

        [JsonProperty("Post OtherDeviation Approved")]
        public object PostOtherDeviationApproved { get; set; }

        [JsonProperty("EventSettlement - Deviation Date")]
        public object EventSettlementDeviationDate { get; set; }

        [JsonProperty("EventSettlement - Deviation Approval")]
        public object EventSettlementDeviationApproval { get; set; }

        [JsonProperty("EventSettlement - Deviation Approval Date")]
        public object EventSettlementDeviationApprovalDate { get; set; }

        [JsonProperty("Finance Deviation Pending")]
        public object FinanceDeviationPending { get; set; }

        [JsonProperty("Sales Deviation Pending")]
        public object SalesDeviationPending { get; set; }

        [JsonProperty("Honorarium Submitted?")]
        public object HonorariumSubmitted { get; set; }
        public object Honorariumamount { get; set; }

        [JsonProperty("IsItincludingGST?")]
        public object IsItincludingGST { get; set; }
        public object AgreementAmount { get; set; }

        [JsonProperty("JV No")]
        public object JVNo { get; set; }

        [JsonProperty("JV Date")]
        public object JVDate { get; set; }

        [JsonProperty("PV No")]
        public object PVNo { get; set; }

        [JsonProperty("PV Date")]
        public object PVDate { get; set; }

        [JsonProperty("PostEventSubmitted?")]
        public string PostEventSubmitted { get; set; }
        public object PostEventBTCExpense { get; set; }
        public object PostEventBTEExpense { get; set; }

        [JsonProperty("ActualAmountGreaterThan50%")]
        public object ActualAmountGreaterThan50Per { get; set; }

        [JsonProperty("Reporting Manager")]
        public object ReportingManager { get; set; }

        [JsonProperty("1 Up Manager")]
        public object _1UpManager { get; set; }
        public string Brands { get; set; }
        public string Panelists { get; set; }
        public string HCP { get; set; }
        public string SlideKits { get; set; }
        public string Invitees { get; set; }
        public string Expenses { get; set; }

        [JsonProperty("Indications Done")]
        public object IndicationsDone { get; set; }

        [JsonProperty("Total Invitees")]
        public double TotalInvitees { get; set; }

        [JsonProperty("Total Attendees")]
        public double TotalAttendees { get; set; }

        [JsonProperty("Approval Status")]
        public object ApprovalStatus { get; set; }

        [JsonProperty("Next Approver")]
        public object NextApprover { get; set; }
        public string Finance { get; set; }

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

        [JsonProperty("Sales Coordinator")]
        public object SalesCoordinator { get; set; }

        [JsonProperty("Finance Head")]
        public string FinanceHead { get; set; }

        [JsonProperty("Total Honorarium Amount")]
        public double TotalHonorariumAmount { get; set; }

        [JsonProperty("Total Travel & Accommodation Amount")]
        public double TotalTravelAccommodationAmount { get; set; }

        [JsonProperty("Total Travel Amount")]
        public double TotalTravelAmount { get; set; }

        [JsonProperty("Total Accommodation Amount")]
        public double TotalAccommodationAmount { get; set; }

        [JsonProperty("Total Local Conveyance")]
        public double TotalLocalConveyance { get; set; }

        [JsonProperty("Total Expense")]
        public double TotalExpense { get; set; }

        [JsonProperty("Total Budget")]
        public double TotalBudget { get; set; }

        [JsonProperty("Total Actual")]
        public double TotalActual { get; set; }

        [JsonProperty("50% Helper")]
        public string _50Helper { get; set; }

        [JsonProperty("Cost per participant - Helper")]
        public object CostperparticipantHelper { get; set; }

        [JsonProperty("BTC exceeds 50% of budget")]
        public object BTCexceeds50ofbudget { get; set; }

        [JsonProperty("Finance Accounts Given Details")]
        public string FinanceAccountsGivenDetails { get; set; }

        [JsonProperty("Finance Treasury Given Details")]
        public string FinanceTreasuryGivenDetails { get; set; }
        public string Role { get; set; }

        [JsonProperty("Advance Amount")]
        public object AdvanceAmount { get; set; }

        [JsonProperty("Total Expense BTC")]
        public object TotalExpenseBTC { get; set; }

        [JsonProperty("Total Expense BTE")]
        public object TotalExpenseBTE { get; set; }

        [JsonProperty("Helper Finance treasury trigger(BTE)")]
        public string HelperFinancetreasurytriggerBTE { get; set; }

        [JsonProperty("Class III Event Code")]
        public object ClassIIIEventCode { get; set; }

        [JsonProperty("Meeting Type")]
        public object MeetingType { get; set; }

        [JsonProperty("Sponsorship Society Name")]
        public object SponsorshipSocietyName { get; set; }

        [JsonProperty("Venue Country")]
        public object VenueCountry { get; set; }

        [JsonProperty("Total HCP Registration Amount")]
        public object TotalHCPRegistrationAmount { get; set; }

        [JsonProperty("Medical Utility Type")]
        public object MedicalUtilityType { get; set; }

        [JsonProperty("Medical Utility Description")]
        public object MedicalUtilityDescription { get; set; }

        [JsonProperty("Valid To")]
        public object ValidTo { get; set; }

        [JsonProperty("Valid From")]
        public object ValidFrom { get; set; }
    }

    public class DeviationDataPayload
    {
        [JsonProperty("Deviation ID")]
        public string DeviationID { get; set; }

        [JsonProperty("EventId/EventRequestId")]
        public string EventIdEventRequestId { get; set; }

        [JsonProperty("Event Topic")]
        public string EventTopic { get; set; }

        [JsonProperty("MIS Code")]
        public object MISCode { get; set; }

        [JsonProperty("HCP Name")]
        public object HCPName { get; set; }

        [JsonProperty("Honorarium Amount")]
        public string HonorariumAmount { get; set; }

        [JsonProperty("Travel & Accommodation Amount")]
        public string TravelAccommodationAmount { get; set; }

        [JsonProperty("Other Expenses")]
        public string OtherExpenses { get; set; }

        [JsonProperty("Deviation Type")]
        public string DeviationType { get; set; }

        [JsonProperty("Event Level")]
        public string EventLevel { get; set; }

        [JsonProperty("Outstanding Events")]
        public string OutstandingEvents { get; set; }
        public string InitiatorName { get; set; }

        [JsonProperty("Initiator Email")]
        public string InitiatorEmail { get; set; }
        public string CreatedOn { get; set; }
        public string Modified { get; set; }
        public string EventType { get; set; }
        public string EventDate { get; set; }

        [JsonProperty("Event End Date")]
        public object EventEndDate { get; set; }
        public string StartTime { get; set; }
        public string EndTime { get; set; }
        public string VenueName { get; set; }
        public string City { get; set; }
        public object IsAdvanceRequired { get; set; }
        public string State { get; set; }

        [JsonProperty("Sales Head approval")]
        public string SalesHeadapproval { get; set; }

        [JsonProperty("Sales Head approval Date")]
        public object SalesHeadapprovalDate { get; set; }

        [JsonProperty("Finance Head Approval")]
        public object FinanceHeadApproval { get; set; }

        [JsonProperty("Finance Head Approval Date")]
        public object FinanceHeadApprovalDate { get; set; }

        [JsonProperty("Approval Status")]
        public object ApprovalStatus { get; set; }

        [JsonProperty("HON-5Workingdays Deviation Date Trigger")]
        public object HON5WorkingdaysDeviationDateTrigger { get; set; }

        [JsonProperty("HON-5Workingdays Deviation Approval")]
        public object HON5WorkingdaysDeviationApproval { get; set; }

        [JsonProperty("HON-5Workingdays Deviation Approval Date")]
        public object HON5WorkingdaysDeviationApprovalDate { get; set; }

        [JsonProperty("POST- Beyond45Days Deviation Date Trigger")]
        public object POSTBeyond45DaysDeviationDateTrigger { get; set; }

        [JsonProperty("POST- Beyond45Days Deviation Approval")]
        public object POSTBeyond45DaysDeviationApproval { get; set; }

        [JsonProperty("POST- Beyond45Days Deviation Approval Date")]
        public object POSTBeyond45DaysDeviationApprovalDate { get; set; }

        [JsonProperty("POST-Lessthan5Invitees Deviation Trigger")]
        public object POSTLessthan5InviteesDeviationTrigger { get; set; }

        [JsonProperty("POST-Lessthan5invitees Deviation Approval")]
        public object POSTLessthan5inviteesDeviationApproval { get; set; }

        [JsonProperty("POST-Lessthan5invitees Deviation Approval Date")]
        public object POSTLessthan5inviteesDeviationApprovalDate { get; set; }

        [JsonProperty("POST-Deviation Costperpaxabove1500 Trigger")]
        public object POSTDeviationCostperpaxabove1500Trigger { get; set; }

        [JsonProperty("POST-Deviation Costperpaxabove1500 Approval ")]
        public object POSTDeviationCostperpaxabove1500Approval { get; set; }

        [JsonProperty("POST-Deviation Costperpaxabove1500 Approval Date")]
        public object POSTDeviationCostperpaxabove1500ApprovalDate { get; set; }

        [JsonProperty("POST-Deviation Costperpaxabove2250 Trigger")]
        public object POSTDeviationCostperpaxabove2250Trigger { get; set; }

        [JsonProperty("POST-Deviation Costperpaxabove2250 Approval ")]
        public object POSTDeviationCostperpaxabove2250Approval { get; set; }

        [JsonProperty("POST-Deviation Costperpaxabove2250 Approval Date")]
        public object POSTDeviationCostperpaxabove2250ApprovalDate { get; set; }

        [JsonProperty("POST-Deviation Change in venue trigger")]
        public object POSTDeviationChangeinvenuetrigger { get; set; }

        [JsonProperty("POST-Deviation Change in venue Approval")]
        public object POSTDeviationChangeinvenueApproval { get; set; }

        [JsonProperty("POST-Deviation Change in venue Approval Date")]
        public object POSTDeviationChangeinvenueApprovalDate { get; set; }

        [JsonProperty("POST-Deviation Change in topic trigger")]
        public object POSTDeviationChangeintopictrigger { get; set; }

        [JsonProperty("POST-Deviation Change in topic Approval")]
        public object POSTDeviationChangeintopicApproval { get; set; }

        [JsonProperty("POST-Deviation Change in topic Approval Date")]
        public object POSTDeviationChangeintopicApprovalDate { get; set; }

        [JsonProperty("POST-Deviation Change in speaker trigger")]
        public object POSTDeviationChangeinspeakertrigger { get; set; }

        [JsonProperty("POST-Deviation Change in speaker Approval")]
        public object POSTDeviationChangeinspeakerApproval { get; set; }

        [JsonProperty("POST-Deviation Change in speaker Approval Date")]
        public object POSTDeviationChangeinspeakerApprovalDate { get; set; }

        [JsonProperty("POST-Deviation Attendees not captured trigger")]
        public object POSTDeviationAttendeesnotcapturedtrigger { get; set; }

        [JsonProperty("POST-Deviation Attendees not captured Approval")]
        public object POSTDeviationAttendeesnotcapturedApproval { get; set; }

        [JsonProperty("POST-Deviation Attendees not captured App Date")]
        public object POSTDeviationAttendeesnotcapturedAppDate { get; set; }

        [JsonProperty("POST-Deviation Speaker not captured trigger")]
        public object POSTDeviationSpeakernotcapturedtrigger { get; set; }

        [JsonProperty("POST-Deviation Speaker not captured  Approval")]
        public object POSTDeviationSpeakernotcapturedApproval { get; set; }

        [JsonProperty("POST-Deviation Speaker not captured  App Date")]
        public object POSTDeviationSpeakernotcapturedAppDate { get; set; }

        [JsonProperty("POST-Deviation Other Deviation Trigger")]
        public object POSTDeviationOtherDeviationTrigger { get; set; }

        [JsonProperty("Other Deviation Type")]
        public object OtherDeviationType { get; set; }

        [JsonProperty("POST-Deviation Other Deviation Approval")]
        public object POSTDeviationOtherDeviationApproval { get; set; }

        [JsonProperty("POST-Deviation Other Deviation Approval Date")]
        public object POSTDeviationOtherDeviationApprovalDate { get; set; }

        [JsonProperty("HCP exceeds 1,00,000 Trigger")]
        public object HCPexceeds100000Trigger { get; set; }

        [JsonProperty("HCP exceeds 1,00,000 FH Approval")]
        public object HCPexceeds100000FHApproval { get; set; }

        [JsonProperty("HCP exceeds 1,00,000 FH Approval Date")]
        public object HCPexceeds100000FHApprovalDate { get; set; }

        [JsonProperty("HCP exceeds 5,00,000 Trigger")]
        public object HCPexceeds500000Trigger { get; set; }

        [JsonProperty("HCP exceeds 5,00,000 Trigger FH Approval")]
        public object HCPexceeds500000TriggerFHApproval { get; set; }

        [JsonProperty("HCP exceeds 5,00,000 Trigger Approval Date")]
        public object HCPexceeds500000TriggerApprovalDate { get; set; }

        [JsonProperty("HCP Honorarium 6,00,000 Exceeded Trigger")]
        public object HCPHonorarium600000ExceededTrigger { get; set; }

        [JsonProperty("HCP Honorarium 6,00,000 Exceeded Approval")]
        public object HCPHonorarium600000ExceededApproval { get; set; }

        [JsonProperty("HCP Honorarium 6,00,000 Exceeded Approval Date")]
        public object HCPHonorarium600000ExceededApprovalDate { get; set; }

        [JsonProperty("Trainer Honorarium 12,00,000 Exceeded Trigger")]
        public object TrainerHonorarium1200000ExceededTrigger { get; set; }

        [JsonProperty("Trainer Honorarium 12,00,000 Exceeded Approval")]
        public object TrainerHonorarium1200000ExceededApproval { get; set; }

        [JsonProperty("Trainer Honorarium 12,00,000 Exceeded Approval Dat")]
        public object TrainerHonorarium1200000ExceededApprovalDate { get; set; }

        [JsonProperty("Travel/Accomodation 3,00,000 Exceeded Trigger")]
        public object TravelAccomodation300000ExceededTrigger { get; set; }

        [JsonProperty("Travel/Accomodation 3,00,000 Exceeded Approval")]
        public object TravelAccomodation300000ExceededApproval { get; set; }

        [JsonProperty("Travel/Accomodation 3,00,000 Exceeded Approval Dat")]
        public object TravelAccomodation300000ExceededApprovalDate { get; set; }

        [JsonProperty("Brochure/Request Letter Uploaded in 2 days Trigger")]
        public object BrochureRequestLetterUploadedin2daysTrigger { get; set; }

        [JsonProperty("Brochure/Request Uploaded in 2 days FH Approval")]
        public object BrochureRequestUploadedin2daysFHApproval { get; set; }

        [JsonProperty("Brochure/Request in 2 days  Approval Date")]
        public object BrochureRequestin2daysApprovalDate { get; set; }

        [JsonProperty("Cost per participant  INR 2250 Trigger ")]
        public object CostperparticipantINR2250Trigger { get; set; }

        [JsonProperty("Cost per participant INR 2250 FH Approval")]
        public object CostperparticipantINR2250FHApproval { get; set; }

        [JsonProperty("Cost per participant INR 2250 Approval Date")]
        public object CostperparticipantINR2250ApprovalDate { get; set; }

        [JsonProperty("YTD Spend on Speaker Trigger")]
        public object YTDSpendonSpeakerTrigger { get; set; }

        [JsonProperty("YTD Spend on Speaker FH Approval")]
        public object YTDSpendonSpeakerFHApproval { get; set; }

        [JsonProperty("YTD Spend on Speaker Approval Date")]
        public object YTDSpendonSpeakerApprovalDate { get; set; }

        [JsonProperty("amount is with applicable limit Trigger")]
        public object amountiswithapplicablelimitTrigger { get; set; }

        [JsonProperty("amount is with applicable limit FH Approval")]
        public object amountiswithapplicablelimitFHApproval { get; set; }

        [JsonProperty("amount is with applicable limit FH Approval Date")]
        public object amountiswithapplicablelimitFHApprovalDate { get; set; }
        public string EventOpen45days { get; set; }
        public object EventOpenSalesHeadApproval { get; set; }

        [JsonProperty("EventOpenSalesHeadApproval Date")]
        public object EventOpenSalesHeadApprovalDate { get; set; }
        public object EventWithin5days { get; set; }

        [JsonProperty("5daysSalesHeadApproval")]
        public object _5daysSalesHeadApproval { get; set; }

        [JsonProperty("5daysSalesHeadApproval date")]
        public object _5daysSalesHeadApprovaldate { get; set; }

        [JsonProperty("PRE-F&B Expense Excluding Tax")]
        public object PREFBExpenseExcludingTax { get; set; }

        [JsonProperty("PRE-F&B Expense Excluding Tax Approval")]
        public object PREFBExpenseExcludingTaxApproval { get; set; }

        [JsonProperty("PRE-F&B Expense Excluding Tax Approval Date")]
        public object PREFBExpenseExcludingTaxApprovalDate { get; set; }

        [JsonProperty("Honorarium Submitted?")]
        public object HonorariumSubmitted { get; set; }

        [JsonProperty("PostEventSubmitted?")]
        public object PostEventSubmitted { get; set; }
        public object PostEventBTCExpense { get; set; }
        public object PostEventBTEExpense { get; set; }

        [JsonProperty("ActualAmountGreaterThan50%")]
        public object ActualAmountGreaterThan50 { get; set; }

        [JsonProperty("1 Up Manager")]
        public object _1UpManager { get; set; }

        [JsonProperty("Reporting Manager")]
        public object ReportingManager { get; set; }
        public object Brands { get; set; }
        public object Panelists { get; set; }
        public object SlideKits { get; set; }
        public object Invitees { get; set; }
        public object Expenses { get; set; }

        [JsonProperty("Total Invitees")]
        public object TotalInvitees { get; set; }

        [JsonProperty("Total Attendees")]
        public object TotalAttendees { get; set; }
        public object Finance { get; set; }

        [JsonProperty("RBM/BM")]
        public object RBMBM { get; set; }

        [JsonProperty("Sales Head Status")]
        public object SalesHeadStatus { get; set; }

        [JsonProperty("Sales Head")]
        public string SalesHead { get; set; }

        [JsonProperty("Marketing Head")]
        public object MarketingHead { get; set; }
        public object Compliance { get; set; }

        [JsonProperty("Finance Treasury")]
        public object FinanceTreasury { get; set; }

        [JsonProperty("Finance Accounts")]
        public object FinanceAccounts { get; set; }

        [JsonProperty("Medical Affairs Head ")]
        public object MedicalAffairsHead { get; set; }

        [JsonProperty("Finance Head")]
        public string FinanceHead { get; set; }

        [JsonProperty("Total Honorarium Amount")]
        public object TotalHonorariumAmount { get; set; }

        [JsonProperty("Total Travel & Accomodation Spend")]
        public object TotalTravelAccomodationSpend { get; set; }

        [JsonProperty("Total Travel Amount")]
        public object TotalTravelAmount { get; set; }

        [JsonProperty("Total Travel & Accommodation Amount")]
        public object TotalTravelAccommodationAmount { get; set; }

        [JsonProperty("Total Local Conveyance")]
        public object TotalLocalConveyance { get; set; }

        [JsonProperty("Total Expense for Invitee")]
        public object TotalExpenseforInvitee { get; set; }

        [JsonProperty("Total Budget")]
        public object TotalBudget { get; set; }

        [JsonProperty("Cost per participant - Helper")]
        public object CostperparticipantHelper { get; set; }

        [JsonProperty("POST-Deviation Excluding GST?")]
        public object POSTDeviationExcludingGST { get; set; }

        [JsonProperty("Finance Head approval2")]
        public object FinanceHeadapproval2 { get; set; }

        [JsonProperty("Finance Head approval3")]
        public object FinanceHeadapproval3 { get; set; }
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
        public double AggregateHonorariumLimit { get; set; }

        [JsonProperty("Aggregate Accommodataion Limit")]
        public double AggregateAccommodataionLimit { get; set; }

        [JsonProperty("Aggregate Honorarium Spent")]
        public double AggregateHonorariumSpent { get; set; }

        [JsonProperty("Aggregate Accommodation Spent")]
        public double AggregateAccommodationSpent { get; set; }

        [JsonProperty("Aggregate Spent on Medical Utility")]
        public double AggregateSpentonMedicalUtility { get; set; }

        [JsonProperty("Aggregate Spent as HCP Consultant")]
        public double AggregateSpentasHCPConsultant { get; set; }

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
        public double AggregatespendonHonorariumTrainer { get; set; }

        [JsonProperty("Aggregate Limit on Accomodation")]
        public double AggregateLimitonAccomodation { get; set; }

        [JsonProperty("Aggregate Honorarium Spent ")]
        public double AggregateHonorariumSpent { get; set; }

        [JsonProperty("Aggregate spend on Accomodation")]
        public double AggregatespendonAccomodation { get; set; }

        [JsonProperty("Aggregate Spent on Medical Utility")]
        public double AggregateSpentonMedicalUtility { get; set; }

        [JsonProperty("Aggregate Spent as HCP Consultant")]
        public double AggregateSpentasHCPConsultant { get; set; }
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
        public string MisCode { get; set; }

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

        //[JsonProperty("Finance Checker-1")]
        //public string FinanceChecker1 { get; set; }

        [JsonProperty("Finance Checker Approval")]
        public string FinanceCheckerApproval { get; set; }

        [JsonProperty("Finance Checker Approval Date")]
        public string FinanceCheckerApprovalDate { get; set; }
        public string Requestor { get; set; }

        [JsonProperty("Initiator Name")]
        public string InitiatorName { get; set; }

        [JsonProperty("Initiator Email")]
        public string InitiatorEmail { get; set; }


        [JsonProperty("Finance Checker")]
        public object FinanceChecker { get; set; }
    }

    public class SpeakerCodeCreationGetPayload
    {
        public string SpeakerName { get; set; }
        public string Speaker_Code { get; set; }

        [JsonProperty("Speaker Code")]
        public string SpeakerCode { get; set; }
        public string MisCode { get; set; }
        public string Division { get; set; }
        public string Speciality { get; set; }
        public string Qualification { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string Country { get; set; }

        [JsonProperty("Contact Number")]
        public string ContactNumber { get; set; }

        [JsonProperty("Speaker Type")]
        public string SpeakerType { get; set; }

        [JsonProperty("Speaker Category")]
        public string SpeakerCategory { get; set; }

        [JsonProperty("Speaker Criteria")]
        public string SpeakerCriteria { get; set; }

        [JsonProperty("Speaker Criteria Details")]
        public string SpeakerCriteriaDetails { get; set; }
        public string Created { get; set; }

        [JsonProperty("Created Date - Helper")]
        public string CreatedDateHelper { get; set; }

        [JsonProperty("Sales Alert Trigger")]
        public string SalesAlertTrigger { get; set; }

        [JsonProperty("Sales Head Approval")]
        public string SalesHeadApproval { get; set; }

        [JsonProperty("Sales Head Approval Date")]
        public string SalesHeadApprovalDate { get; set; }

        [JsonProperty("Medical Affairs Alert Trigger")]
        public string MedicalAffairsAlertTrigger { get; set; }

        [JsonProperty("Medical Affairs Head Approval")]
        public string MedicalAffairsHeadApproval { get; set; }

        [JsonProperty("Medical Affairs Head Approval Date")]
        public string MedicalAffairsHeadApprovalDate { get; set; }

        [JsonProperty("Sales Head")]
        public string SalesHead { get; set; }

        [JsonProperty("Medical Affairs Head")]
        public string MedicalAffairsHead { get; set; }

        [JsonProperty("Initiator Name")]
        public string InitiatorName { get; set; }

        [JsonProperty("Initiator Email")]
        public string InitiatorEmail { get; set; }
    }

    public class TrainerCodeCreationGetPayload
    {
        public string TrainerName { get; set; }
        public string Number { get; set; }

        [JsonProperty("TrainerTybe - Shortcode")]
        public string TrainerTybeShortcode { get; set; }

        [JsonProperty("Trainer Code")]
        public string TrainerCode { get; set; }

        [JsonProperty("Trainer Brand")]
        public string TrainerBrand { get; set; }

        [JsonProperty("Trainer Type")]
        public string TrainerType { get; set; }
        public string MisCode { get; set; }
        public string Division { get; set; }
        public string Speciality { get; set; }
        public string Qualification { get; set; }
        public object Address { get; set; }
        public object City { get; set; }
        public object State { get; set; }
        public object Country { get; set; }

        [JsonProperty("Contact Number")]
        public object ContactNumber { get; set; }

        [JsonProperty("Trained by")]
        public object Trainedby { get; set; }

        [JsonProperty("Trainer CV")]
        public object TrainerCV { get; set; }

        [JsonProperty("Trainer certificate")]
        public object Trainercertificate { get; set; }

        [JsonProperty("Trained on")]
        public object Trainedon { get; set; }

        [JsonProperty("Trainer Category")]
        public object TrainerCategory { get; set; }

        [JsonProperty("Trainer Criteria")]
        public object TrainerCriteria { get; set; }

        [JsonProperty("Trainer Criteria Details")]
        public object TrainerCriteriaDetails { get; set; }
        public string Created { get; set; }

        [JsonProperty("Created Date - Helper")]
        public string CreatedDateHelper { get; set; }

        [JsonProperty("Sales Alert Trigger")]
        public string SalesAlertTrigger { get; set; }

        [JsonProperty("Medical Affairs Alert Trigger")]
        public string MedicalAffairsAlertTrigger { get; set; }

        [JsonProperty("Sales Head")]
        public string SalesHead { get; set; }

        [JsonProperty("Sales Head Approval")]
        public object SalesHeadApproval { get; set; }

        [JsonProperty("Sales Head Approval Date")]
        public string SalesHeadApprovalDate { get; set; }

        [JsonProperty("Medical Affairs Head")]
        public string MedicalAffairsHead { get; set; }

        [JsonProperty("Medical Affairs Head Approval")]
        public object MedicalAffairsHeadApproval { get; set; }

        [JsonProperty("Medical Affairs Head Approval Date")]
        public object MedicalAffairsHeadApprovalDate { get; set; }

        [JsonProperty("Initiator Name")]
        public object InitiatorName { get; set; }

        [JsonProperty("Initiator Email")]
        public object InitiatorEmail { get; set; }
    }

    public class VendorCodeCreationGetPayload
    {
        public string VendorId { get; set; }
        public string VendorAccount { get; set; }
        public string MisCode { get; set; }
        public string BeneficiaryName { get; set; }
        public string PanCardName { get; set; }
        public string PanNumber { get; set; }
        public string BankAccountNumber { get; set; }

        [JsonProperty("Bank Name")]
        public object BankName { get; set; }
        public string IfscCode { get; set; }

        [JsonProperty("Swift Code")]
        public string SwiftCode { get; set; }

        [JsonProperty("IBN Number")]
        public string IBNNumber { get; set; }

        [JsonProperty("Email ")]
        public string Email { get; set; }

        [JsonProperty("Pancard Document")]
        public string PancardDocument { get; set; }

        [JsonProperty("Cheque Document")]
        public string ChequeDocument { get; set; }

        [JsonProperty("Tax Residence Certificate")]
        public string TaxResidenceCertificate { get; set; }

        [JsonProperty("Initiator Name")]
        public object InitiatorName { get; set; }

        [JsonProperty("Initiator Email")]
        public object InitiatorEmail { get; set; }

        [JsonProperty("Tax Residence Certificate Date")]
        public object TaxResidenceCertificateDate { get; set; }

        [JsonProperty("Finance Checker")]
        public string FinanceChecker { get; set; }

        [JsonProperty("Finance Checker  approval")]
        public string FinanceCheckerapproval { get; set; }

        [JsonProperty("Finance Checker Approval Date")]
        public object FinanceCheckerApprovalDate { get; set; }
    }

}
