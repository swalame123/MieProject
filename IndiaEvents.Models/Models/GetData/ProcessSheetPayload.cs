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


}
