using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IndiaEvents.Models.Models.EventTypeSheets
{
    public class ApprovalAndRejectionFlowInPreEvent
    {
        public string? EventId { get; set; }
        public string? RBMStatus { get; set; }
        public string? SalesHeadStatus { get; set; }
        public string? MarketingHeadStatus { get; set; }
        public string? FinanceTreasuryStatus { get; set; }
        public string? MedicalAffairsHeadStatus { get; set; }
        public string? ComplianceStatus { get; set; }
        public string? Comments { get; set; }
    }
    public class ApprovalAndRejectionFlowInDeviation
    {
        public string? EventId { get; set; }
        public string? SalesHeadStatus { get; set; }
        public string? FinanceHeadStatus { get; set; }
        public string? Comments { get; set; }
    }



    public class ApprovalAndRejectionFlowInHonorarium
    {
        public string? EventId { get; set; }
        public string? RBMStatus { get; set; }
        public string? SalesHeadStatus { get; set; }
        public string? MarketingHeadStatus { get; set; }
        public string? MedicalAffairsHeadStatus { get; set; }
        public string? ComplianceStatus { get; set; }
        public string? Comments { get; set; }
    }
    public class ApprovalAndRejectionFlowInPostSettlement
    {
        public string? EventId { get; set; }
        public string? RBMStatus { get; set; }
        public string? SalesHeadStatus { get; set; }
        public string? MarketingHeadStatus { get; set; }
        public string? MedicalAffairsHeadStatus { get; set; }
        public string? ComplianceStatus { get; set; }
        public string? Comments { get; set; }
    }
}
