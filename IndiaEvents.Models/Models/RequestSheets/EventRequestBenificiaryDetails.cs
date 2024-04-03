using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IndiaEvents.Models.Models.RequestSheets
{
    public class EventRequestBenificiaryDetails
    {

        public string? Currency { get; set; }
        public string? EnterCurrencyType { get; set; }
        public string? BenificiaryName { get; set; }
        public string? BankAccountNumber { get; set; }
        public string? BankName { get; set; }
        public string? NameasPerPAN { get; set; }
        public string? PANCardNumber { get; set; }
        public string? IFSCCode { get; set; }
        public string? IbnNumber  { get; set; }
        public string? SwiftCode { get; set; }
        public DateTime? TaxResidenceCertificateDate { get; set; }
        public string? EmailID { get; set; }
    }
}
