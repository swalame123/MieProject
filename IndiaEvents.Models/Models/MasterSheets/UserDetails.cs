using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IndiaEvents.Models.Models.MasterSheets
{
    public class UserDetails
    {


        public string unique_name { get; set; }
        public string email { get; set; }
        
        public string role { get; set; }
        public string reportingmanager { get; set; }
        public string firstLevelManager { get; set; }
        public string RBM_BM { get; set; }
        public string SalesHead { get; set; }
        public string FinanceHead { get; set; }
        public string MarketingHead { get; set; }
        public string ComplianceHead { get; set; }
        public string MedicalAffairsHead { get; set; }
        public string FinanceTreasury { get; set; }
        public string FinanceChecker { get; set; }
        public string FinanceAccounts { get; set; }
        public string SalesCoordinator { get; set; }
        public string MarketingCoordinator { get; set; }
       
    }
}
