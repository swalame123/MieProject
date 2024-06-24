using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IndiaEvents.Models.Models.MasterSheets
{
    public class UserDetails
    {


        public string Username { get; set; }
        public string Email { get; set; }
        public string Password { get; set; }
        public string Role { get; set; }
        public string ReportingManager { get; set; }
        public string FirstLevelManager { get; set; }
        public string RBM_BM { get; set; }
        public string SalesHead { get; set; }
        public string FinanceHead { get; set; }
        public string MarketingHead { get; set; }
        public string Compliance { get; set; }
        public string MedicalAffairsHead { get; set; }
        public string FinanceTreasury { get; set; }
        public string FinanceChecker { get; set; }
        public string FinanceAccounts { get; set; }
        public string SalesCoordinator { get; set; }
        public string MarketingCoordinator { get; set; }
        public string IsActive { get; set; }
    }
}
