using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IndiaEvents.Models.Models.RequestSheets
{

    public class EventRequestDeviationsData
    {

        public double HonorariumAmountExcludingTax { get; set; }
        //public int AccommodationAmountExcludingTax { get; set; }
        public double TravelorAccomodationAmountExcludingTax { get; set; }
        public double OtherExpenseAmountExcludingTax { get; set; }
        //public int FoodAndBeveragesAmountExcludingTax { get; set; }
        //public int AgregateSpendAmountExcludingTax { get; set; }
        public string? DeviationFile { get; set; }
        public string? HcpName { get; set; }
        public string? MisCode { get; set; }
    }
}
