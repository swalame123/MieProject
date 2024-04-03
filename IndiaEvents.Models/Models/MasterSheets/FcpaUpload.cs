using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IndiaEvents.Models.Models.MasterSheets
{
    public class FcpaUpload
    {
        public string? SelectedType { get; set; }
        public string? UploadFile { get; set; }
        public DateTime? FcpaDate { get; set; }
        public string? MisCode { get; set; }
        public string? NA_Id { get; set; }
        
    }
}
