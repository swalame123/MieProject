using IndiaEventsWebApi.Models.MasterSheets.CodeCreation;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IndiaEvents.Models.Data
{

    public class SampleDbContext : DbContext
    {
        public SampleDbContext(DbContextOptions options) : base(options)
        {
        }
        public DbSet<VendorCodeGeneration> vendorCodeGenerations { get; set; }
    }
}
