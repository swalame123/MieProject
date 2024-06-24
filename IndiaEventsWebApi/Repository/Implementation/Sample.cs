using IndiaEvents.Models.Data;
using IndiaEventsWebApi.Models.MasterSheets.CodeCreation;
using IndiaEventsWebApi.Repository.Interface;

namespace IndiaEventsWebApi.Repository.Implementation
{
    public class Sample : ISample
    {
        private readonly SampleDbContext _context;
        public Sample(SampleDbContext context)
        {
            this._context = context;
        }
        public async Task<VendorCodeGeneration> CreateAsync(VendorCodeGeneration vendorCodeGeneration)
        {

            await _context.vendorCodeGenerations.AddAsync(vendorCodeGeneration);
            await _context.SaveChangesAsync();

            return vendorCodeGeneration;
        }
    }
}
