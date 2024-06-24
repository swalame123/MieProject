using IndiaEventsWebApi.Models.MasterSheets.CodeCreation;

namespace IndiaEventsWebApi.Repository.Interface
{
    public interface ISample
    {
        Task<VendorCodeGeneration> CreateAsync(VendorCodeGeneration vendorCodeGeneration);
    }
}
