namespace IndiaEventsWebApi.Models.RequestSheets
{
    public class EventRequestBrandsList
    {
        public string? BrandName { get; set; }
        public string? PercentAllocation { get; set; }
        public string? ProjectId { get; set; }
    }
    public class UpdateEventRequestBrandsList
    {
        public string BrandId { get; set; }
        public string? BrandName { get; set; }
        public string? PercentAllocation { get; set; }
        public string? ProjectId { get; set; }
    }

    
}
