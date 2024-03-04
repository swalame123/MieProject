using IndiaEventsWebApi.Models.EventTypeSheets;
using Microsoft.AspNetCore.Http;

namespace IndiaEventsWebApi.Models
{
    public class FileUploadModel
    {
        public IFormFile File { get; set; }
        public Class1 FormData {  get; set; }
    }
}
