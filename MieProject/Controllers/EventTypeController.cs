using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace MieProject.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class EventTypeController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;

        public EventTypeController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

        }
    }
}
