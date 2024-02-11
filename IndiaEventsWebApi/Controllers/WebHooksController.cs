using IndiaEventsWebApi.Models.RequestSheets;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using NLog.Fluent;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Text;

namespace IndiaEventsWebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class WebHooksController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;
        private readonly SmartsheetClient _smartsheetClient;
        private readonly ILogger<WebHooksController> _logger;
        public WebHooksController(IConfiguration configuration, ILogger<WebHooksController> logger, SmartsheetClient smartsheetClient)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
            _smartsheetClient = smartsheetClient;
            _logger = logger;
        }
        string BaseHref => $"{HttpContext.Request.Scheme}://{HttpContext.Request.Host}";


        [HttpPost(Name = "WebHook")]
        public async Task<IActionResult> PostData()
        {
           
                try
                {
                   
                    Dictionary<string, string> requestHeaders = ExtractRequestHeaders();
                    string rawContent = ExtractRawContent();

                    
                    var webhookPayload = JsonConvert.DeserializeObject<Root>(rawContent);

                    
                    var attachedFile = Attachementfile(webhookPayload);

                    
                    var challenge = requestHeaders.GetValueOrDefault("challenge");

                   
                    LogWebhookRequest(requestHeaders);

                   
                    return Ok(new Webhook { smartsheetHookResponse = challenge });
                }
                catch (Exception ex)
                {
                  
                    Log.Error($"Error occurred in Webhook API controller PostData method: {ex.Message} at {DateTime.Now}");

                   
                    return BadRequest(ex.StackTrace);
                }
            
        }

            private Dictionary<string, string> ExtractRequestHeaders()
            {
                Dictionary<string, string> requestHeaders = new Dictionary<string, string>();

                foreach (var header in Request.Headers)
                {
                    requestHeaders.Add(header.Key, header.Value.ToString());
                }

                return requestHeaders;
            }

            private string ExtractRawContent()
            {
                using (var reader = new StreamReader(Request.Body, encoding: Encoding.UTF8, detectEncodingFromByteOrderMarks: false))
                {
                    return reader.ReadToEndAsync().Result;
                }
            }

            private void LogWebhookRequest(Dictionary<string, string> requestHeaders)
            {
               
                Log.Info(string.Join(";", requestHeaders.Select(x => x.Key + "=" + x.Value).ToArray()));
                
            }

            private string Attachementfile(Root webhookPayload)
            {
                

                return "Path to attached file";
            }

        private class Root
        {
        }
        public class Webhook
        {
            public string smartsheetHookResponse { get; set; }

            
        }
    }

}








































//    [HttpPost("WebhookHandler")]
//    public IActionResult WebhookHandler([FromBody] WebhookPayload payload)
//    {
//        try
//        {
//            // Validate payload, check event type, etc.
//            if (payload.EventType == "UPDATE" && payload.ResourceType == "ROW")
//            {
//                // Check if the updated column is "Status" and the new value is "completed"
//                if (payload.Change != null &&
//                    payload.Change.ColumnId == "your_status_column_id" &&
//                    payload.Change.NewValue == "completed")
//                {
//                    // Extract necessary information from the payload, e.g., EventID
//                    string EventID = payload.RowId.ToString(); // Replace with the actual way to get EventID

//                    // Call your existing logic to generate PDF
//                    GenerateSummaryPDF(EventID);
//                }
//            }

//            return Ok();
//        }
//        catch (Exception ex)
//        {
//            // Log and handle exceptions
//            // You might want to return an error response here if needed
//            return BadRequest(ex.Message);
//        }
//    }
//}

//// This class represents the structure of the Smartsheet webhook payload.
//// Please replace it with the actual structure provided by Smartsheet.
//public class WebhookPayload
//{
//    public string EventType { get; set; }
//    public string ResourceType { get; set; }
//    public long? RowId { get; set; }
//    public Change Change { get; set; }
//}

//public class Change
//{
//    public string ColumnId { get; set; }
//    public string NewValue { get; set; }

