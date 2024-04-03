using Aspose.Cells;
using IndiaEventsWebApi.Helper;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;

namespace IndiaEventsWebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class InvoiceBase64Controller : ControllerBase
    {

        private readonly string accessToken;
        private readonly IConfiguration configuration;
        private readonly SmartsheetClient smartsheet;

        private readonly Sheet sheet;

        public InvoiceBase64Controller(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
            string sheetId = configuration.GetSection("SmartsheetSettings:EventSettlement").Value;

            smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
        }



        [HttpGet("GetInvoiceBase64")]
        public IActionResult GetInvoiceBase64(string EventId)
        {
            var dataArray = new Dictionary<string, List<string>>();
            var targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == EventId));

            if (targetRow != null)
            {
                var attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet.Id.Value, targetRow.Id.Value, null);
                var url = "";
                foreach (var attachment in attachments.Data)
                {
                    if (attachment != null)
                    {
                        var AID = (long)attachment.Id;
                        var Name = attachment.Name.Split(".")[0];
                        var file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet.Id.Value, AID);
                        url = file.Url;
                        using (HttpClient client = new HttpClient())
                        {
                            var fileContent = client.GetByteArrayAsync(url).Result;
                            var base64String = Convert.ToBase64String(fileContent);
                            string concatname = attachment.Name.Split('-')[1];
                            var Data = $"{concatname}:{base64String}";
                            if (Name.Split("-")[1] == "BTEInvoice")
                            {
                                if (!dataArray.ContainsKey("BTEInvoice"))
                                {
                                    dataArray["BTEInvoice"] = new List<string>();
                                }
                                dataArray["BTEInvoice"].Add(Data);
                            }
                            else if (Name.Split("-")[1] == "BTCInvoice")
                            {
                                if (!dataArray.ContainsKey("BTCInvoice"))
                                {
                                    dataArray["BTCInvoice"] = new List<string>();
                                }
                                dataArray["BTCInvoice"].Add(Data);
                            }
                        }
                    }
                }
            }

            return Ok(dataArray);
        }


    }
}
