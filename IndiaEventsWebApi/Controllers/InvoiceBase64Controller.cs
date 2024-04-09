
using IndiaEventsWebApi.Helper;
using IndiaEventsWebApi.Junk.Test;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Serilog;
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

        private readonly string ExpenseSheet;
        private readonly string PanelSheet;

        public InvoiceBase64Controller(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
            ExpenseSheet = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
            PanelSheet = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;

            smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

        }



        [HttpPost("GetInvoiceBase64")]



        public IActionResult GetInvoiceBase64(InvoiceIds formdata)
        {
            Dictionary<string, string> idUrlMap = new Dictionary<string, string>(); 

            try
            {
                if (formdata.ExpenseId.Count > 0)
                {
                    Sheet sheet = SheetHelper.GetSheetById(smartsheet, ExpenseSheet);
                    foreach (var id in formdata.ExpenseId)
                    {
                        Row targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == id));

                        if (targetRow != null)
                        {
                            PaginatedResult<Attachment> attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet.Id.Value, targetRow.Id.Value, null);

                            foreach (var attachment in attachments.Data)
                            {
                                if (attachment != null && attachment.Name.Contains("Invoice"))
                                {
                                    long AID = (long)attachment.Id;
                                    string Name = attachment.Name.Split(".")[0];
                                    Attachment file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet.Id.Value, AID);
                                    idUrlMap[id] = file.Url; 
                                }
                            }
                        }
                    }
                }

                var resultArray = idUrlMap.Select(kv => new { Id = kv.Key, Url = kv.Value }).ToArray();

                return Ok(resultArray);
            }
            catch (Exception ex)
            {
                Log.Error($"Error occurred on Webhook apicontroller Attachementfile method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(BadRequest(ex.Message));
            }
        }















        //public IActionResult GetInvoiceBase64(InvoiceIds formdata)
        //{
        //    Dictionary<string, List<string>> dataArray = new();
        //    try
        //    {
        //        if (formdata.ExpenseId.Count > 0)
        //        {
        //            Sheet sheet = SheetHelper.GetSheetById(smartsheet, ExpenseSheet);
        //            foreach (var id in formdata.ExpenseId)
        //            {
        //                Row targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == id));
        //                if (targetRow != null)
        //                {
        //                    PaginatedResult<Attachment> attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet.Id.Value, targetRow.Id.Value, null);
        //                    string url = "";
        //                    foreach (var attachment in attachments.Data)
        //                    {
        //                        if (attachment != null && attachment.Name.Contains("Invoice"))
        //                        {
        //                            long AID = (long)attachment.Id;
        //                            string Name = attachment.Name.Split(".")[0];
        //                            Attachment file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet.Id.Value, AID);
        //                            url = file.Url;

        //                        }

        //                    }
        //                }
        //            }

        //        }
        //        if (formdata.PanelistId.Count > 0)
        //        {
        //            Sheet sheet = SheetHelper.GetSheetById(smartsheet, PanelSheet);
        //            //var targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == EventId));
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Log.Error($"Error occured on Webhook apicontroller Attachementfile method {ex.Message} at {DateTime.Now}");
        //        Log.Error(ex.StackTrace);
        //        return BadRequest(BadRequest(ex.Message));
        //    }
        //    //var targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == EventId));

        //    //if (targetRow != null)
        //    //{
        //    //    var attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet.Id.Value, targetRow.Id.Value, null);
        //    //    var url = "";
        //    //    foreach (var attachment in attachments.Data)
        //    //    {
        //    //        if (attachment != null)
        //    //        {
        //    //            var AID = (long)attachment.Id;
        //    //            var Name = attachment.Name.Split(".")[0];
        //    //            var file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet.Id.Value, AID);
        //    //            url = file.Url;
        //    //            using (HttpClient client = new HttpClient())
        //    //            {
        //    //                var fileContent = client.GetByteArrayAsync(url).Result;
        //    //                var base64String = Convert.ToBase64String(fileContent);
        //    //                string concatname = attachment.Name.Split('-')[1];
        //    //                var Data = $"{concatname}:{base64String}";
        //    //                if (Name.Split("-")[1] == "BTEInvoice")
        //    //                {
        //    //                    if (!dataArray.ContainsKey("BTEInvoice"))
        //    //                    {
        //    //                        dataArray["BTEInvoice"] = new List<string>();
        //    //                    }
        //    //                    dataArray["BTEInvoice"].Add(Data);
        //    //                }
        //    //                else if (Name.Split("-")[1] == "BTCInvoice")
        //    //                {
        //    //                    if (!dataArray.ContainsKey("BTCInvoice"))
        //    //                    {
        //    //                        dataArray["BTCInvoice"] = new List<string>();
        //    //                    }
        //    //                    dataArray["BTCInvoice"].Add(Data);
        //    //                }
        //    //            }
        //    //        }
        //    //    }
        //    //}

        //    return Ok(dataArray);
        //}


    }

}
