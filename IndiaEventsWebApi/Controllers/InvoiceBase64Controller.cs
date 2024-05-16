
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
                    var resultArray = idUrlMap.Select(kv => new { Id = kv.Key, Url = kv.Value }).ToArray();
                    return Ok(resultArray);
                }

                else if (formdata.PanelistId.Count > 0)
                {
                    Sheet sheet = SheetHelper.GetSheetById(smartsheet, PanelSheet);
                    foreach (var id in formdata.PanelistId)
                    {
                        Row targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == id));

                        if (targetRow != null)
                        {
                            PaginatedResult<Attachment> attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet.Id.Value, targetRow.Id.Value, null);

                            foreach (var attachment in attachments.Data)
                            {

                                if (attachment != null && attachment.Name.ToLower().Contains("agreement"))
                                {
                                    long AID = (long)attachment.Id;
                                    string Name = attachment.Name;
                                    Attachment file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet.Id.Value, AID);

                                    idUrlMap[Name + "*" + id] = file.Url;
                                }
                            }
                        }
                    }
                    var resultArray = idUrlMap.Select(kv => new { Id = kv.Key.Split("*")[1], Name = kv.Key.Split("*")[0], Url = kv.Value }).ToArray();
                    return Ok(resultArray);
                }
                return Ok();



            }
            catch (Exception ex)
            {
                Log.Error($"Error occurred on Webhook apicontroller Attachementfile method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(BadRequest(ex.Message));
            }
        }

        //[HttpPost("GetInvoiceBase64WithAttachmentsIds")]
        //public IActionResult GetInvoiceBase64WithAttachmentsIds(InvoiceIds formdata)
        //{
        //    Dictionary<string, string> idUrlMap = new Dictionary<string, string>();

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

        //                    foreach (var attachment in attachments.Data)
        //                    {
        //                        if (attachment != null && attachment.Name.Contains("Invoice"))
        //                        {
        //                            long AID = (long)attachment.Id;
        //                            string Name = attachment.Name.Split(".")[0];
        //                            Attachment file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet.Id.Value, AID);
        //                            idUrlMap[id] = file.Url;
        //                        }
        //                    }
        //                }
        //            }
        //            var resultArray = idUrlMap.Select(kv => new { Id = kv.Key, Url = kv.Value }).ToArray();
        //            return Ok(resultArray);
        //        }

        //        else if (formdata.PanelistId.Count > 0)
        //        {
        //            Sheet sheet = SheetHelper.GetSheetById(smartsheet, PanelSheet);
        //            foreach (var id in formdata.PanelistId)
        //            {
        //                Row targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == id));

        //                if (targetRow != null)
        //                {
        //                    PaginatedResult<Attachment> attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet.Id.Value, targetRow.Id.Value, null);

        //                    foreach (var attachment in attachments.Data)
        //                    {

        //                        if (attachment != null && attachment.Name.ToLower().Contains("agreement"))
        //                        {
        //                            long AID = (long)attachment.Id;
        //                            string Name = attachment.Name;
        //                            Attachment file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet.Id.Value, AID);

        //                            idUrlMap[Name + "*" + id] = file.Url;
        //                        }
        //                    }
        //                }
        //            }
        //            var resultArray = idUrlMap.Select(kv => new { Id = kv.Key.Split("*")[1], Name = kv.Key.Split("*")[0], Url = kv.Value }).ToArray();
        //            return Ok(resultArray);
        //        }
        //        return Ok();



        //    }
        //    catch (Exception ex)
        //    {
        //        Log.Error($"Error occurred on Webhook apicontroller Attachementfile method {ex.Message} at {DateTime.Now}");
        //        Log.Error(ex.StackTrace);
        //        return BadRequest(BadRequest(ex.Message));
        //    }
        //}




    }

}
