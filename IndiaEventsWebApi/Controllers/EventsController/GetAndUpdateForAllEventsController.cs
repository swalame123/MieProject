using IndiaEventsWebApi.Helper;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;

namespace IndiaEventsWebApi.Controllers.EventsController
{
    [Route("api/[controller]")]
    [ApiController]
    public class GetAndUpdateForAllEventsController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;
        private readonly SmartsheetClient smartsheet;
        private readonly string sheetId1;
        private readonly string sheetId2;
        private readonly string sheetId3;
        private readonly string sheetId4;
        private readonly string sheetId5;
        private readonly string sheetId6;
        private readonly string sheetId7;
        private readonly string sheetId8;
        private readonly string sheetId9;
        public GetAndUpdateForAllEventsController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
            smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            sheetId1 = configuration.GetSection("SmartsheetSettings:Class1").Value;
            sheetId2 = configuration.GetSection("SmartsheetSettings:EventRequestBrandsList").Value;
            sheetId3 = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
            sheetId4 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
            sheetId5 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpSlideKit").Value;
            sheetId6 = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
            sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;
            sheetId8 = configuration.GetSection("SmartsheetSettings:EventRequestBeneficiary").Value;
            sheetId9 = configuration.GetSection("SmartsheetSettings:EventRequestProductBrandsList").Value;
        }
        [HttpGet("GetDataFromAllSheetsUsingEventId")]
        public IActionResult GetDataFromAllSheetsUsingEventId(string eventId)
        {
            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
            //Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
            //Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
            //Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
            //Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
            //Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
            //Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);
            //Sheet sheet8 = SheetHelper.GetSheetById(smartsheet, sheetId8);
            //Sheet sheet9 = SheetHelper.GetSheetById(smartsheet, sheetId9);
            Dictionary<string, object> DraftData = new();
            Row existingRow = sheet1.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == eventId));

            if (existingRow != null)
            {
                List<string> columnNames = new List<string>();
                foreach (Column column in sheet1.Columns)
                {
                    columnNames.Add(column.Title);
                }
                for (int i = 0; i < columnNames.Count; i++)
                {
                    var x = existingRow.Cells[i].Value;
                    if (columnNames[i] == "Brands")
                    {
                        if (x != null || x == "")
                        {
                            List<object> brandsList = SheetHelper.ConvertToJsonObject(existingRow.Cells[i].Value.ToString());
                            DraftData[columnNames[i]] = brandsList;
                        }

                    }
                    else if (columnNames[i] == "Panelists")
                    {
                        if (x != null || x == "")
                        {
                            List<object> Panelists = SheetHelper.ConvertToJsonObject(existingRow.Cells[i].Value.ToString());
                            DraftData[columnNames[i]] = Panelists;
                        }
                    }
                    else if (columnNames[i] == "SlideKits")
                    {
                        if (x != null || x == "")
                        {
                            List<object> SlideKits = SheetHelper.ConvertToJsonObject(existingRow.Cells[i].Value.ToString());
                            DraftData[columnNames[i]] = SlideKits;
                        }
                    }
                    else if (columnNames[i] == "Invitees")
                    {
                        if (x != null || x == "")
                        {
                            List<object> Invitees = SheetHelper.ConvertToJsonObject(existingRow.Cells[i].Value.ToString());
                            DraftData[columnNames[i]] = Invitees;
                        }
                    }


                    else if (columnNames[i] == "Expenses")
                    {
                        if (x != null || x == "")
                        {
                            List<object> Expenses = SheetHelper.ConvertToJsonObject(existingRow.Cells[i].Value.ToString());
                            DraftData[columnNames[i]] = Expenses;

                        }
                    }
                    else
                    {
                        DraftData[columnNames[i]] = existingRow.Cells[i].Value;
                    }


                }
                var attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet1.Id.Value, existingRow.Id.Value, null);

                List<Dictionary<string, object>> attachmentsList = new List<Dictionary<string, object>>();

                foreach (var attachment in attachments.Data)
                {
                    var AID = (long)attachment.Id;
                    var file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet1.Id.Value, AID);
                    Dictionary<string, object> attachmentInfo = new Dictionary<string, object>
                    {
                        { "Name", file.Name },
                        { "Url", file.Url }
                    };
                    attachmentsList.Add(attachmentInfo);
                }
                DraftData["Attachments"] = attachmentsList;
            }
            return Ok(DraftData);





        }

    }




}
