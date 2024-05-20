using IndiaEvents.Models.Models.Draft;
using IndiaEvents.Models.Models.EventTypeSheets;
using IndiaEventsWebApi.Helper;
using IndiaEventsWebApi.Models;
using IndiaEventsWebApi.Models.EventTypeSheets;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI.HPSF;
using Serilog;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Security.Policy;
using System.Text;
using static Org.BouncyCastle.Bcpg.Attr.ImageAttrib;

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
        private readonly string sheetId10;
        private readonly string sheetId11;
        public GetAndUpdateForAllEventsController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
            smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            sheetId1 = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;
            sheetId2 = configuration.GetSection("SmartsheetSettings:EventRequestBrandsList").Value;
            sheetId3 = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
            sheetId4 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
            sheetId5 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpSlideKit").Value;
            sheetId6 = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
            sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;
            sheetId8 = configuration.GetSection("SmartsheetSettings:EventRequestBeneficiary").Value;
            sheetId9 = configuration.GetSection("SmartsheetSettings:EventRequestProductBrandsList").Value;
            sheetId10 = configuration.GetSection("SmartsheetSettings:EventSettlement").Value;
            sheetId11 = configuration.GetSection("SmartsheetSettings:HonorariumPayment").Value;
        }
        #region
        //[HttpGet("GetDataFromAllSheetsUsingEventId")]
        //public IActionResult GetDataFromAllSheetsUsingEventId(string eventId)
        //{
        //    Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
        //    //Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
        //    //Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
        //    //Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
        //    //Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
        //    //Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
        //    //Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);
        //    //Sheet sheet8 = SheetHelper.GetSheetById(smartsheet, sheetId8);
        //    //Sheet sheet9 = SheetHelper.GetSheetById(smartsheet, sheetId9);
        //    Dictionary<string, object> DraftData = new();
        //    Row existingRow = sheet1.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == eventId));

        //    if (existingRow != null)
        //    {
        //        List<string> columnNames = new List<string>();
        //        foreach (Column column in sheet1.Columns)
        //        {
        //            columnNames.Add(column.Title);
        //        }
        //        for (int i = 0; i < columnNames.Count; i++)
        //        {
        //            var x = existingRow.Cells[i].Value;
        //            if (columnNames[i] == "Brands")
        //            {
        //                if (x != null || x == "")
        //                {
        //                    List<object> brandsList = SheetHelper.ConvertToJsonObject(existingRow.Cells[i].Value.ToString());
        //                    DraftData[columnNames[i]] = brandsList;
        //                }

        //            }
        //            else if (columnNames[i] == "Panelists")
        //            {
        //                if (x != null || x == "")
        //                {
        //                    List<object> Panelists = SheetHelper.ConvertToJsonObject(existingRow.Cells[i].Value.ToString());
        //                    DraftData[columnNames[i]] = Panelists;
        //                }
        //            }
        //            else if (columnNames[i] == "SlideKits")
        //            {
        //                if (x != null || x == "")
        //                {
        //                    List<object> SlideKits = SheetHelper.ConvertToJsonObject(existingRow.Cells[i].Value.ToString());
        //                    DraftData[columnNames[i]] = SlideKits;
        //                }
        //            }
        //            else if (columnNames[i] == "Invitees")
        //            {
        //                if (x != null || x == "")
        //                {
        //                    List<object> Invitees = SheetHelper.ConvertToJsonObject(existingRow.Cells[i].Value.ToString());
        //                    DraftData[columnNames[i]] = Invitees;
        //                }
        //            }


        //            else if (columnNames[i] == "Expenses")
        //            {
        //                if (x != null || x == "")
        //                {
        //                    List<object> Expenses = SheetHelper.ConvertToJsonObject(existingRow.Cells[i].Value.ToString());
        //                    DraftData[columnNames[i]] = Expenses;

        //                }
        //            }
        //            else
        //            {
        //                DraftData[columnNames[i]] = existingRow.Cells[i].Value;
        //            }


        //        }
        //        var attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet1.Id.Value, existingRow.Id.Value, null);

        //        List<Dictionary<string, object>> attachmentsList = new List<Dictionary<string, object>>();

        //        foreach (var attachment in attachments.Data)
        //        {
        //            var AID = (long)attachment.Id;
        //            var file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet1.Id.Value, AID);
        //            Dictionary<string, object> attachmentInfo = new Dictionary<string, object>
        //            {
        //                { "Name", file.Name },
        //                { "Url", file.Url }
        //            };
        //            attachmentsList.Add(attachmentInfo);
        //        }
        //        DraftData["Attachments"] = attachmentsList;
        //    }
        //    return Ok(DraftData);





        //}


        //[HttpGet("SampleGetDataFromAllSheetsUsingEventIdInPreEvent")]
        //public IActionResult SampleGetDataFromAllSheetsUsingEventId(string eventId)
        //{
        //    List<UpdateDataForClassI> formData = new List<UpdateDataForClassI>();

        //    Dictionary<string, object> resultData = new Dictionary<string, object>();
        //    List<Sheet> sheets = new List<Sheet>
        //    {
        //        SheetHelper.GetSheetById(smartsheet, sheetId1),
        //        SheetHelper.GetSheetById(smartsheet, sheetId2),
        //        SheetHelper.GetSheetById(smartsheet, sheetId3),
        //        SheetHelper.GetSheetById(smartsheet, sheetId4),
        //        SheetHelper.GetSheetById(smartsheet, sheetId5),
        //        SheetHelper.GetSheetById(smartsheet, sheetId6),
        //        SheetHelper.GetSheetById(smartsheet, sheetId7),
        //        SheetHelper.GetSheetById(smartsheet, sheetId8),
        //        SheetHelper.GetSheetById(smartsheet, sheetId9)
        //    };
        //    foreach (var sheet in sheets)
        //    {
        //        List<Dictionary<string, object>> rowsData = new List<Dictionary<string, object>>();

        //        foreach (var row in sheet.Rows)
        //        {
        //            if (row.Cells.Any(c => c.DisplayValue == eventId))
        //            {
        //                Dictionary<string, object> rowData = new Dictionary<string, object>();

        //                List<string> columnNames = new List<string>();
        //                foreach (Column column in sheet.Columns)
        //                {
        //                    columnNames.Add(column.Title);
        //                }

        //                for (int i = 0; i < columnNames.Count; i++)
        //                {
        //                    rowData[columnNames[i]] = row.Cells[i].Value;
        //                }

        //                var attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet.Id.Value, row.Id.Value, null);
        //                List<Dictionary<string, object>> attachmentsList = new List<Dictionary<string, object>>();
        //                foreach (var attachment in attachments.Data)
        //                {
        //                    var AID = (long)attachment.Id;
        //                    var file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet.Id.Value, AID);
        //                    Dictionary<string, object> attachmentInfo = new Dictionary<string, object>
        //                    {
        //                        { "Name", file.Name },
        //                        { "Url", file.Url }
        //                    };
        //                    attachmentsList.Add(attachmentInfo);
        //                }

        //                rowData["Attachments"] = attachmentsList;
        //                rowsData.Add(rowData);
        //            }
        //        }

        //        if (rowsData.Count > 0)
        //        {
        //            resultData[sheet.Name] = rowsData;
        //        }
        //    }
        //    //foreach (var formdata in resultData)
        //    //{
        //    //    if (formdata.ContainsKey("Event Topic"))
        //    //    {
        //    //        specificData["Event Topic"] = resultData["Event Topic"];
        //    //    }
        //    //}




        //    string jsonData = JsonConvert.SerializeObject(resultData);

        //    //return Ok(jsonData);
        //    //return Ok(specificData);





        //    return Ok(resultData);
        //    //return Ok(extractedData);




        //}

        #endregion

        [HttpGet("GetDataFromAllSheetsUsingEventIdInPreEvent")]
        public IActionResult GetDataFromAllSheetsUsingEventId(string eventId)
        {
            Dictionary<string, object> resultData = new();

            List<UpdateDataForClassI> formData = new();

            List<Dictionary<string, object>> eventDetails = new();
            List<Dictionary<string, object>> BrandseventDetails = new();
            List<Dictionary<string, object>> InviteeseventDetails = new();
            List<Dictionary<string, object>> PaneleventDetails = new();
            List<Dictionary<string, object>> SlideKiteventDetails = new();
            List<Dictionary<string, object>> ExpenseeventDetails = new();
            List<Dictionary<string, object>> attachmentsList = new List<Dictionary<string, object>>();
            List<Dictionary<string, object>> DeviationsattachmentsList = new List<Dictionary<string, object>>();
            List<Dictionary<string, object>> attachmentInfoFiles = new();

            Sheet sheet1 = (Sheet)SheetHelper.GetSheetById(smartsheet, sheetId1);
            List<string> columnNames = new List<string>();
            foreach (Column column in sheet1.Columns)
            {
                columnNames.Add(column.Title);
            }

            foreach (var row in sheet1.Rows)
            {

                if (row.Cells.Any(c => c.DisplayValue == eventId))
                {
                    List<string> columnsToInclude = new List<string> { "Event Date", "Event Topic", "Class III Event Code", "Start Time",
                                                                        "End Time", "State", "Venue Name", 
                                                                        "City", "Brands", "Panelists", "SlideKits", "Invitees", 
                                                                        "MIPL Invitees", "Expenses", "Meeting Type" };

                    Dictionary<string, object> rowData = new Dictionary<string, object>();
                    for (int i = 0; i < columnNames.Count; i++)
                    {
                        if (columnsToInclude.Contains(columnNames[i]))
                        {
                            rowData[columnNames[i]] = row.Cells[i].Value;
                        }
                    }
                    PaginatedResult<Attachment> attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet1.Id.Value, row.Id.Value, null);


                    if (attachments.Data != null || attachments.Data.Count > 0)
                    {
                        foreach (var attachment in attachments.Data)
                        {
                            long AID = (long)attachment.Id;
                            Attachment file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet1.Id.Value, AID);

                            Dictionary<string, object> attachmentInfoData = new()
                            {
                                { "Name", file.Name },
                                { "Id", file.Id },
                                { "base64", SheetHelper.UrlToBaseValue(file.Url) }
                            };
                            attachmentInfoFiles.Add(attachmentInfoData);
                        }

                    }


                    eventDetails.Add(rowData);
                }
            }

            Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
            List<string> BrandsColumnNames = new List<string>();
            foreach (Column column in sheet2.Columns)
            {
                BrandsColumnNames.Add(column.Title);
            }

            foreach (var row in sheet2.Rows)
            {

                if (row.Cells.Any(c => c.DisplayValue == eventId))
                {
                    List<string> BrandscolumnsToInclude = new List<string> { "BrandRequestID", "Brands", "% Allocation", "Project ID" };

                    Dictionary<string, object> BrandsrowData = new Dictionary<string, object>();
                    for (int i = 0; i < BrandsColumnNames.Count; i++)
                    {
                        if (BrandscolumnsToInclude.Contains(BrandsColumnNames[i]))
                        {
                            BrandsrowData[BrandsColumnNames[i]] = row.Cells[i].Value;
                        }
                    }
                    PaginatedResult<Attachment> attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet2.Id.Value, row.Id.Value, null);

                    List<Dictionary<string, object>> BrandsattachmentsList = new();
                    if (attachments.Data != null || attachments.Data.Count > 0)
                    {
                        foreach (var attachment in attachments.Data)
                        {
                            long AID = (long)attachment.Id;
                            Attachment file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet2.Id.Value, AID);
                            Dictionary<string, object> attachmentInfo = new()
                            {
                                { "Name", file.Name },
                                { "Id", file.Id },

                                { "base64", SheetHelper.UrlToBaseValue(file.Url) }
                            };
                            BrandsattachmentsList.Add(attachmentInfo);
                        }
                        BrandsrowData["Attachments"] = BrandsattachmentsList;
                    }
                    BrandseventDetails.Add(BrandsrowData);
                }
            }
            Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
            List<string> InviteesColumnNames = new List<string>();
            foreach (Column column in sheet3.Columns)
            {
                InviteesColumnNames.Add(column.Title);
            }
            foreach (var row in sheet3.Rows)
            {
                if (row.Cells.Any(c => c.DisplayValue == eventId))
                {
                    List<string> InviteescolumnsToInclude = new List<string> { "INV", "Invitee Source", "HCPName", "MISCode", "Employee Code", "LocalConveyance", "BTC/BTE", "LocalConveyance", "Speciality", "Lc Amount Excluding Tax", "LcAmount" };
                    Dictionary<string, object> InviteesrowData = new Dictionary<string, object>();
                    for (int i = 0; i < InviteesColumnNames.Count; i++)
                    {
                        if (InviteescolumnsToInclude.Contains(InviteesColumnNames[i]))
                        {
                            InviteesrowData[InviteesColumnNames[i]] = row.Cells[i].Value;
                        }
                    }
                    PaginatedResult<Attachment> attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet3.Id.Value, row.Id.Value, null);

                    List<Dictionary<string, object>> InviteesattachmentsList = new();
                    if (attachments.Data != null || attachments.Data.Count > 0)
                    {
                        foreach (var attachment in attachments.Data)
                        {
                            long AID = (long)attachment.Id;
                            Attachment file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet3.Id.Value, AID);
                            Dictionary<string, object> attachmentInfo = new()
                            {
                                { "Name", file.Name },
                                { "Id", file.Id },
                                { "base64" , SheetHelper.UrlToBaseValue(file.Url) }
                            };
                            InviteesattachmentsList.Add(attachmentInfo);
                        }
                        InviteesrowData["Attachments"] = InviteesattachmentsList;
                    }
                    InviteeseventDetails.Add(InviteesrowData);
                }
            }
            Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
            List<string> PanelColumnNames = new List<string>();
            foreach (Column column in sheet4.Columns)
            {
                PanelColumnNames.Add(column.Title);
            }
            foreach (var row in sheet4.Rows)
            {
                if (row.Cells.Any(c => c.DisplayValue == eventId))
                {
                    List<string> PanelcolumnsToInclude = new List<string> { "Panelist ID", "SpeakerCode", "TrainerCode", "Tier", "Qualification", "Speciality", "Country",
                        "Rationale", "Speciality", "FCPA Date", "LcAmount", "PresentationDuration", "PanelSessionPreparationDuration", "PanelDiscussionDuration", "QASessionDuration",
                        "BriefingSession", "TotalSessionHours", "HcpRole", "HCPName", "MISCode", "HCP Type", "ExpenseType", "HonorariumRequired", "HonorariumAmount", "Honorarium Amount Excluding Tax",
                        "Travel BTC/BTE", "Mode of Travel", "Travel Excluding Tax", "Travel", "Accomodation Excluding Tax", "Accomodation","Accomodation BTC/BTE",
                        "Local Conveyance Excluding Tax", "LocalConveyance", "LC BTC/BTE",
                        "PAN card name", "Bank Account Number", "IFSC Code", "Bank Name", "Currency", "Other Currency", "Beneficiary Name", "Pan Number", "Global FMV", "Swift Code" };

                    Dictionary<string, object> PanelrowData = new Dictionary<string, object>();
                    for (int i = 0; i < PanelColumnNames.Count; i++)
                    {
                        if (PanelcolumnsToInclude.Contains(PanelColumnNames[i]))
                        {
                            PanelrowData[PanelColumnNames[i]] = row.Cells[i].Value;
                        }
                    }
                    PaginatedResult<Attachment> attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet4.Id.Value, row.Id.Value, null);

                    List<Dictionary<string, object>> PanelattachmentsList = new();
                    if (attachments.Data != null || attachments.Data.Count > 0)
                    {
                        foreach (var attachment in attachments.Data)
                        {
                            long AID = (long)attachment.Id;
                            Attachment file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet4.Id.Value, AID);
                            Dictionary<string, object> attachmentInfo = new()
                            {
                                { "Name", file.Name },
                                { "Id", file.Id },
                                { "base64" , SheetHelper.UrlToBaseValue(file.Url) }
                            };
                            PanelattachmentsList.Add(attachmentInfo);
                        }
                        PanelrowData["Attachments"] = PanelattachmentsList;
                    }
                    PaneleventDetails.Add(PanelrowData);
                }
            }
            Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
            List<string> SlideKitColumnNames = new List<string>();
            foreach (Column column in sheet5.Columns)
            {
                SlideKitColumnNames.Add(column.Title);
            }
            foreach (var row in sheet5.Rows)
            {
                if (row.Cells.Any(c => c.DisplayValue == eventId))
                {
                    List<string> SlideKitcolumnsToInclude = new List<string> { "SlideKit ID", "HCP Name", "MIS", "HcpType", "Slide Kit Type", "SlideKit Document" };

                    Dictionary<string, object> SlideKitrowData = new Dictionary<string, object>();
                    for (int i = 0; i < SlideKitColumnNames.Count; i++)
                    {
                        if (SlideKitcolumnsToInclude.Contains(SlideKitColumnNames[i]))
                        {
                            SlideKitrowData[SlideKitColumnNames[i]] = row.Cells[i].Value;
                        }
                    }
                    PaginatedResult<Attachment> attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet5.Id.Value, row.Id.Value, null);

                    List<Dictionary<string, object>> SlideKitattachmentsList = new();
                    if (attachments.Data != null || attachments.Data.Count > 0)
                    {
                        foreach (var attachment in attachments.Data)
                        {
                            long AID = (long)attachment.Id;
                            Attachment file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet5.Id.Value, AID);
                            Dictionary<string, object> attachmentInfo = new()
                            {
                                { "Name", file.Name },
                                { "Id", file.Id },
                                { "base64" , SheetHelper.UrlToBaseValue(file.Url) }
                            };
                            SlideKitattachmentsList.Add(attachmentInfo);
                        }
                        SlideKitrowData["Attachments"] = SlideKitattachmentsList;
                    }
                    SlideKiteventDetails.Add(SlideKitrowData);
                }
            }
            Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
            List<string> ExpenseColumnNames = new List<string>();
            foreach (Column column in sheet6.Columns)
            {
                ExpenseColumnNames.Add(column.Title);
            }
            foreach (var row in sheet6.Rows)
            {
                if (row.Cells.Any(c => c.DisplayValue == eventId))
                {
                    List<string> ExpensecolumnsToInclude = new List<string> { "Expenses ID", "Expense", "BTC/BTE", "Amount Excluding Tax", "Amount" };

                    Dictionary<string, object> ExpenserowData = new Dictionary<string, object>();
                    for (int i = 0; i < ExpenseColumnNames.Count; i++)
                    {
                        if (ExpensecolumnsToInclude.Contains(ExpenseColumnNames[i]))
                        {
                            ExpenserowData[ExpenseColumnNames[i]] = row.Cells[i].Value;
                        }
                    }
                    ExpenseeventDetails.Add(ExpenserowData);
                }
            }
            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);
            List<string> DeviationscolumnNames = new List<string>();
            foreach (Column column in sheet7.Columns)
            {
                DeviationscolumnNames.Add(column.Title);
            }
            foreach (var row in sheet7.Rows)
            {
                if (row.Cells.Any(c => c.DisplayValue == eventId))
                {
                    List<string> DeviationscolumnsToInclude = new List<string> { "Deviation Type" };

                    Dictionary<string, object> DeviationsattachmentInfo = new Dictionary<string, object>();
                    for (int i = 0; i < DeviationscolumnNames.Count; i++)
                    {
                        if (DeviationscolumnsToInclude.Contains(DeviationscolumnNames[i]))
                        {
                            var val = row.Cells[i].Value.ToString();
                            var attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet7.Id.Value, row.Id.Value, null);
                            foreach (var attachment in attachments.Data)
                            {
                                var AID = (long)attachment.Id;
                                var file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet7.Id.Value, AID);
                                DeviationsattachmentInfo[val] = SheetHelper.UrlToBaseValue(file.Url);
                            }
                        }
                    }
                    DeviationsattachmentsList.Add(DeviationsattachmentInfo);
                }
            }
            resultData["eventDetails"] = eventDetails;
            resultData["Files"] = attachmentInfoFiles;
            resultData["Brands"] = BrandseventDetails;
            resultData["Invitees"] = InviteeseventDetails;
            resultData["PanelDetails"] = PaneleventDetails;
            resultData["SlideKitSelection"] = SlideKiteventDetails;
            resultData["ExpenseSelection"] = ExpenseeventDetails;
            resultData["Deviation"] = DeviationsattachmentsList;
            return Ok(resultData);
        }

        [HttpGet("GetDataFromAllSheetsByEventIdInPreEvent")]
        public IActionResult GetDataFromAllSheetsByEventIdInPreEvent(string eventId)
        {
            Dictionary<string, object> resultData = new();

            List<UpdateDataForClassI> formData = new();

            List<Dictionary<string, object>> eventDetails = new();
            List<Dictionary<string, object>> BrandseventDetails = new();
            List<Dictionary<string, object>> InviteeseventDetails = new();
            List<Dictionary<string, object>> PaneleventDetails = new();
            List<Dictionary<string, object>> SlideKiteventDetails = new();
            List<Dictionary<string, object>> ExpenseeventDetails = new();
            List<Dictionary<string, object>> BeneficiarteventDetails = new();
            List<Dictionary<string, object>> ProductBrandsListeventDetails = new();
            List<Dictionary<string, object>> attachmentsList = new List<Dictionary<string, object>>();
            List<Dictionary<string, object>> DeviationsattachmentsList = new List<Dictionary<string, object>>();
            List<Dictionary<string, object>> attachmentInfoFiles = new();

            Sheet sheet1 = (Sheet)SheetHelper.GetSheetById(smartsheet, sheetId1);
            List<string> columnNames = new List<string>();
            foreach (Column column in sheet1.Columns)
            {
                columnNames.Add(column.Title);
            }
            foreach (var row in sheet1.Rows.Where(row => row.Cells.Any(c => c.DisplayValue == eventId)))
            {
                Dictionary<string, object> rowData = new Dictionary<string, object>
                {
                    { "EventDate", GetValueOrDefault(row, "Event Date") },
                    { "EventEndDate", GetValueOrDefault(row, "End Date") },
                    { "EventTopic", GetValueOrDefault(row, "Event Topic") },
                    { "ClassIIIEventCode", GetValueOrDefault(row, "Class III Event Code") },
                    { "StartTime", GetValueOrDefault(row, "Start Time") },
                    { "EndTime", GetValueOrDefault(row, "End Time") },
                    { "State", GetValueOrDefault(row, "State") },
                    { "VenueName", GetValueOrDefault(row, "Venue Name") },
                    { "City", GetValueOrDefault(row, "City") },
                    { "Brands", GetValueOrDefault(row, "Brands") },
                    { "Panelists", GetValueOrDefault(row, "Panelists") },
                    { "SlideKits", GetValueOrDefault(row, "SlideKits") },
                    { "Invitees", GetValueOrDefault(row, "Invitees") },
                    { "MIPLInvitees", GetValueOrDefault(row, "MIPL Invitees") },
                    { "Expenses", GetValueOrDefault(row, "Expenses") },
                    { "MeetingType", GetValueOrDefault(row, "Meeting Type") },
                    { "ModeOfTraining", GetValueOrDefault(row, "Mode of Training") },
                    { "EmergencySupport", GetValueOrDefault(row, "Emergency Support") },
                    { "IsFacilityCharges", GetValueOrDefault(row, "Facility Charges") },
                    { "FacilityChargesBTC/BTE", GetValueOrDefault(row, "Facility Charges BTC/BTE") },
                    { "FacilityChargesIncludingTax", GetValueOrDefault(row, "Total Facility Charges Including Tax") },
                    { "FacilityChargesExcludingTax", GetValueOrDefault(row, "Facility Charges Excluding Tax") },
                    { "HOTWebinarType", GetValueOrDefault(row, "HOT Webinar Type") },
                    { "HOTWebinarVendorName", GetValueOrDefault(row, "HOT Webinar Vendor Name") },
                    { "VenueSelectionChecklist", GetValueOrDefault(row, "Venue Selection Checklist") },
                    { "EmergencyContactNo", GetValueOrDefault(row, "Emergency Contact No") },
                    { "IsAnesthetistRequired", GetValueOrDefault(row, "Anesthetist Required?") },
                    { "AnesthetistBTC/BTE", GetValueOrDefault(row, "Anesthetist BTC/BTE") },
                    { "AnesthetistExcludingTax", GetValueOrDefault(row, "Anesthetist Excluding Tax") },
                    { "AnesthetistincludingTax", GetValueOrDefault(row, "Anesthetist including Tax") }
                };
                PaginatedResult<Attachment> attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet1.Id.Value, row.Id.Value, null);

                if (attachments.Data != null && attachments.Data.Count > 0)
                {
                    foreach (var attachment in attachments.Data)
                    {
                        long AID = (long)attachment.Id;
                        Attachment file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet1.Id.Value, AID);
                        Dictionary<string, object> attachmentInfoData = new Dictionary<string, object>
                        {
                            { "Name", file.Name },
                            { "Id", file.Id },
                            { "base64", SheetHelper.UrlToBaseValue(file.Url) }
                        };
                        attachmentInfoFiles.Add(attachmentInfoData);
                    }
                }

                eventDetails.Add(rowData);
            }


            string? GetValueOrDefault(Row row, string columnName)
            {
                int columnIndex = columnNames.IndexOf(columnName);
                if (columnIndex >= 0 && columnIndex < row.Cells.Count)
                {
                    return row.Cells[columnIndex].Value?.ToString();
                }
                return null;
            }



            Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
            List<string> BrandsColumnNames = new List<string>();
            foreach (Column column in sheet2.Columns)
            {
                BrandsColumnNames.Add(column.Title);
            }

            foreach (var row in sheet2.Rows)
            {

                if (row.Cells.Any(c => c.DisplayValue == eventId))
                {
                    List<string> BrandscolumnsToInclude = new List<string> { "BrandRequestID", "Brands", "% Allocation", "Project ID" };

                    Dictionary<string, object> BrandsrowData = new Dictionary<string, object>();
                    for (int i = 0; i < BrandsColumnNames.Count; i++)
                    {
                        if (BrandscolumnsToInclude.Contains(BrandsColumnNames[i]))
                        {
                            BrandsrowData[BrandsColumnNames[i]] = row.Cells[i].Value;
                        }
                    }
                    PaginatedResult<Attachment> attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet2.Id.Value, row.Id.Value, null);

                    List<Dictionary<string, object>> BrandsattachmentsList = new();
                    if (attachments.Data != null || attachments.Data.Count > 0)
                    {
                        foreach (var attachment in attachments.Data)
                        {
                            long AID = (long)attachment.Id;
                            Attachment file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet2.Id.Value, AID);
                            Dictionary<string, object> attachmentInfo = new()
                            {
                                { "Name", file.Name },
                                { "Id", file.Id },
                                { "base64", SheetHelper.UrlToBaseValue(file.Url) }
                            };
                            BrandsattachmentsList.Add(attachmentInfo);
                        }
                        BrandsrowData["Attachments"] = BrandsattachmentsList;
                    }
                    BrandseventDetails.Add(BrandsrowData);
                }
            }
            Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
            List<string> InviteesColumnNames = new List<string>();
            foreach (Column column in sheet3.Columns)
            {
                InviteesColumnNames.Add(column.Title);
            }
            foreach (var row in sheet3.Rows)
            {
                if (row.Cells.Any(c => c.DisplayValue == eventId))
                {
                    List<string> InviteescolumnsToInclude = new List<string> { "INV", "Invitee Source", "HCPName", "MISCode", "Employee Code", "LocalConveyance", "BTC/BTE", "LocalConveyance", "Speciality", "Lc Amount Excluding Tax", "LcAmount" };
                    Dictionary<string, object> InviteesrowData = new Dictionary<string, object>();
                    for (int i = 0; i < InviteesColumnNames.Count; i++)
                    {
                        if (InviteescolumnsToInclude.Contains(InviteesColumnNames[i]))
                        {
                            InviteesrowData[InviteesColumnNames[i]] = row.Cells[i].Value;
                        }
                    }
                    PaginatedResult<Attachment> attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet3.Id.Value, row.Id.Value, null);

                    List<Dictionary<string, object>> InviteesattachmentsList = new();
                    if (attachments.Data != null || attachments.Data.Count > 0)
                    {
                        foreach (var attachment in attachments.Data)
                        {
                            long AID = (long)attachment.Id;
                            Attachment file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet3.Id.Value, AID);
                            Dictionary<string, object> attachmentInfo = new()
                            {
                                { "Name", file.Name },
                                { "Id", file.Id },
                                { "base64" , SheetHelper.UrlToBaseValue(file.Url) }
                            };
                            InviteesattachmentsList.Add(attachmentInfo);
                        }
                        InviteesrowData["Attachments"] = InviteesattachmentsList;
                    }
                    InviteeseventDetails.Add(InviteesrowData);
                }
            }
            Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
            List<string> PanelColumnNames = new List<string>();
            foreach (Column column in sheet4.Columns)
            {
                PanelColumnNames.Add(column.Title);
            }
            foreach (var row in sheet4.Rows)
            {
                if (row.Cells.Any(c => c.DisplayValue == eventId))
                {
                    List<string> PanelcolumnsToInclude = new List<string> { "Panelist ID", "SpeakerCode", "TrainerCode", "Tier", "Qualification", "Speciality", "Country",
                        "Rationale", "Speciality", "FCPA Date", "LcAmount", "PresentationDuration", "PanelSessionPreparationDuration", "PanelDiscussionDuration", "QASessionDuration",
                        "BriefingSession", "TotalSessionHours", "HcpRole", "HCPName", "MISCode", "HCP Type", "ExpenseType", "HonorariumRequired", "HonorariumAmount", "Honorarium Amount Excluding Tax",
                        "Travel BTC/BTE", "Mode of Travel", "Travel Excluding Tax", "Travel", "Accomodation Excluding Tax", "Accomodation","Accomodation BTC/BTE","Annual Trainer Agreement Valid?",
                        "Local Conveyance Excluding Tax", "LocalConveyance", "LC BTC/BTE",
                        "PAN card name", "Bank Account Number", "IFSC Code", "Bank Name", "Currency", "Other Currency", "Beneficiary Name", "Pan Number", "Global FMV", "Swift Code" };

                    Dictionary<string, object> PanelrowData = new Dictionary<string, object>();
                    for (int i = 0; i < PanelColumnNames.Count; i++)
                    {
                        if (PanelcolumnsToInclude.Contains(PanelColumnNames[i]))
                        {
                            PanelrowData[PanelColumnNames[i]] = row.Cells[i].Value;
                        }
                    }
                    PaginatedResult<Attachment> attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet4.Id.Value, row.Id.Value, null);

                    List<Dictionary<string, object>> PanelattachmentsList = new();
                    if (attachments.Data != null || attachments.Data.Count > 0)
                    {
                        foreach (var attachment in attachments.Data)
                        {
                            long AID = (long)attachment.Id;
                            Attachment file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet4.Id.Value, AID);
                            Dictionary<string, object> attachmentInfo = new()
                            {
                                { "Name", file.Name },
                                { "Id", file.Id },
                                { "base64" , SheetHelper.UrlToBaseValue(file.Url) }
                            };
                            PanelattachmentsList.Add(attachmentInfo);
                        }
                        PanelrowData["Attachments"] = PanelattachmentsList;
                    }
                    PaneleventDetails.Add(PanelrowData);
                }
            }
            Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
            List<string> SlideKitColumnNames = new List<string>();
            foreach (Column column in sheet5.Columns)
            {
                SlideKitColumnNames.Add(column.Title);
            }
            foreach (var row in sheet5.Rows)
            {
                if (row.Cells.Any(c => c.DisplayValue == eventId))
                {
                    List<string> SlideKitcolumnsToInclude = new List<string> { "SlideKit ID", "HCP Name", "MIS", "HcpType", "Slide Kit Type", "SlideKit Document" };

                    Dictionary<string, object> SlideKitrowData = new Dictionary<string, object>();
                    for (int i = 0; i < SlideKitColumnNames.Count; i++)
                    {
                        if (SlideKitcolumnsToInclude.Contains(SlideKitColumnNames[i]))
                        {
                            SlideKitrowData[SlideKitColumnNames[i]] = row.Cells[i].Value;
                        }
                    }
                    PaginatedResult<Attachment> attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet5.Id.Value, row.Id.Value, null);

                    List<Dictionary<string, object>> SlideKitattachmentsList = new();
                    if (attachments.Data != null || attachments.Data.Count > 0)
                    {
                        foreach (var attachment in attachments.Data)
                        {
                            long AID = (long)attachment.Id;
                            Attachment file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet5.Id.Value, AID);
                            Dictionary<string, object> attachmentInfo = new()
                            {
                                { "Name", file.Name },
                                { "Id", file.Id },
                                { "base64" , SheetHelper.UrlToBaseValue(file.Url) }
                            };
                            SlideKitattachmentsList.Add(attachmentInfo);
                        }
                        SlideKitrowData["Attachments"] = SlideKitattachmentsList;
                    }
                    SlideKiteventDetails.Add(SlideKitrowData);
                }
            }
            Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
            List<string> ExpenseColumnNames = new List<string>();
            foreach (Column column in sheet6.Columns)
            {
                ExpenseColumnNames.Add(column.Title);
            }
            foreach (var row in sheet6.Rows)
            {
                if (row.Cells.Any(c => c.DisplayValue == eventId))
                {
                    List<string> ExpensecolumnsToInclude = new List<string> { "Expenses ID", "Expense", "BTC/BTE", "Amount Excluding Tax", "Amount" };

                    Dictionary<string, object> ExpenserowData = new Dictionary<string, object>();
                    for (int i = 0; i < ExpenseColumnNames.Count; i++)
                    {
                        if (ExpensecolumnsToInclude.Contains(ExpenseColumnNames[i]))
                        {
                            ExpenserowData[ExpenseColumnNames[i]] = row.Cells[i].Value;
                        }
                    }
                    ExpenseeventDetails.Add(ExpenserowData);
                }
            }
            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);
            List<string> DeviationscolumnNames = new List<string>();
            foreach (Column column in sheet7.Columns)
            {
                DeviationscolumnNames.Add(column.Title);
            }
            foreach (var row in sheet7.Rows)
            {
                if (row.Cells.Any(c => c.DisplayValue == eventId))
                {
                    List<string> DeviationscolumnsToInclude = new List<string> { "Deviation Type" };

                    Dictionary<string, object> DeviationsattachmentInfo = new Dictionary<string, object>();
                    for (int i = 0; i < DeviationscolumnNames.Count; i++)
                    {
                        if (DeviationscolumnsToInclude.Contains(DeviationscolumnNames[i]))
                        {
                            var val = row.Cells[i].Value.ToString();
                            var attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet7.Id.Value, row.Id.Value, null);
                            foreach (var attachment in attachments.Data)
                            {
                                var AID = (long)attachment.Id;
                                var file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet7.Id.Value, AID);
                                DeviationsattachmentInfo[val] = SheetHelper.UrlToBaseValue(file.Url);
                            }
                        }
                    }
                    DeviationsattachmentsList.Add(DeviationsattachmentInfo);
                }
            }


            Sheet sheet8 = (Sheet)SheetHelper.GetSheetById(smartsheet, sheetId8);
            List<string> EventRequestBeneficiary = new List<string>();
            foreach (Column column in sheet8.Columns)
            {
                EventRequestBeneficiary.Add(column.Title);
            }
            foreach (var row in sheet8.Rows.Where(row => row.Cells.Any(c => c.DisplayValue == eventId)))
            {
                Dictionary<string, object> BeneficiaryrowData = new Dictionary<string, object>
                {
                    { "BenfId", GetValueOrDefaults(row, "Benf Id") },
                    { "EventType", GetValueOrDefaults(row, "EventType") },
                    { "EventDate", GetValueOrDefaults(row, "EventDate") },
                    { "VenueName", GetValueOrDefaults(row, "VenueName") },
                    { "State", GetValueOrDefaults(row, "State") },
                    { "IsFacilityCharges", GetValueOrDefaults(row, "Facility Charges") },
                    { "IsAnesthetistRequired", GetValueOrDefaults(row, "Anesthetist Required?") },
                    { "TypeOfBeneficiary", GetValueOrDefaults(row, "Type of Beneficiary") },
                    { "Currency", GetValueOrDefaults(row, "Currency") },
                    { "OtherCurrency", GetValueOrDefaults(row, "Other Currency") },
                    { "BeneficiaryName", GetValueOrDefaults(row, "Beneficiary Name") },
                    { "BankAccountNumber", GetValueOrDefaults(row, "Bank Account Number") },
                    { "BankName", GetValueOrDefaults(row, "Bank Name") },
                    { "PANcardName", GetValueOrDefaults(row, "PAN card name") },
                    { "PanNumber", GetValueOrDefaults(row, "Pan Number") },
                    { "IFSCCode", GetValueOrDefaults(row, "IFSC Code") },
                    { "EmailId", GetValueOrDefaults(row, "Email Id") },
                    { "SwiftCode", GetValueOrDefaults(row, "City") },
                    { "IBNNumber", GetValueOrDefaults(row, "IBN Number") },
                    { "TaxResidenceCertificateDate", GetValueOrDefaults(row, "Tax Residence Certificate Date") }

                };
                PaginatedResult<Attachment> attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet8.Id.Value, row.Id.Value, null);

                List<Dictionary<string, object>> SlideKitattachmentsList = new();
                if (attachments.Data != null || attachments.Data.Count > 0)
                {
                    foreach (var attachment in attachments.Data)
                    {
                        long AID = (long)attachment.Id;
                        Attachment file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet8.Id.Value, AID);
                        Dictionary<string, object> attachmentInfo = new()
                            {
                                { "Name", file.Name },
                                { "Id", file.Id },
                                { "base64" , SheetHelper.UrlToBaseValue(file.Url) }
                            };
                        SlideKitattachmentsList.Add(attachmentInfo);
                    }
                    BeneficiaryrowData["Attachments"] = SlideKitattachmentsList;
                }



                BeneficiarteventDetails.Add(BeneficiaryrowData);
            }


            string? GetValueOrDefaults(Row row, string columnName)
            {
                int columnIndex = EventRequestBeneficiary.IndexOf(columnName);
                if (columnIndex >= 0 && columnIndex < row.Cells.Count)
                {
                    return row.Cells[columnIndex].Value?.ToString();
                }
                return null;
            }


            Sheet sheet9 = (Sheet)SheetHelper.GetSheetById(smartsheet, sheetId9);
            List<string> EventRequestProductBrandsList = new List<string>();
            foreach (Column column in sheet9.Columns)
            {
                EventRequestProductBrandsList.Add(column.Title);
            }
            foreach (var row in sheet9.Rows.Where(row => row.Cells.Any(c => c.DisplayValue == eventId)))
            {
                Dictionary<string, object> ProductBrandsListrowData = new Dictionary<string, object>
                {
                    { "Product ID", GetValuesOrDefaults(row, "Product ID") },
                    { "EventType", GetValuesOrDefaults(row, "EventType") },
                    { "EventDate", GetValuesOrDefaults(row, "EventDate") },
                    { "EventTopic", GetValuesOrDefaults(row, "Event Topic") },
                    { "ProductBrand", GetValuesOrDefaults(row, "Product Brand") },
                    { "ProductName", GetValuesOrDefaults(row, "Product Name") },
                    { "NoOfSamplesRequired", GetValuesOrDefaults(row, "No of Samples Required") },


                };
                PaginatedResult<Attachment> attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet9.Id.Value, row.Id.Value, null);

                List<Dictionary<string, object>> ProductBrandsListattachmentsList = new();
                if (attachments.Data != null || attachments.Data.Count > 0)
                {
                    foreach (var attachment in attachments.Data)
                    {
                        long AID = (long)attachment.Id;
                        Attachment file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet9.Id.Value, AID);
                        Dictionary<string, object> attachmentInfo = new()
                            {
                                { "Name", file.Name },
                                { "Id", file.Id },
                                { "base64" , SheetHelper.UrlToBaseValue(file.Url) }
                            };
                        ProductBrandsListattachmentsList.Add(attachmentInfo);
                    }
                    ProductBrandsListrowData["Attachments"] = ProductBrandsListattachmentsList;
                }
                ProductBrandsListeventDetails.Add(ProductBrandsListrowData);
            }

            string? GetValuesOrDefaults(Row row, string columnName)
            {
                int columnIndex = EventRequestProductBrandsList.IndexOf(columnName);
                if (columnIndex >= 0 && columnIndex < row.Cells.Count)
                {
                    return row.Cells[columnIndex].Value?.ToString();
                }
                return null;
            }











            resultData["eventDetails"] = eventDetails;
            resultData["Files"] = attachmentInfoFiles;
            resultData["Brands"] = BrandseventDetails;
            resultData["Invitees"] = InviteeseventDetails;
            resultData["PanelDetails"] = PaneleventDetails;
            resultData["SlideKitSelection"] = SlideKiteventDetails;
            resultData["ExpenseSelection"] = ExpenseeventDetails;
            resultData["Deviation"] = DeviationsattachmentsList;
            resultData["EventRequestBeneficiary"] = BeneficiarteventDetails;
            resultData["EventRequestProductBrandsList"] = ProductBrandsListeventDetails;

            return Ok(resultData);
        }

        [HttpPut("UpdateClassIPreEvent")]
        public IActionResult UpdateClassIPreEvent(UpdateDataForClassI formDataList)
        {
            var eventId = formDataList.EventDetails.Id;
            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);
            #region
            //StringBuilder addedBrandsData = new();
            //StringBuilder addedInviteesData = new();
            //StringBuilder addedMEnariniInviteesData = new();
            //StringBuilder addedHcpData = new();
            //StringBuilder addedSlideKitData = new();
            //StringBuilder addedExpences = new();

            //int addedSlideKitDataNo = 1;
            //int addedHcpDataNo = 1;
            //int addedInviteesDataNo = 1;
            //int addedInviteesDataNoforMenarini = 1;
            //int addedBrandsDataNo = 1;
            //int addedExpencesNo = 1;


            //foreach (var formdata in formDataList.ExpenseSelection)
            //{
            //    string rowData = $"{addedExpencesNo}. {formdata.Expense} | AmountExcludingTax: {formdata.ExpenseAmountExcludingTax}| Amount: {formdata.ExpenseAmountIncludingTax} | {formdata.ExpenseType}";
            //    addedExpences.AppendLine(rowData);
            //    addedExpencesNo++;

            //}

            //string Expense = addedExpences.ToString();
            //foreach (var formdata in formDataList.SlideKitSelection)
            //{
            //    string rowData = $"{addedSlideKitDataNo}. {formdata.HcpName} | {formdata.SlideKitType}";
            //    addedSlideKitData.AppendLine(rowData);
            //    addedSlideKitDataNo++;
            //}
            //string slideKit = addedSlideKitData.ToString();
            //foreach (var formdata in formDataList.BrandSelection)
            //{
            //    string rowData = $"{addedBrandsDataNo}. {formdata.brandName} | {formdata.projectId} | {formdata.percentageAllocation}";
            //    addedBrandsData.AppendLine(rowData);
            //    addedBrandsDataNo++;
            //}
            //string brand = addedBrandsData.ToString();
            //foreach (var formdata in formDataList.InviteeSelection)
            //{
            //    if (formdata.InviteeFrom == "Menarini Employees")
            //    {
            //        string row = $"{addedInviteesDataNoforMenarini}. {formdata.Name}";
            //        addedMEnariniInviteesData.AppendLine(row);
            //        addedInviteesDataNoforMenarini++;
            //    }
            //    else
            //    {
            //        string rowData = $"{addedInviteesDataNo}. {formdata.Name}";
            //        addedInviteesData.AppendLine(rowData);
            //        addedInviteesDataNo++;
            //    }

            //}
            //string Invitees = addedInviteesData.ToString();
            //string MenariniInvitees = addedMEnariniInviteesData.ToString();
            //foreach (var formdata in formDataList.PanelSelection)
            //{

            //    string rowData = $"{addedHcpDataNo}. {formdata.HcpRole} |{formdata.HcpName} | Honr.Amt: {formdata.HonarariumAmountIncludingTax} |Trav.&Acc.Amt: {formdata.TravelAmountIncludingTax + formdata.AccomdationIncludingTax} ";
            //    addedHcpData.AppendLine(rowData);
            //    addedHcpDataNo++;

            //}
            //string HCP = addedHcpData.ToString();
            #endregion

            Row? targetRow = sheet1.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == formDataList.EventDetails.Id));
            long UpdatedId = 0;
            if (targetRow != null)
            {
                try
                {
                    Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.EventDetails.EventTopic });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Start Time"), Value = formDataList.EventDetails.StartTime });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "End Time"), Value = formDataList.EventDetails.EndTime });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Venue Name"), Value = formDataList.EventDetails.VenueName });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Type"), Value = formDataList.EventDetails.EventType });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Date"), Value = formDataList.EventDetails.EventDate });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "State"), Value = formDataList.EventDetails.State });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "City"), Value = formDataList.EventDetails.City });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Meeting Type"), Value = formDataList.EventDetails.MeetingType });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Brands"), Value = formDataList.EventDetails.Brands });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Panelists"), Value = formDataList.EventDetails.Panelists });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "SlideKits"), Value = formDataList.EventDetails.SlideKits });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Invitees"), Value = formDataList.EventDetails.Invitees });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "MIPL Invitees"), Value = formDataList.EventDetails.MIPLInvitees });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Expenses"), Value = formDataList.EventDetails.Expenses });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "BTE Expense Details"), Value = formDataList.EventDetails.ExpenseDataBTE });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, " Total Expense BTC"), Value = formDataList.EventDetails.TotalExpenseBTC });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense BTE"), Value = formDataList.EventDetails.TotalExpenseBTE });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Honorarium Amount"), Value = formDataList.EventDetails.TotalHonorariumAmount });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel & Accommodation Amount"), Value = formDataList.EventDetails.TotalTravelAccommodationAmount });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Accommodation Amount"), Value = formDataList.EventDetails.TotalAccomodationAmount });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Budget Amount"), Value = formDataList.EventDetails.TotalBudget });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Local Conveyance"), Value = formDataList.EventDetails.TotalLocalConveyance });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel Amount"), Value = formDataList.EventDetails.TotalTravelAmount });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense"), Value = formDataList.EventDetails.TotalExpense });

                    IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRow });
                    long uId = updatedRow[0].Id.Value;
                    UpdatedId = uId;
                    if (formDataList.EventDetails.IsFilesUpload == "Yes")
                    {
                        foreach (var p in formDataList.EventDetails.Files)
                        {

                            string[] words = p.FileBase64.Split(':');
                            string r = words[0];
                            string q = words[1];
                            string name = r.Split(".")[0];
                            string filePath = SheetHelper.testingFile(q, name);
                            Row addedRow = updatedRow[0];
                            if (p.Id != null)
                            {
                                Attachment Updateattachment = smartsheet.SheetResources.AttachmentResources.VersioningResources.AttachNewVersion(
                                    sheet1.Id.Value, (long)p.Id, filePath, "application/msword");
                            }
                            else
                            {
                                Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                   sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                            }

                            if (System.IO.File.Exists(filePath))
                            {
                                SheetHelper.DeleteFile(filePath);
                            }
                        }

                    }
                }
                catch (Exception ex)
                {
                    Log.Error($"Error occured on EventSettlementController method {ex.Message} at {DateTime.Now}");
                    Log.Error(ex.StackTrace);
                    return BadRequest(ex.Message);
                }
            }
           

            if (formDataList.IsDeviationUpload == "Yes")
            {
                List<string> DeviationNames = new List<string>();
                foreach (var p in formDataList.DeviationDetails)
                {

                    string[] words = p.DeviationFile.Split(':')[0].Split("*");
                    string r = words[1];
                    DeviationNames.Add(r);
                }
                foreach (var pp in formDataList.DeviationDetails)
                {
                    foreach (var deviationname in DeviationNames)
                    {
                        string file = deviationname.Split(".")[0];

                        if (pp.DeviationFile.Split(':')[0].Split("*")[1] == deviationname)
                        {
                            try
                            {
                                Row newRow7 = new()
                                {
                                    Cells = new List<Cell>()
                                };
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.EventDetails.EventTopic });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Type"), Value = formDataList.EventDetails.EventType });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Date"), Value = formDataList.EventDetails.EventDate });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Start Time"), Value = formDataList.EventDetails.StartTime });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "End Time"), Value = formDataList.EventDetails.EndTime });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Venue Name"), Value = formDataList.EventDetails.VenueName });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "City"), Value = formDataList.EventDetails.City });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "State"), Value = formDataList.EventDetails.State });

                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "MIS Code"), Value = SheetHelper.MisCodeCheck(pp.MisCode) });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HCP Name"), Value = pp.HcpName });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Honorarium Amount"), Value = pp.HonorariumAmountExcludingTax });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Travel & Accommodation Amount"), Value = pp.TravelorAccomodationAmountExcludingTax });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Other Expenses"), Value = pp.OtherExpenseAmountExcludingTax });

                                if (file == "30DaysDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:30DaysDeviationFile").Value });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = "Yes" });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = formDataList.EventDetails.EventOpen30dayscount });
                                }
                                else if (file == "7DaysDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:5DaysDeviationFile").Value });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = "Yes" });

                                }
                                else if (file == "ExpenseExcludingTax")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:ExpenseExcludingTax").Value });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "PRE-F&B Expense Excluding Tax"), Value = "Yes" });
                                }
                                else if (file.Contains("Travel_Accomodation3LExceededFile"))
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:Travel_Accomodation3LExceededFile").Value });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Travel/Accomodation 3,00,000 Exceeded Trigger"), Value = "Yes" });//formDataList.class1.FB_Expense_Excluding_Tax });
                                }
                                else if (file.Contains("TrainerHonorarium12LExceededFile"))
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:TrainerHonorarium12LExceededFile").Value/*"Honorarium Aggregate Limit of 12,00,000 is Exceeded"*/ });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Trainer Honorarium 12,00,000 Exceeded Trigger"), Value = "Yes" }); //formDataList.class1.FB_Expense_Excluding_Tax });
                                }
                                else if (file.Contains("HCPHonorarium6LExceededFile"))
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:HCPHonorarium6LExceededFile").Value/*"Honorarium Aggregate Limit of 6,00,000 is Exceeded"*/ });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HCP Honorarium 6,00,000 Exceeded Trigger"), Value = "Yes" }); // formDataList.class1.FB_Expense_Excluding_Tax });
                                }

                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.EventDetails.Sales_Head });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formDataList.EventDetails.FinanceHead });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Name"), Value = formDataList.EventDetails.InitiatorName });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.EventDetails.Initiator_Email });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Coordinator"), Value = formDataList.EventDetails.SalesCoordinator });

                                IList<Row> addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });

                                int j = 1;
                                foreach (var p in formDataList.DeviationDetails)
                                {
                                    string[] nameSplit = p.DeviationFile.Split("*");
                                    string[] words = nameSplit[1].Split(':');
                                    string r = words[0];
                                    string q = words[1];
                                    if (deviationname == r)
                                    {
                                        string name = nameSplit[0];
                                        string filePath = SheetHelper.testingFile(q, name);
                                        Row addedRow = addeddeviationrow[0];
                                        Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                                        // Attachment attachmentinmain = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet1.Id.Value, targetRow.Id.Value, filePath, "application/msword");
                                        j++;
                                        if (System.IO.File.Exists(filePath))
                                        {
                                            SheetHelper.DeleteFile(filePath);
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                return BadRequest(ex.Message);
                            }

                        }

                    }

                }
            }

            if (formDataList.BrandSelection.Count > 0)
            {
                Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
                List<long> rowIdsToDelete = new List<long>();
                foreach (Row row in sheet2.Rows)
                {
                    if (row.Cells.Any(cell => cell.DisplayValue == formDataList.EventDetails.Id))
                    {
                        rowIdsToDelete.Add((long)row.Id);
                    }
                }
                if (rowIdsToDelete.Count > 0)
                {
                    smartsheet.SheetResources.RowResources.DeleteRows(sheet2.Id.Value, rowIdsToDelete.ToArray(), true);
                }
                List<Row> newRows2 = new();
                foreach (var formdata in formDataList.BrandSelection)
                {
                    Row newRow2 = new()
                    {
                        Cells = new List<Cell>()
                        {
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "% Allocation"), Value = formdata.PercentageAllocation },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Brands"), Value = formdata.BrandName },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Project ID"), Value = formdata.ProjectId },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "EventId/EventRequestId"), Value =  formDataList.EventDetails.Id }
                        }
                    };

                    newRows2.Add(newRow2);
                }
                smartsheet.SheetResources.RowResources.AddRows(sheet2.Id.Value, newRows2.ToArray());



            }

            if (formDataList.PanelSelection.Count > 0)
            {

                Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
                List<long> rowIdsToDelete = new List<long>();
                foreach (Row row in sheet4.Rows)
                {
                    if (row.Cells.Any(cell => cell.DisplayValue == formDataList.EventDetails.Id))
                    {
                        rowIdsToDelete.Add((long)row.Id);
                    }
                }
                if (rowIdsToDelete.Count > 0)
                {
                    smartsheet.SheetResources.RowResources.DeleteRows(sheet4.Id.Value, rowIdsToDelete.ToArray(), true);
                }
                foreach (var formData in formDataList.PanelSelection)
                {
                    Row newRow1 = new()
                    {
                        Cells = new List<Cell>()
                    };
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HcpRole"), Value = formData.HcpRole });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "MISCode"), Value = SheetHelper.MisCodeCheck(formData.MisCode) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Travel"), Value = formData.TravelAmountIncludingTax });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TotalSpend"), Value = formData.FinalAmount });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Accomodation"), Value = formData.AccomdationIncludingTax });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "LocalConveyance"), Value = formData.LocalConveyanceIncludingTax });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "SpeakerCode"), Value = formData.SpeakerCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TrainerCode"), Value = formData.TrainerCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HonorariumRequired"), Value = formData.HonorariumRequired });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "AgreementAmount"), Value = formData.AgreementAmount });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HonorariumAmount"), Value = formData.HonarariumAmountIncludingTax });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Speciality"), Value = formData.Speciality });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Topic"), Value = formDataList.EventDetails.EventTopic });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Type"), Value = formDataList.EventDetails.EventType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Venue name"), Value = formDataList.EventDetails.VenueName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Date Start"), Value = formDataList.EventDetails.EventDate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCPName"), Value = formData.HcpName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PAN card name"), Value = formData.PanCardName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "ExpenseType"), Value = formData.ExpenseType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Bank Account Number"), Value = formData.BankAccountNumber });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Bank Name"), Value = formData.BankName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "IFSC Code"), Value = formData.IFSCCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "FCPA Date"), Value = formData.FcpaIssueDate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Currency"), Value = formData.Currency });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Honorarium Amount Excluding Tax"), Value = formData.HonarariumAmountExcludingTax });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Travel Excluding Tax"), Value = formData.TravelExcludingTax });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Accomodation Excluding Tax"), Value = formData.AccomdationExcludingTax });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Local Conveyance Excluding Tax"), Value = formData.LocalConveyanceExcludingTax });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "LC BTC/BTE"), Value = formData.LcBtcorBte });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Travel BTC/BTE"), Value = formData.TravelBtcorBte });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Accomodation BTC/BTE"), Value = formData.AccomodationBtcorBte });

                    if (formData.Currency == "Others")
                    {
                        newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Other Currency"), Value = formData.OtherCurrencyType });
                    }
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Beneficiary Name"), Value = formData.BeneficiaryName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Pan Number"), Value = formData.PanNumber });

                    if (formData.HcpRole == "Others")
                    {

                        newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Other Type"), Value = formData.OthersType });
                    }

                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Tier"), Value = formData.Tier });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCP Type"), Value = formData.GOorNGO });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PresentationDuration"), Value = formData.PresentationDuration });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PanelSessionPreparationDuration"), Value = formData.PanelSessionPreperationDuration });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PanelDiscussionDuration"), Value = formData.PanelDisscussionDuration });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "QASessionDuration"), Value = formData.QaSessionDuration });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "BriefingSession"), Value = formData.BriefingSession });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TotalSessionHours"), Value = formData.TotalSessionHours });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Rationale"), Value = formData.Rationale });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "EventId/EventRequestId"), Value = eventId });


                    IList<Row> row = smartsheet.SheetResources.RowResources.AddRows(sheet4.Id.Value, new Row[] { newRow1 });
                    if (formData.IsFilesUpload == "Yes")
                    {
                        foreach (string p in formData.Files)
                        {
                            string[] words = p.Split(':');
                            string r = words[0];
                            string q = words[1];
                            string name = r.Split(".")[0];
                            string filePath = SheetHelper.testingFile(q, name);
                            Row addedRow = row[0];
                            Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                   sheet4.Id.Value, addedRow.Id.Value, filePath, "application/msword");


                            if (System.IO.File.Exists(filePath))
                            {
                                SheetHelper.DeleteFile(filePath);
                            }
                        }
                    }


                }


            }

            if (formDataList.SlideKitSelection.Count > 0)
            {
                Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
                List<long> rowIdsToDelete = new List<long>();
                foreach (Row row in sheet5.Rows)
                {
                    if (row.Cells.Any(cell => cell.DisplayValue == formDataList.EventDetails.Id))
                    {
                        rowIdsToDelete.Add((long)row.Id);
                    }
                }
                if (rowIdsToDelete.Count > 0)
                {
                    smartsheet.SheetResources.RowResources.DeleteRows(sheet5.Id.Value, rowIdsToDelete.ToArray(), true);
                }

                foreach (var formdata in formDataList.SlideKitSelection)
                {
                    Row newRow5 = new()
                    {
                        Cells = new List<Cell>()
                    };

                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "MIS"), Value = SheetHelper.MisCodeCheck(formdata.MisCode) });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "Slide Kit Type"), Value = formdata.SlideKitType });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "SlideKit Document"), Value = formdata.SlideKitOption });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "EventId/EventRequestId"), Value = eventId });


                    IList<Row> row = smartsheet.SheetResources.RowResources.AddRows(sheet5.Id.Value, new Row[] { newRow5 });
                    if (formdata.IsFilesUpload == "Yes")
                    {
                        foreach (string p in formdata.Files)
                        {
                            string[] words = p.Split(':');
                            string r = words[0];
                            string q = words[1];
                            string name = r.Split(".")[0];
                            string filePath = SheetHelper.testingFile(q, name);
                            Row addedRow = row[0];
                            Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                   sheet5.Id.Value, addedRow.Id.Value, filePath, "application/msword");


                            if (System.IO.File.Exists(filePath))
                            {
                                SheetHelper.DeleteFile(filePath);
                            }
                        }
                    }
                }
            }

            if (formDataList.InviteeSelection.Count > 0)
            {
                Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
                List<long> rowIdsToDelete = new List<long>();
                foreach (Row row in sheet3.Rows)
                {
                    if (row.Cells.Any(cell => cell.DisplayValue == formDataList.EventDetails.Id))
                    {
                        rowIdsToDelete.Add((long)row.Id);
                    }
                }
                if (rowIdsToDelete.Count > 0)
                {
                    smartsheet.SheetResources.RowResources.DeleteRows(sheet3.Id.Value, rowIdsToDelete.ToArray(), true);
                }
                List<Row> newRows3 = new();
                foreach (var formdata in formDataList.InviteeSelection)
                {
                    Row newRow3 = new()
                    {
                        Cells = new List<Cell>()
                        {
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "HCPName"), Value = formdata.Name },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Designation"), Value = formdata.Designation },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Employee Code"), Value = formdata.EmployeeCode },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "LocalConveyance"), Value = formdata.IsLocalConveyance },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "BTC/BTE"), Value = formdata.LocalConveyanceType },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "LcAmount"), Value = formdata.LocalConveyanceAmountIncludingTax },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Lc Amount Excluding Tax"), Value = formdata.LocalConveyanceAmountExcludingTax },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "EventId/EventRequestId"), Value = eventId },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Invitee Source"), Value = formdata.InviteeFrom },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "HCP Type"), Value = formdata.HCPType },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Speciality"), Value = formdata.Speciality },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "MISCode"), Value = SheetHelper.MisCodeCheck(formdata.MisCode )},
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Topic"), Value = formDataList.EventDetails.EventTopic },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Type"), Value = formDataList.EventDetails.EventType },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Venue name"), Value = formDataList.EventDetails.VenueName },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Date Start"), Value = formDataList.EventDetails.EventDate },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event End Date"), Value = formDataList.EventDetails.EventDate }

                        }
                    };
                    newRows3.Add(newRow3);
                }
                smartsheet.SheetResources.RowResources.AddRows(sheet3.Id.Value, newRows3.ToArray());



            }

            if (formDataList.ExpenseSelection.Count > 0)
            {
                Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
                List<long> rowIdsToDelete = new List<long>();
                foreach (Row row in sheet6.Rows)
                {
                    if (row.Cells.Any(cell => cell.DisplayValue == formDataList.EventDetails.Id))
                    {
                        rowIdsToDelete.Add((long)row.Id);
                    }
                }
                if (rowIdsToDelete.Count > 0)
                {
                    smartsheet.SheetResources.RowResources.DeleteRows(sheet6.Id.Value, rowIdsToDelete.ToArray(), true);
                }
                List<Row> newRows6 = new();
                foreach (var formdata in formDataList.ExpenseSelection)
                {
                    Row newRow6 = new()
                    {
                        Cells = new List<Cell>()
                        {
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Expense"), Value = formdata.Expense },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "EventId/EventRequestID"), Value = eventId },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount Excluding Tax"), Value = formdata.ExpenseAmountExcludingTax },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount"), Value = formdata.ExpenseAmountIncludingTax },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formdata.ExpenseType },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Topic"), Value = formDataList.EventDetails.EventTopic },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Type"), Value = formDataList.EventDetails.EventType },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Venue name"), Value = formDataList.EventDetails.VenueName },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Date Start"), Value = formDataList.EventDetails.EventDate },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event End Date"), Value = formDataList.EventDetails.EventDate }
                        }
                    };
                    newRows6.Add(newRow6);
                }
                smartsheet.SheetResources.RowResources.AddRows(sheet6.Id.Value, newRows6.ToArray());



            }

            return Ok(new
            { Message = "Updated Successfully" });
        }

        [HttpPut("UpdateHandsOnPreEvent")]
        public IActionResult UpdateHandsOnPreEvent(UpdateDataForHandsOnTraining formDataList)
        {
            var eventId = formDataList.EventDetails.Id;
            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);

            Row? targetRow = sheet1.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == formDataList.EventDetails.Id));
            long UpdatedId = 0;
            if (targetRow != null)
            {
                try
                {
                    Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.EventDetails.EventName });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Start Time"), Value = formDataList.EventDetails.StartTime });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "End Time"), Value = formDataList.EventDetails.EndTime });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Venue Name"), Value = formDataList.EventDetails.VenueName });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Type"), Value = formDataList.EventDetails.EventType });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Date"), Value = formDataList.EventDetails.EventDate });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "State"), Value = formDataList.EventDetails.State });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "City"), Value = formDataList.EventDetails.City });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Brands"), Value = formDataList.EventDetails.Brands });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Panelists"), Value = formDataList.EventDetails.Panelists });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "SlideKits"), Value = formDataList.EventDetails.SlideKits });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Invitees"), Value = formDataList.EventDetails.Invitees });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "MIPL Invitees"), Value = formDataList.EventDetails.MIPLInvitees });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Expenses"), Value = formDataList.EventDetails.Expenses });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, " Total Expense BTC"), Value = formDataList.EventDetails.TotalExpenseBTC });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense BTE"), Value = formDataList.EventDetails.TotalExpenseBTE });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Honorarium Amount"), Value = formDataList.EventDetails.TotalHonorariumAmount });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel & Accommodation Amount"), Value = formDataList.EventDetails.TotalTravelAccommodationAmount });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Accommodation Amount"), Value = formDataList.EventDetails.TotalAccomodationAmount });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Budget Amount"), Value = formDataList.EventDetails.TotalBudget });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Local Conveyance"), Value = formDataList.EventDetails.TotalLocalConveyance });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel Amount"), Value = formDataList.EventDetails.TotalTravelAmount });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense"), Value = formDataList.EventDetails.TotalExpense });

                    IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRow });
                    long uId = updatedRow[0].Id.Value;
                    UpdatedId = uId;
                    if (formDataList.EventDetails.IsFilesUpload == "Yes")
                    {
                        foreach (var p in formDataList.EventDetails.Files)
                        {

                            string[] words = p.FileBase64.Split(':');
                            string r = words[0];
                            string q = words[1];
                            string name = r.Split(".")[0];
                            string filePath = SheetHelper.testingFile(q, name);
                            Row addedRow = updatedRow[0];
                            if (p.Id != null)
                            {
                                Attachment Updateattachment = smartsheet.SheetResources.AttachmentResources.VersioningResources.AttachNewVersion(
                                    sheet1.Id.Value, (long)p.Id, filePath, "application/msword");
                            }
                            else
                            {
                                Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                   sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                            }

                            if (System.IO.File.Exists(filePath))
                            {
                                SheetHelper.DeleteFile(filePath);
                            }
                        }

                    }
                }
                catch (Exception ex)
                {
                    Log.Error($"Error occured on EventSettlementController method {ex.Message} at {DateTime.Now}");
                    Log.Error(ex.StackTrace);
                    return BadRequest(ex.Message);
                }
            }

            if (formDataList.IsDeviationUpload == "Yes")
            {
                List<string> DeviationNames = new List<string>();
                foreach (var p in formDataList.DeviationDetails)
                {

                    string[] words = p.DeviationFile.Split(':')[0].Split("*");
                    string r = words[1];
                    DeviationNames.Add(r);
                }
                foreach (var pp in formDataList.DeviationDetails)
                {
                    foreach (var deviationname in DeviationNames)
                    {
                        string file = deviationname.Split(".")[0];

                        if (pp.DeviationFile.Split(':')[0].Split("*")[1] == deviationname)
                        {
                            try
                            {
                                Row newRow7 = new()
                                {
                                    Cells = new List<Cell>()
                                };
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.EventDetails.EventName });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Type"), Value = formDataList.EventDetails.EventType });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Date"), Value = formDataList.EventDetails.EventDate });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Start Time"), Value = formDataList.EventDetails.StartTime });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "End Time"), Value = formDataList.EventDetails.EndTime });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Venue Name"), Value = formDataList.EventDetails.VenueName });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "City"), Value = formDataList.EventDetails.City });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "State"), Value = formDataList.EventDetails.State });

                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "MIS Code"), Value = SheetHelper.MisCodeCheck(pp.MisCode) });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HCP Name"), Value = pp.HcpName });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Honorarium Amount"), Value = pp.HonorariumAmountExcludingTax });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Travel & Accommodation Amount"), Value = pp.TravelorAccomodationAmountExcludingTax });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Other Expenses"), Value = pp.OtherExpenseAmountExcludingTax });

                                if (file == "30DaysDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:30DaysDeviationFile").Value });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = "Yes" });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = formDataList.EventDetails.EventOpen30dayscount });
                                }
                                else if (file == "7DaysDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:5DaysDeviationFile").Value });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = "Yes" });

                                }
                                else if (file == "ExpenseExcludingTax")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:ExpenseExcludingTax").Value });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "PRE-F&B Expense Excluding Tax"), Value = "Yes" });
                                }
                                else if (file.Contains("Travel_Accomodation3LExceededFile"))
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:Travel_Accomodation3LExceededFile").Value });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Travel/Accomodation 3,00,000 Exceeded Trigger"), Value = "Yes" });//formDataList.class1.FB_Expense_Excluding_Tax });
                                }
                                else if (file.Contains("TrainerHonorarium12LExceededFile"))
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:TrainerHonorarium12LExceededFile").Value/*"Honorarium Aggregate Limit of 12,00,000 is Exceeded"*/ });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Trainer Honorarium 12,00,000 Exceeded Trigger"), Value = "Yes" }); //formDataList.class1.FB_Expense_Excluding_Tax });
                                }
                                else if (file.Contains("HCPHonorarium6LExceededFile"))
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:HCPHonorarium6LExceededFile").Value/*"Honorarium Aggregate Limit of 6,00,000 is Exceeded"*/ });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HCP Honorarium 6,00,000 Exceeded Trigger"), Value = "Yes" }); // formDataList.class1.FB_Expense_Excluding_Tax });
                                }

                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.EventDetails.Sales_Head });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formDataList.EventDetails.FinanceHead });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Name"), Value = formDataList.EventDetails.InitiatorName });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.EventDetails.Initiator_Email });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Coordinator"), Value = formDataList.EventDetails.SalesCoordinator });
                                IList<Row> addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });

                                int j = 1;
                                foreach (var p in formDataList.DeviationDetails)
                                {
                                    string[] nameSplit = p.DeviationFile.Split("*");
                                    string[] words = nameSplit[1].Split(':');
                                    string r = words[0];
                                    string q = words[1];
                                    if (deviationname == r)
                                    {
                                        string name = nameSplit[0];
                                        string filePath = SheetHelper.testingFile(q, name);
                                        Row addedRow = addeddeviationrow[0];
                                        Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                                        Attachment attachmentinmain = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet1.Id.Value, targetRow.Id.Value, filePath, "application/msword");
                                        j++;
                                        if (System.IO.File.Exists(filePath))
                                        {
                                            SheetHelper.DeleteFile(filePath);
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                return BadRequest(ex.Message);
                            }

                        }

                    }

                }
            }

            if (formDataList.BrandSelection.Count > 0)
            {
                Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
                List<long> rowIdsToDelete = new List<long>();
                foreach (Row row in sheet2.Rows)
                {
                    if (row.Cells.Any(cell => cell.DisplayValue == formDataList.EventDetails.Id))
                    {
                        rowIdsToDelete.Add((long)row.Id);
                    }
                }
                if (rowIdsToDelete.Count > 0)
                {
                    smartsheet.SheetResources.RowResources.DeleteRows(sheet2.Id.Value, rowIdsToDelete.ToArray(), true);
                }
                List<Row> newRows2 = new();
                foreach (var formdata in formDataList.BrandSelection)
                {
                    Row newRow2 = new()
                    {
                        Cells = new List<Cell>()
                        {
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "% Allocation"), Value = formdata.PercentageAllocation },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Brands"), Value = formdata.BrandName },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Project ID"), Value = formdata.ProjectId },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "EventId/EventRequestId"), Value =  formDataList.EventDetails.Id }
                        }
                    };

                    newRows2.Add(newRow2);
                }
                smartsheet.SheetResources.RowResources.AddRows(sheet2.Id.Value, newRows2.ToArray());



            }

            if (formDataList.PanelSelection.Count > 0)
            {

                Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
                List<long> rowIdsToDelete = new List<long>();
                foreach (Row row in sheet4.Rows)
                {
                    if (row.Cells.Any(cell => cell.DisplayValue == formDataList.EventDetails.Id))
                    {
                        rowIdsToDelete.Add((long)row.Id);
                    }
                }
                if (rowIdsToDelete.Count > 0)
                {
                    smartsheet.SheetResources.RowResources.DeleteRows(sheet4.Id.Value, rowIdsToDelete.ToArray(), true);
                }
                foreach (var formData in formDataList.PanelSelection)
                {
                    Row newRow1 = new()
                    {
                        Cells = new List<Cell>()
                    };
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HcpRole"), Value = formData.HcpRole });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "MISCode"), Value = SheetHelper.MisCodeCheck(formData.MisCode) });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Travel"), Value = formData.TravelAmountIncludingTax });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TotalSpend"), Value = formData.FinalAmount });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Accomodation"), Value = formData.AccomdationIncludingTax });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "LocalConveyance"), Value = formData.LocalConveyanceIncludingTax });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "SpeakerCode"), Value = formData.SpeakerCode });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TrainerCode"), Value = formData.TrainerCode });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HonorariumRequired"), Value = formData.HonorariumRequired });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "AgreementAmount"), Value = formData.AgreementAmount });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HonorariumAmount"), Value = formData.HonarariumAmountIncludingTax });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Speciality"), Value = formData.Speciality });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Topic"), Value = formDataList.EventDetails.EventTopic });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Type"), Value = formDataList.EventDetails.EventType });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Venue name"), Value = formDataList.EventDetails.VenueName });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Date Start"), Value = formDataList.EventDetails.EventDate });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCPName"), Value = formData.HcpName });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PAN card name"), Value = formData.PanCardName });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "ExpenseType"), Value = formData.ExpenseType });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Bank Account Number"), Value = formData.BankAccountNumber });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Bank Name"), Value = formData.BankName });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "IFSC Code"), Value = formData.IFSCCode });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "FCPA Date"), Value = formData.FcpaIssueDate });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Currency"), Value = formData.Currency });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Honorarium Amount Excluding Tax"), Value = formData.HonarariumAmountExcludingTax });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Travel Excluding Tax"), Value = formData.TravelExcludingTax });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Accomodation Excluding Tax"), Value = formData.AccomdationExcludingTax });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Local Conveyance Excluding Tax"), Value = formData.LocalConveyanceExcludingTax });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "LC BTC/BTE"), Value = formData.LcBtcorBte });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Travel BTC/BTE"), Value = formData.TravelBtcorBte });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Accomodation BTC/BTE"), Value = formData.AccomodationBtcorBte });

                    //if (formData.Currency == "Others")
                    //{
                    //    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Other Currency"), Value = formData.OtherCurrencyType });
                    //}
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Beneficiary Name"), Value = formData.BeneficiaryName });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Pan Number"), Value = formData.PanNumber });

                    //if (formData.HcpRole == "Others")
                    //{

                    //    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Other Type"), Value = formData.OthersType });
                    //}

                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Tier"), Value = formData.Tier });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCP Type"), Value = formData.GOorNGO });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PresentationDuration"), Value = formData.PresentationDuration });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PanelSessionPreparationDuration"), Value = formData.PanelSessionPreperationDuration });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PanelDiscussionDuration"), Value = formData.PanelDisscussionDuration });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "QASessionDuration"), Value = formData.QaSessionDuration });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "BriefingSession"), Value = formData.BriefingSession });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TotalSessionHours"), Value = formData.TotalSessionHours });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Rationale"), Value = formData.Rationale });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "EventId/EventRequestId"), Value = eventId });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "MISCode"), Value = SheetHelper.MisCodeCheck(formData.MISCode) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HcpRole"), Value = formData.HCPRole });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCPName"), Value = formData.TrainerName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TrainerCode"), Value = formData.TrainerCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Qualification"), Value = formData.TrainerQualification });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Country"), Value = formData.TrainerCountry });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Speciality"), Value = formData.Speciality });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Tier"), Value = formData.TrainerCategory });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCP Type"), Value = formData.TrainerType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Rationale"), Value = formData.Rationale });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "FCPA Date"), Value = formData.FCPAIssueDate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HonorariumRequired"), Value = formData.IsHonorariumApplicable });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Annual Trainer Agreement Valid?"), Value = formData.IsAnnualTrainerAgreementValid });

                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PresentationDuration"), Value = formData.Presentation_Speaking_WorkshopDuration });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PanelSessionPreparationDuration"), Value = formData.DevelopmentofPresentationPanelSessionPreparation });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PanelDiscussionDuration"), Value = formData.PaneldiscussionSessionduration });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "QASessionDuration"), Value = formData.QASession });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "BriefingSession"), Value = formData.Speaker_TrainerBriefing });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TotalSessionHours"), Value = formData.TotalNoOfHours });

                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "LC BTC/BTE"), Value = formData.IsLCBTC_BTE });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Accomodation BTC/BTE"), Value = formData.IsAccomodationBTC_BTE });


                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Honorarium Amount Excluding Tax"), Value = formData.HonorariumAmountexcludingTax });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HonorariumAmount"), Value = formData.HonorariumAmountincludingTax });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "AgreementAmount"), Value = formData.AgreementAmount });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "YTD Spend Including Current Event"), Value = formData.YTDspendIncludingCurrentEvent });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Global FMV"), Value = formData.IsGlobalFMVCheck });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "ExpenseType"), Value = formData.ExpenseType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Mode of Travel"), Value = formData.TravelSelection });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Travel BTC/BTE"), Value = formData.IsExpenseBTC_BTE });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Travel Excluding Tax"), Value = formData.TravelAmountExcludingTax });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Travel"), Value = formData.TravelAmountIncludingTax });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Accomodation Excluding Tax"), Value = formData.AccomodationAmountExcludingTax });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Accomodation"), Value = formData.AccomodationAmountIncludingTax });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Local Conveyance Excluding Tax"), Value = formData.LocalConveyanceAmountexcludingTax });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "LocalConveyance"), Value = formData.LocalConveyanceAmountincludingTax });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Travel/Accomodation Spend Including Current Event"), Value = formData.TravelandAccomodationspendincludingcurrentevent });


                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Currency"), Value = formData.BenificiaryDetailsData.Currency });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Other Currency"), Value = formData.BenificiaryDetailsData.EnterCurrencyType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Beneficiary Name"), Value = formData.BenificiaryDetailsData.BenificiaryName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Bank Account Number"), Value = formData.BenificiaryDetailsData.BankAccountNumber });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Bank Name"), Value = formData.BenificiaryDetailsData.BankName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PAN card name"), Value = formData.BenificiaryDetailsData.NameasPerPAN });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Pan Number"), Value = formData.BenificiaryDetailsData.PANCardNumber });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "IFSC Code"), Value = formData.BenificiaryDetailsData.IFSCCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Email Id"), Value = formData.BenificiaryDetailsData.EmailID });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "IBN Number"), Value = formData.BenificiaryDetailsData.IbnNumber });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Swift Code"), Value = formData.BenificiaryDetailsData.SwiftCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Tax Residence Certificate"), Value = formData.BenificiaryDetailsData.TaxResidenceCertificateDate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "AgreementAmount"), Value = formData.AgreementAmount });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Speciality"), Value = formData.Speciality });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Topic"), Value = formDataList.EventDetails.EventName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Type"), Value = formDataList.EventDetails.EventType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Venue name"), Value = formDataList.EventDetails.VenueName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Date Start"), Value = formDataList.EventDetails.EventDate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event End Date"), Value = formDataList.EventDetails.EventDate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TotalSpend"), Value = formData.FinalAmount });

                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "EventId/EventRequestId"), Value = eventId });

                    IList<Row> row = smartsheet.SheetResources.RowResources.AddRows(sheet4.Id.Value, new Row[] { newRow1 });
                    foreach (string p in formData.TrainerFiles)
                    {
                        string[] words = p.Split(':');
                        string r = words[0];
                        string q = words[1];
                        string name = r.Split(".")[0];
                        string filePath = SheetHelper.testingFile(q, name);
                        Row addedRow = row[0];
                        Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                               sheet4.Id.Value, addedRow.Id.Value, filePath, "application/msword");



                        if (System.IO.File.Exists(filePath))
                        {
                            SheetHelper.DeleteFile(filePath);
                        }
                    }


                }


            }

            if (formDataList.SlideKitSelection.Count > 0)
            {
                Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
                List<long> rowIdsToDelete = new List<long>();
                foreach (Row row in sheet5.Rows)
                {
                    if (row.Cells.Any(cell => cell.DisplayValue == formDataList.EventDetails.Id))
                    {
                        rowIdsToDelete.Add((long)row.Id);
                    }
                }
                if (rowIdsToDelete.Count > 0)
                {
                    smartsheet.SheetResources.RowResources.DeleteRows(sheet5.Id.Value, rowIdsToDelete.ToArray(), true);
                }

                foreach (var formdata in formDataList.SlideKitSelection)
                {
                    Row newRow5 = new()
                    {
                        Cells = new List<Cell>()
                    };
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "MIS"), Value = SheetHelper.MisCodeCheck(formdata.MISCode) });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "HCP Name"), Value = formdata.TrainerName });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "Slide Kit Type"), Value = formdata.SlideKitSelection });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "SlideKit Document"), Value = formdata.SlideKitDocument });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "EventId/EventRequestId"), Value = eventId });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "Fillers Indication"), Value = formdata.IndicatorsForFillers });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "Threads Indication"), Value = formdata.IndicatorsForThreads });

                    IList<Row> row = smartsheet.SheetResources.RowResources.AddRows(sheet5.Id.Value, new Row[] { newRow5 });
                    if (formdata.IsUpload == "Yes")
                    {
                        string[] words = formdata.DocToUpload.Split(':');
                        string r = words[0];
                        string q = words[1];
                        string name = r.Split(".")[0];
                        string filePath = SheetHelper.testingFile(q, name);
                        Row addedRow = row[0];
                        Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                               sheet5.Id.Value, addedRow.Id.Value, filePath, "application/msword");



                        if (System.IO.File.Exists(filePath))
                        {
                            SheetHelper.DeleteFile(filePath);
                        }
                    }
                }
            }

            if (formDataList.InviteeSelection.Count > 0)
            {
                Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
                List<long> rowIdsToDelete = new List<long>();
                foreach (Row row in sheet3.Rows)
                {
                    if (row.Cells.Any(cell => cell.DisplayValue == formDataList.EventDetails.Id))
                    {
                        rowIdsToDelete.Add((long)row.Id);
                    }
                }
                if (rowIdsToDelete.Count > 0)
                {
                    smartsheet.SheetResources.RowResources.DeleteRows(sheet3.Id.Value, rowIdsToDelete.ToArray(), true);
                }
                List<Row> newRows3 = new();
                foreach (var formdata in formDataList.InviteeSelection)
                {
                    Row newRow3 = new()
                    {
                        Cells = new List<Cell>()
                    };
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "MISCode"), Value = SheetHelper.MisCodeCheck(formdata.MisCode) });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "HCPName"), Value = formdata.AttenderName });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "HCP Type"), Value = formdata.AttenderType });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Designation"), Value = formdata.Designation });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Employee Code"), Value = formdata.EmployeeCode });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "LocalConveyance"), Value = formdata.IsLocalConveyance });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "BTC/BTE"), Value = formdata.IsBtcorBte });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "LcAmount"), Value = formdata.LocalConveyanceAmountIncludingTax });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Lc Amount Excluding Tax"), Value = formdata.LocalConveyanceAmountExcludingTax });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "EventId/EventRequestId"), Value = eventId });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Invitee Source"), Value = formdata.InviteedFrom });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Qualification"), Value = formdata.Qualification });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Experience"), Value = formdata.Experiance });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Topic"), Value = formDataList.EventDetails.EventName });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Type"), Value = formDataList.EventDetails.EventType });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Venue name"), Value = formDataList.EventDetails.VenueName });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Date Start"), Value = formDataList.EventDetails.EventDate });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event End Date"), Value = formDataList.EventDetails.EventDate });

                    IList<Row> row = smartsheet.SheetResources.RowResources.AddRows(sheet3.Id.Value, new Row[] { newRow3 });

                    foreach (string p in formdata.AttenderFiles)
                    {
                        string[] words = p.Split(':');
                        string r = words[0];
                        string q = words[1];
                        string name = r.Split(".")[0];
                        string filePath = SheetHelper.testingFile(q, name);
                        Row addedRow = row[0];
                        Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                               sheet3.Id.Value, addedRow.Id.Value, filePath, "application/msword");



                        if (System.IO.File.Exists(filePath))
                        {
                            SheetHelper.DeleteFile(filePath);
                        }
                    }
                }
            }

            if (formDataList.ExpenseSelection.Count > 0)
            {
                Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
                List<long> rowIdsToDeleteInExpense = new List<long>();
                foreach (Row row in sheet6.Rows)
                {
                    if (row.Cells.Any(cell => cell.DisplayValue == formDataList.EventDetails.Id))
                    {
                        rowIdsToDeleteInExpense.Add((long)row.Id);
                    }
                }
                if (rowIdsToDeleteInExpense.Count > 0)
                {
                    smartsheet.SheetResources.RowResources.DeleteRows(sheet6.Id.Value, rowIdsToDeleteInExpense.ToArray(), true);
                }
                List<Row> newRows6 = new();
                foreach (var formdata in formDataList.ExpenseSelection)
                {
                    Row newRow6 = new()
                    {
                        Cells = new List<Cell>()
                        {
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Expense"), Value = formdata.ExpenseType },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "EventId/EventRequestID"), Value = eventId },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount Excluding Tax"), Value = formdata.ExpenseAmountExcludingTax },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount"), Value = formdata.ExpenseAmountIncludingTax },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formdata.IsBtcorBte },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Topic"), Value = formDataList.EventDetails.EventName },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Type"), Value = formDataList.EventDetails.EventType },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Venue name"), Value = formDataList.EventDetails.VenueName },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Date Start"), Value = formDataList.EventDetails.EventDate },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event End Date"), Value = formDataList.EventDetails.EventDate }
                        }
                    };
                    newRows6.Add(newRow6);
                }
                smartsheet.SheetResources.RowResources.AddRows(sheet6.Id.Value, newRows6.ToArray());



            }

            if (formDataList.EventDetails.VenueBenificiaryDetailsData != null)
            {
                Sheet sheet8 = SheetHelper.GetSheetById(smartsheet, sheetId8);
                Row newRow8 = new()
                {
                    Cells = new List<Cell>()
                };

                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "EventId/EventRequestId"), Value = eventId });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "EventType"), Value = formDataList.EventDetails.EventType });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "EventDate"), Value = formDataList.EventDetails.EventDate });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "VenueName"), Value = formDataList.EventDetails.VenueName });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "City"), Value = formDataList.EventDetails.City });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "State"), Value = formDataList.EventDetails.State });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Other Currency"), Value = formDataList.EventDetails.VenueBenificiaryDetailsData.EnterCurrencyType });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Beneficiary Name"), Value = formDataList.EventDetails.VenueBenificiaryDetailsData.BenificiaryName });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Bank Account Number"), Value = formDataList.EventDetails.VenueBenificiaryDetailsData.BankAccountNumber });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Facility Charges"), Value = formDataList.EventDetails.IsVenueFacilityCharges });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Anesthetist Required?"), Value = formDataList.EventDetails.IsAnesthetistRequired });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Currency"), Value = formDataList.EventDetails.VenueBenificiaryDetailsData.Currency });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Bank Name"), Value = formDataList.EventDetails.VenueBenificiaryDetailsData.BankName });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "PAN card name"), Value = formDataList.EventDetails.VenueBenificiaryDetailsData.NameasPerPAN });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Pan Number"), Value = formDataList.EventDetails.VenueBenificiaryDetailsData.PANCardNumber });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "IFSC Code"), Value = formDataList.EventDetails.VenueBenificiaryDetailsData.IFSCCode });
                //newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, ""), Value = formDataList.HandsOnTraining.BenificiaryDetailsData.IbnNumber });
                //newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formDataList.HandsOnTraining.BenificiaryDetailsData.SwiftCode });
                //newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formDataList.HandsOnTraining.BenificiaryDetailsData.TaxResidenceCertificateDate });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Email Id"), Value = formDataList.EventDetails.VenueBenificiaryDetailsData.EmailID });

                smartsheet.SheetResources.RowResources.AddRows(sheet8.Id.Value, new Row[] { newRow8 });
            }

            if (formDataList.EventDetails.AnaestheticBenificiaryDetailsData != null)
            {
                Sheet sheet8 = SheetHelper.GetSheetById(smartsheet, sheetId8);
                Row newRow8 = new()
                {
                    Cells = new List<Cell>()
                };

                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "EventId/EventRequestId"), Value = eventId });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "EventType"), Value = formDataList.EventDetails.EventType });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "EventDate"), Value = formDataList.EventDetails.EventDate });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "VenueName"), Value = formDataList.EventDetails.VenueName });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "City"), Value = formDataList.EventDetails.City });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "State"), Value = formDataList.EventDetails.State });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Other Currency"), Value = formDataList.EventDetails.AnaestheticBenificiaryDetailsData.EnterCurrencyType });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Beneficiary Name"), Value = formDataList.EventDetails.AnaestheticBenificiaryDetailsData.BenificiaryName });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Bank Account Number"), Value = formDataList.EventDetails.AnaestheticBenificiaryDetailsData.BankAccountNumber });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Facility Charges"), Value = formDataList.EventDetails.IsVenueFacilityCharges });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Anesthetist Required?"), Value = formDataList.EventDetails.IsAnesthetistRequired });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Currency"), Value = formDataList.EventDetails.AnaestheticBenificiaryDetailsData.Currency });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Bank Name"), Value = formDataList.EventDetails.AnaestheticBenificiaryDetailsData.BankName });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "PAN card name"), Value = formDataList.EventDetails.AnaestheticBenificiaryDetailsData.NameasPerPAN });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Pan Number"), Value = formDataList.EventDetails.AnaestheticBenificiaryDetailsData.PANCardNumber });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "IFSC Code"), Value = formDataList.EventDetails.AnaestheticBenificiaryDetailsData.IFSCCode });
                //newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, ""), Value = formDataList.HandsOnTraining.BenificiaryDetailsData.IbnNumber });
                //newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formDataList.HandsOnTraining.BenificiaryDetailsData.SwiftCode });
                //newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formDataList.HandsOnTraining.BenificiaryDetailsData.TaxResidenceCertificateDate });
                newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Email Id"), Value = formDataList.EventDetails.AnaestheticBenificiaryDetailsData.EmailID });

                smartsheet.SheetResources.RowResources.AddRows(sheet8.Id.Value, new Row[] { newRow8 });
            }



            return Ok(new
            { Message = "Updated Successfully" });
        }

        [HttpPut("UpdateStallFabricationPreEvent")]
        public IActionResult UpdateStallFabricationPreEvent(UpdateDataForStall formDataList)
        {

            var eventId = formDataList.EventDetails.Id;
            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);
            Row? targetRow = sheet1.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == formDataList.EventDetails.Id));
            long UpdatedId = 0;

            if (targetRow != null)
            {
                try
                {
                    Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.EventDetails.EventTopic });

                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Type"), Value = formDataList.EventDetails.EventType });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Date"), Value = formDataList.EventDetails.EventStartDate });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "End Date"), Value = formDataList.EventDetails.EventEndDate });

                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Brands"), Value = formDataList.EventDetails.BrandsData });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "BTE Expense Details"), Value = formDataList.EventDetails.ExpenseDataBTE });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Class III Event Code"), Value = formDataList.EventDetails.ClassIIIEventCode });
                    //updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Invitees"), Value = formDataList.EventDetails.Invitees });
                    //updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "MIPL Invitees"), Value = formDataList.EventDetails.MIPLInvitees });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Expenses"), Value = formDataList.EventDetails.ExpenseData });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, " Total Expense BTC"), Value = formDataList.EventDetails.TotalExpenseBTC });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense BTE"), Value = formDataList.EventDetails.TotalExpenseBTE });
                    //updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Honorarium Amount"), Value = formDataList.EventDetails.TotalHonorariumAmount });
                    //updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel & Accommodation Amount"), Value = formDataList.EventDetails.TotalTravelAccommodationAmount });
                    //updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Accommodation Amount"), Value = formDataList.EventDetails.TotalAccomodationAmount });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Budget Amount"), Value = formDataList.EventDetails.TotalBudget });
                    //updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Local Conveyance"), Value = formDataList.EventDetails.TotalLocalConveyance });
                    //updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel Amount"), Value = formDataList.EventDetails.TotalTravelAmount });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense"), Value = formDataList.EventDetails.TotalExpense });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "IsAdvanceRequired"), Value = formDataList.EventDetails.IsAdvanceRequired });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Advance Amount"), Value = formDataList.EventDetails.AdvanceAmount });




                    IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRow });
                    long uId = updatedRow[0].Id.Value;
                    UpdatedId = uId;
                    if (formDataList.EventDetails.IsFilesUpload == "Yes")
                    {
                        foreach (var p in formDataList.EventDetails.Files)
                        {

                            string[] words = p.FileBase64.Split(':');
                            string r = words[0];
                            string q = words[1];
                            string name = r.Split(".")[0];
                            string filePath = SheetHelper.testingFile(q, name);
                            Row addedRow = updatedRow[0];
                            if (p.Id != null)
                            {
                                Attachment Updateattachment = smartsheet.SheetResources.AttachmentResources.VersioningResources.AttachNewVersion(
                                    sheet1.Id.Value, (long)p.Id, filePath, "application/msword");
                            }
                            else
                            {
                                Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                   sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                            }

                            if (System.IO.File.Exists(filePath))
                            {
                                SheetHelper.DeleteFile(filePath);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log.Error($"Error occured on EventSettlementController method {ex.Message} at {DateTime.Now}");
                    Log.Error(ex.StackTrace);
                    return BadRequest(ex.Message);
                }
            }
            if (formDataList.IsDeviationUpload == "Yes")
            {
                List<string> DeviationNames = new List<string>();
                foreach (var p in formDataList.DeviationDetails)
                {

                    string[] words = p.DeviationFile.Split(':')[0].Split("*");
                    string r = words[1];
                    DeviationNames.Add(r);
                }
                foreach (var pp in formDataList.DeviationDetails)
                {
                    foreach (var deviationname in DeviationNames)
                    {
                        string file = deviationname.Split(".")[0];

                        if (pp.DeviationFile.Split(':')[0].Split("*")[1] == deviationname)
                        {
                            try
                            {
                                Row newRow7 = new()
                                {
                                    Cells = new List<Cell>()
                                };
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.EventDetails.EventTopic });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Type"), Value = formDataList.EventDetails.EventType });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Date"), Value = formDataList.EventDetails.EventStartDate });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "End Date"), Value = formDataList.EventDetails.EventEndDate });

                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "MIS Code"), Value = SheetHelper.MisCodeCheck(pp.MisCode) });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HCP Name"), Value = pp.HcpName });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Honorarium Amount"), Value = pp.HonorariumAmountExcludingTax });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Travel & Accommodation Amount"), Value = pp.TravelorAccomodationAmountExcludingTax });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Other Expenses"), Value = pp.OtherExpenseAmountExcludingTax });

                                if (file == "30DaysDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:30DaysDeviationFile").Value });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = "Yes" });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = formDataList.EventDetails.EventOpen30dayscount });
                                }
                                else if (file == "7DaysDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:5DaysDeviationFile").Value });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = "Yes" });

                                }
                                else if (file == "ExpenseExcludingTax")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:ExpenseExcludingTax").Value });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "PRE-F&B Expense Excluding Tax"), Value = "Yes" });
                                }
                                else if (file.Contains("Travel_Accomodation3LExceededFile"))
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:Travel_Accomodation3LExceededFile").Value });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Travel/Accomodation 3,00,000 Exceeded Trigger"), Value = "Yes" });//formDataList.class1.FB_Expense_Excluding_Tax });
                                }
                                else if (file.Contains("TrainerHonorarium12LExceededFile"))
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:TrainerHonorarium12LExceededFile").Value/*"Honorarium Aggregate Limit of 12,00,000 is Exceeded"*/ });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Trainer Honorarium 12,00,000 Exceeded Trigger"), Value = "Yes" }); //formDataList.class1.FB_Expense_Excluding_Tax });
                                }
                                else if (file.Contains("HCPHonorarium6LExceededFile"))
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:HCPHonorarium6LExceededFile").Value/*"Honorarium Aggregate Limit of 6,00,000 is Exceeded"*/ });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HCP Honorarium 6,00,000 Exceeded Trigger"), Value = "Yes" }); // formDataList.class1.FB_Expense_Excluding_Tax });
                                }

                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.EventDetails.Sales_Head });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formDataList.EventDetails.FinanceHead });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Name"), Value = formDataList.EventDetails.InitiatorName });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.EventDetails.Initiator_Email });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Coordinator"), Value = formDataList.EventDetails.SalesCoordinator });
                                IList<Row> addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });

                                int j = 1;
                                foreach (var p in formDataList.DeviationDetails)
                                {
                                    string[] nameSplit = p.DeviationFile.Split("*");
                                    string[] words = nameSplit[1].Split(':');
                                    string r = words[0];
                                    string q = words[1];
                                    if (deviationname == r)
                                    {
                                        string name = nameSplit[0];
                                        string filePath = SheetHelper.testingFile(q, name);
                                        Row addedRow = addeddeviationrow[0];
                                        Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                                        Attachment attachmentinmain = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet1.Id.Value, targetRow.Id.Value, filePath, "application/msword");
                                        j++;
                                        if (System.IO.File.Exists(filePath))
                                        {
                                            SheetHelper.DeleteFile(filePath);
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                return BadRequest(ex.Message);
                            }

                        }

                    }

                }
            }


            if (formDataList.BrandSelection.Count > 0)
            {
                Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
                List<long> rowIdsToDelete = new List<long>();
                foreach (Row row in sheet2.Rows)
                {
                    if (row.Cells.Any(cell => cell.DisplayValue == formDataList.EventDetails.Id))
                    {
                        rowIdsToDelete.Add((long)row.Id);
                    }
                }
                if (rowIdsToDelete.Count > 0)
                {
                    smartsheet.SheetResources.RowResources.DeleteRows(sheet2.Id.Value, rowIdsToDelete.ToArray(), true);
                }
                List<Row> newRows2 = new();
                foreach (var formdata in formDataList.BrandSelection)
                {
                    Row newRow2 = new()
                    {
                        Cells = new List<Cell>()
                        {
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "% Allocation"), Value = formdata.PercentageAllocation },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Brands"), Value = formdata.BrandName },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Project ID"), Value = formdata.ProjectId },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "EventId/EventRequestId"), Value =  formDataList.EventDetails.Id }
                        }
                    };

                    newRows2.Add(newRow2);
                }
                smartsheet.SheetResources.RowResources.AddRows(sheet2.Id.Value, newRows2.ToArray());



            }


            if (formDataList.ExpenseSelection.Count > 0)
            {
                Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
                List<long> rowIdsToDelete = new List<long>();
                foreach (Row row in sheet6.Rows)
                {
                    if (row.Cells.Any(cell => cell.DisplayValue == formDataList.EventDetails.Id))
                    {
                        rowIdsToDelete.Add((long)row.Id);
                    }
                }
                if (rowIdsToDelete.Count > 0)
                {
                    smartsheet.SheetResources.RowResources.DeleteRows(sheet6.Id.Value, rowIdsToDelete.ToArray(), true);
                }
                List<Row> newRows6 = new();
                foreach (var formdata in formDataList.ExpenseSelection)
                {
                    Row newRow6 = new()
                    {
                        Cells = new List<Cell>()
                        {
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Expense"), Value = formdata.Expense },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "EventId/EventRequestID"), Value = eventId },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount Excluding Tax"), Value = formdata.ExpenseAmountExcludingTax },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount"), Value = formdata.ExpenseAmountIncludingTax },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formdata.ExpenseType },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Topic"), Value = formDataList.EventDetails.EventTopic },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Type"), Value = formDataList.EventDetails.EventType },
                            //new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Venue name"), Value = formDataList.EventDetails.VenueName },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Date Start"), Value = formDataList.EventDetails.EventStartDate },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event End Date"), Value = formDataList.EventDetails.EventEndDate }
                        }
                    };
                    newRows6.Add(newRow6);
                }
                smartsheet.SheetResources.RowResources.AddRows(sheet6.Id.Value, newRows6.ToArray());



            }

            return Ok(new
            { Message = "Updated Successfully" });






        }

        [HttpPut("UpdateStallFabricationPostEvent")]
        public IActionResult UpdateStallFabricationPostEvent(UpdateEventSettlementData formDataList)
        {
            var eventId = formDataList.EventId;
            Sheet sheet10 = SheetHelper.GetSheetById(smartsheet, sheetId10);
            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);
            Row? targetRow = sheet10.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == eventId));
            long UpdatedId = 0;
            if (targetRow != null)
            {
                try
                {
                    Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet10, "Budget Amount"), Value = formDataList.TotalBudgetAmount });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet10, "Total Actual"), Value = formDataList.TotalActuals });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet10, "Advance Utilized For Event"), Value = formDataList.AdvanceUtilizedForEvents });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet10, "Pay Back Amount To Company"), Value = formDataList.PayBackAmountToCompany });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet10, "Total Expense"), Value = formDataList.TotalExpenses });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet10, "Additional Amount Needed To Pay For Initiator"), Value = formDataList.AdditionalAmountNeededToPayForInitiator });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet10, "Advance Amount"), Value = formDataList.AdvanceProvided });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet10, "Total Expense BTC"), Value = formDataList.TotalBtcAmount });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet10, "Total Expense BTE"), Value = formDataList.TotalBteAmount });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet10, "Expense Details"), Value = formDataList.ExpenseDetails });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet10, "Total Expense Details"), Value = formDataList.TotalExpenseDetails });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet10, "Advance Details"), Value = formDataList.AdvanceDetails });





                    IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet10.Id.Value, new Row[] { updateRow });
                    long uId = updatedRow[0].Id.Value;
                    UpdatedId = uId;
                    if (formDataList.IsFilesUpload == "Yes")
                    {
                        foreach (var p in formDataList.Files)
                        {

                            string[] words = p.FileBase64.Split(':');
                            string r = words[0];
                            string q = words[1];
                            string name = r.Split(".")[0];
                            string filePath = SheetHelper.testingFile(q, name);
                            Row addedRow = updatedRow[0];
                            if (p.Id != null && p.Id != 0)
                            {
                                Attachment Updateattachment = smartsheet.SheetResources.AttachmentResources.VersioningResources.AttachNewVersion(
                                    sheet10.Id.Value, (long)p.Id, filePath, "application/msword");
                            }
                            else
                            {
                                Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                   sheet10.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                            }

                            if (System.IO.File.Exists(filePath))
                            {
                                SheetHelper.DeleteFile(filePath);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log.Error($"Error occured on EventSettlementController method {ex.Message} at {DateTime.Now}");
                    Log.Error(ex.StackTrace);
                    return BadRequest(ex.Message);
                }

            }
            if (formDataList.IsDeviationUpload == "Yes")
            {
                List<string> DeviationNames = new List<string>();
                foreach (var p in formDataList.DeviationDetails)
                {

                    string[] words = p.DeviationFile.Split(':')[0].Split("*");
                    string r = words[1];
                    DeviationNames.Add(r);
                }
                foreach (var pp in formDataList.DeviationDetails)
                {
                    foreach (var deviationname in DeviationNames)
                    {
                        string file = deviationname.Split(".")[0];

                        if (pp.DeviationFile.Split(':')[0].Split("*")[1] == deviationname)
                        {
                            try
                            {
                                Row newRow7 = new()
                                {
                                    Cells = new List<Cell>()
                                };
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.EventTopic });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Type"), Value = formDataList.EventType });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Date"), Value = formDataList.EventStartDate });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "End Date"), Value = formDataList.EventEndDate });

                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "MIS Code"), Value = SheetHelper.MisCodeCheck(pp.MisCode) });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HCP Name"), Value = pp.HcpName });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Honorarium Amount"), Value = pp.HonorariumAmountExcludingTax });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Travel & Accommodation Amount"), Value = pp.TravelorAccomodationAmountExcludingTax });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Other Expenses"), Value = pp.OtherExpenseAmountExcludingTax });
                                if (file == "30DaysDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "POST- Beyond45Days Deviation Date Trigger"), Value = "Yes" });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInEventSettlement:30DaysDeviationFile").Value });
                                }
                                else if (file == "Lessthan5InviteesDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "POST-Lessthan5Invitees Deviation Trigger"), Value = "Yes" });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInEventSettlement:Lessthan5InviteesDeviationFile").Value });
                                }
                                else if (file == "ExcludingGSTDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "POST-Deviation Excluding GST?"), Value = "Yes" });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInEventSettlement:ExcludingGSTDeviationFile").Value });
                                }
                                else if (file == "Change in venue")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "POST-Deviation Change in venue trigger"), Value = "Yes" });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = file });

                                }
                                else if (file == "Change in topic")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "POST-Deviation Change in topic trigger"), Value = "Yes" });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = file });

                                }
                                else if (file == "Change in speaker")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "POST-Deviation Change in speaker trigger"), Value = "Yes" });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = file });

                                }
                                else if (file == "Attendees not captured in photographs")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "POST-Deviation Attendees not captured trigger"), Value = "Yes" });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = file });

                                }
                                else if (file == "Speaker not captured in photographs")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "POST-Deviation Speaker not captured trigger"), Value = "Yes" });//POST-Deviation Speaker not captured  trigger
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = file });
                                }
                                else
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "POST-Deviation Other Deviation Trigger"), Value = "Yes" });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Others" });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Other Deviation Type"), Value = file });
                                }
                                //if (file == "30DaysDeviationFile")
                                //{
                                //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:30DaysDeviationFile").Value });
                                //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = "Yes" });
                                //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = SheetHelper.NumCheck(formDataList.EventOpen30dayscount) });
                                //}
                                //else if (file == "7DaysDeviationFile")
                                //{
                                //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:5DaysDeviationFile").Value });
                                //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = "Yes" });

                                //}
                                //else if (file == "ExpenseExcludingTax")
                                //{
                                //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:ExpenseExcludingTax").Value });
                                //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "PRE-F&B Expense Excluding Tax"), Value = "Yes" });
                                //}
                                //else if (file.Contains("Travel_Accomodation3LExceededFile"))
                                //{
                                //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:Travel_Accomodation3LExceededFile").Value });
                                //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Travel/Accomodation 3,00,000 Exceeded Trigger"), Value = "Yes" });//formDataList.class1.FB_Expense_Excluding_Tax });
                                //}
                                //else if (file.Contains("TrainerHonorarium12LExceededFile"))
                                //{
                                //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:TrainerHonorarium12LExceededFile").Value/*"Honorarium Aggregate Limit of 12,00,000 is Exceeded"*/ });
                                //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Trainer Honorarium 12,00,000 Exceeded Trigger"), Value = "Yes" }); //formDataList.class1.FB_Expense_Excluding_Tax });
                                //}
                                //else if (file.Contains("HCPHonorarium6LExceededFile"))
                                //{
                                //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:HCPHonorarium6LExceededFile").Value/*"Honorarium Aggregate Limit of 6,00,000 is Exceeded"*/ });
                                //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HCP Honorarium 6,00,000 Exceeded Trigger"), Value = "Yes" }); // formDataList.class1.FB_Expense_Excluding_Tax });
                                //}

                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.Sales_Head });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formDataList.FinanceHead });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Name"), Value = formDataList.InitiatorName });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.Initiator_Email });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Coordinator"), Value = formDataList.SalesCoordinator });
                                IList<Row> addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });

                                int j = 1;
                                foreach (var p in formDataList.DeviationDetails)
                                {
                                    string[] nameSplit = p.DeviationFile.Split("*");
                                    string[] words = nameSplit[1].Split(':');
                                    string r = words[0];
                                    string q = words[1];
                                    if (deviationname == r)
                                    {
                                        string name = nameSplit[0];
                                        string filePath = SheetHelper.testingFile(q, name);
                                        Row addedRow = addeddeviationrow[0];
                                        Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                                        Attachment attachmentinmain = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet10.Id.Value, targetRow.Id.Value, filePath, "application/msword");
                                        j++;
                                        if (System.IO.File.Exists(filePath))
                                        {
                                            SheetHelper.DeleteFile(filePath);
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                return BadRequest(ex.Message);
                            }

                        }

                    }

                }
            }

            return Ok(new
            { Message = "Updated Successfully" });
        }

        [HttpPut("UpdateRegectionHonorariumData")]
        public IActionResult UpdateRegectionHonorariumData(UpdateHonorariumPaymentListPh2 formData)
        {
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId11);
            var eventId = formData.EventId;
            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);
            Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == eventId));
            if (targetRow != null)
            {
                if (formData.IsFilesUpload == "Yes")
                {
                    foreach (var p in formData.Files)
                    {

                        string[] words = p.FileBase64.Split(':');
                        string r = words[0];
                        string q = words[1];
                        string name = r.Split(".")[0];
                        string filePath = SheetHelper.testingFile(q, name);
                        //Row addedRow = updatedRow[0];
                        if (p.Id != null)
                        {
                            Attachment Updateattachment = smartsheet.SheetResources.AttachmentResources.VersioningResources.AttachNewVersion(
                                sheet.Id.Value, (long)p.Id, filePath, "application/msword");
                        }
                        else
                        {
                            Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                               sheet.Id.Value, targetRow.Id.Value, filePath, "application/msword");
                        }

                        if (System.IO.File.Exists(filePath))
                        {
                            SheetHelper.DeleteFile(filePath);
                        }
                    }
                }
            }
            if (formData.IsDeviationUpload == "Yes")
            {
                try
                {
                    Row newRow7 = new()
                    {
                        Cells = new List<Cell>()
                    };
                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formData.EventTopic });
                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Type"), Value = formData.EventType });
                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Date"), Value = formData.EventDate });
                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Start Time"), Value = formData.StartTime });
                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "End Time"), Value = formData.EndTime });
                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Venue Name"), Value = formData.VenueName });
                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "City"), Value = formData.City });
                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "State"), Value = formData.State });
                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInHonorarium:5WorkingdaysDeviationDateTrigger").Value });
                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HON-5Workingdays Deviation Date Trigger"), Value = formData.IsDeviationUpload });
                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formData.SalesHeadEmail });
                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formData.FinanceHead });
                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Name"), Value = formData.InitiatorName });
                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formData.InitiatorEmail });
                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Coordinator"), Value = formData.SalesCoordinator });
                    IList<Row> addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });

                    foreach (string p in formData.DeviationFiles)
                    {
                        string[] words = p.Split(':');
                        string r = words[0];
                        string q = words[1];
                        string name = r.Split("*")[0];
                        string filePath = SheetHelper.testingFile(q, name);
                        Row addedRow = addeddeviationrow[0];
                        Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                        Attachment attachmentDeviation = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet.Id.Value, targetRow.Id.Value, filePath, "application/msword");


                        if (System.IO.File.Exists(filePath))
                        {
                            SheetHelper.DeleteFile(filePath);
                        }
                    }
                }
                catch (Exception ex)
                {
                    return BadRequest(ex.Message);
                }
            }
            return Ok(new
            { Message = "Updated Successfully" });
        }

        [HttpDelete("DeleteFilesFromPostSettlementSheet")]
        public IActionResult DeleteFilesFromPostSettlementSheet(DeleteFilesArray formDataList)
        {
            var eventId = formDataList.EventId;
            Sheet sheet10 = SheetHelper.GetSheetById(smartsheet, sheetId10);

            Row? targetRow = sheet10.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == eventId));
            if (targetRow != null)
            {
                PaginatedResult<Attachment> attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet10.Id.Value, targetRow.Id.Value, null);
                if (attachments.Data != null || attachments.Data.Count > 0)
                {
                    foreach (var attachment in attachments.Data)
                    {
                        long Id = attachment.Id.Value;
                        if (formDataList.AttachmentIds.Contains(Id.ToString()))
                        {
                            smartsheet.SheetResources.AttachmentResources.DeleteAttachment(sheet10.Id.Value, Id);
                        }
                    }
                    return Ok(new
                    { Message = "Deleted Successfully" });
                }
            }
            else
            {
                return Ok(new
                { Message = "Row not found" });
            }
            return Ok();
        }
    }
}












