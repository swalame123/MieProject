using IndiaEvents.Models.Models.Draft;
using IndiaEventsWebApi.Helper;
using IndiaEventsWebApi.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Serilog;
using Smartsheet.Api;
using Smartsheet.Api.Models;
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


        //[HttpGet("GetDataFromAllSheetsUsingEventIdInPreEvent")]
        //public IActionResult GetDataFromAllSheetsUsingEventId(string eventId)
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
                    List<string> columnsToInclude = new List<string> { "EventDate", "Event Topic", "StartTime", "EndTime", "State", "VenueName", "City", "Brands", "Panelists", "SlideKits", "Invitees", "MIPL Invitees", "Expenses", "Meeting Type" };

                    Dictionary<string, object> rowData = new Dictionary<string, object>();
                    for (int i = 0; i < columnNames.Count; i++)
                    {
                        if (columnsToInclude.Contains(columnNames[i]))
                        {
                            rowData[columnNames[i]] = row.Cells[i].Value;
                        }
                    }
                    var attachments = smartsheet.SheetResources.RowResources.AttachmentResources.ListAttachments(sheet1.Id.Value, row.Id.Value, null);

                    Dictionary<string, object> attachmentInfo = new Dictionary<string, object>();
                    foreach (var attachment in attachments.Data)
                    {
                        var AID = (long)attachment.Id;
                        var file = smartsheet.SheetResources.AttachmentResources.GetAttachment(sheet1.Id.Value, AID);

                        var fileId = (long)attachment.Id;
                        //attachmentInfo[file.Name] = file.Url;
                        Dictionary<string, object> attachmentInfoData = new()
                            {
                                { "Name", file.Name },
                                { "Id", file.Id },
                                { "Url", file.Url }
                            };
                        attachmentInfo[file.Name] = attachmentInfoData;



                    }
                    attachmentsList.Add(attachmentInfo);

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

                                { "Url", file.Url }
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
                                { "Url", file.Url }
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
                                { "Url", file.Url }
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
                                { "Url", file.Url }
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
                                DeviationsattachmentInfo[val] = file.Url;
                            }
                        }
                    }
                    DeviationsattachmentsList.Add(DeviationsattachmentInfo);
                }
            }
            resultData["eventDetails"] = eventDetails;
            resultData["Files"] = attachmentsList;
            resultData["Brands"] = BrandseventDetails;
            resultData["Invitees"] = InviteeseventDetails;
            resultData["PanelDetails"] = PaneleventDetails;
            resultData["SlideKitSelection"] = SlideKiteventDetails;
            resultData["ExpenseSelection"] = ExpenseeventDetails;
            resultData["Deviation"] = DeviationsattachmentsList;
            return Ok(resultData);
        }

        [HttpPut("UpdateClassIPreEvent")]
        public IActionResult UpdateClassIPreEvent(UpdateDataForClassI formDataList)
        {
            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
            //Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
            //Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
            //Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
            //Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
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
            if (targetRow == null)
            {
                try
                {
                    Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.EventDetails.EventTopic });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "StartTime"), Value = formDataList.EventDetails.StartTime });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EndTime"), Value = formDataList.EventDetails.EndTime });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "VenueName"), Value = formDataList.EventDetails.VenueName });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "State"), Value = formDataList.EventDetails.State });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "City"), Value = formDataList.EventDetails.City });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Meeting Type"), Value = formDataList.EventDetails.MeetingType });
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
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Budget"), Value = formDataList.EventDetails.TotalBudget });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Local Conveyance"), Value = formDataList.EventDetails.TotalLocalConveyance });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel Amount"), Value = formDataList.EventDetails.TotalTravelAmount });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense"), Value = formDataList.EventDetails.TotalExpense });

                    IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRow });
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

            if (formDataList.BrandSelection.Count > 0)
            {
                Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
                try
                {
                    foreach (var formdata in formDataList.BrandSelection)
                    {
                        Row? BrandstargetRow = sheet2.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == formdata.Id));
                        if (targetRow == null)
                        {

                            Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Brands"), Value = formdata.BrandName });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "% Allocation"), Value = formdata.PercentageAllocation });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Project ID"), Value = formdata.ProjectId });

                            IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet2.Id.Value, new Row[] { updateRow });
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

            if (formDataList.PanelSelection.Count > 0)
            {
                Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
                try
                {
                    foreach (var formdata in formDataList.PanelSelection)
                    {
                        Row? BrandstargetRow = sheet4.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == formdata.Id));
                        if (targetRow == null)
                        {

                            Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "SpeakerCode"), Value = formdata.SpeakerCode });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TrainerCode"), Value = formdata.TrainerCode });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Speciality"), Value = formdata.Speciality });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Tier"), Value = formdata.Tier });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Qualification"), Value = formdata.Qualification });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Country"), Value = formdata.Country });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Rationale"), Value = formdata.Rationale });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "FCPA Date"), Value = formdata.FcpaIssueDate });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PresentationDuration"), Value = formdata.PresentationDuration });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PanelSessionPreparationDuration"), Value = formdata.PanelSessionPreperationDuration });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PanelDiscussionDuration"), Value = formdata.PanelDisscussionDuration });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "QASessionDuration"), Value = formdata.QaSessionDuration });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "BriefingSession"), Value = formdata.BriefingSession });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TotalSessionHours"), Value = formdata.TotalSessionHours });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HcpRole"), Value = formdata.HcpRole });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCPName"), Value = formdata.HcpName });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "MISCode"), Value = formdata.MisCode });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCP Type"), Value = formdata.GOorNGO });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "ExpenseType"), Value = formdata.ExpenseType });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HonorariumRequired"), Value = formdata.HonorariumRequired });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HonorariumAmount"), Value = formdata.HonarariumAmountIncludingTax });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Honorarium Amount Excluding Tax"), Value = formdata.HonarariumAmountExcludingTax });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Travel"), Value = formdata.TravelAmountIncludingTax });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Travel Excluding Tax"), Value = formdata.TravelExcludingTax });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Travel BTC/BTE"), Value = formdata.TravelBtcorBte });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "LocalConveyance"), Value = formdata.LocalConveyanceIncludingTax });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Local Conveyance Excluding Tax"), Value = formdata.LocalConveyanceExcludingTax });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "LC BTC/BTE"), Value = formdata.LcBtcorBte });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Accomodation"), Value = formdata.AccomdationIncludingTax });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Accomodation Excluding Tax"), Value = formdata.AccomdationExcludingTax });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Accomodation BTC/BTE"), Value = formdata.AccomodationBtcorBte });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PAN card name"), Value = formdata.PanCardName });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Bank Account Number"), Value = formdata.BankAccountNumber });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "IFSC Code"), Value = formdata.IFSCCode });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Bank Name"), Value = formdata.BankName });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Currency"), Value = formdata.Currency });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Other Currency"), Value = formdata.OtherCurrencyType });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Beneficiary Name"), Value = formdata.BeneficiaryName });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Pan Number"), Value = formdata.PanNumber });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Global FMV"), Value = formdata.IsGlobalFMVCheck });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Swift Code"), Value = formdata.SwiftCode });


                            IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet4.Id.Value, new Row[] { updateRow });
                            if (formdata.IsFilesUpload == "Yes")
                            {
                                foreach (var p in formdata.Files)
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
                                            sheet4.Id.Value, (long)p.Id, filePath, "application/msword");
                                    }
                                    else
                                    {
                                        Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                           sheet4.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                                    }

                                    if (System.IO.File.Exists(filePath))
                                    {
                                        SheetHelper.DeleteFile(filePath);
                                    }
                                }

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

            if (formDataList.SlideKitSelection.Count > 0)
            {
                Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId2);
                try
                {
                    foreach (var formdata in formDataList.SlideKitSelection)
                    {
                        Row? BrandstargetRow = sheet5.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == formdata.Id));
                        if (targetRow == null)
                        {

                            Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "HCP Name"), Value = formdata.HcpName });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "MIS"), Value = formdata.MisCode });
                            //updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "EventTopic"), Value = formdata.HcpType });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "Slide Kit Type"), Value = formdata.SlideKitType });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "SlideKit Document"), Value = formdata.SlideKitOption });

                            IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet5.Id.Value, new Row[] { updateRow });
                            if (formdata.IsFilesUpload == "Yes")
                            {
                                foreach (var p in formdata.Files)
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
                                            sheet5.Id.Value, (long)p.Id, filePath, "application/msword");
                                    }
                                    else
                                    {
                                        Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                           sheet5.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                                    }

                                    if (System.IO.File.Exists(filePath))
                                    {
                                        SheetHelper.DeleteFile(filePath);
                                    }
                                }

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

            if (formDataList.InviteeSelection.Count > 0)
            {
                Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId2);
                try
                {
                    foreach (var formdata in formDataList.InviteeSelection)
                    {
                        Row? InviteetargetRow = sheet3.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == formdata.Id));
                        if (targetRow == null)
                        {

                            Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "StartTime"), Value = formdata.InviteeFrom });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "HCPName"), Value = formdata.Name });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "MISCode"), Value = formdata.MisCode });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Employee Code"), Value = formdata.EmployeeCode });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "LocalConveyance"), Value = formdata.IsLocalConveyance });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "BTC/BTE"), Value = formdata.LocalConveyanceType });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Speciality"), Value = formdata.Speciality });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "LcAmount"), Value = formdata.LocalConveyanceAmountIncludingTax });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Lc Amount Excluding Tax"), Value = formdata.LocalConveyanceAmountExcludingTax });
                            IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet3.Id.Value, new Row[] { updateRow });

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

            if (formDataList.ExpenseSelection.Count > 0)
            {
                Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
                try
                {
                    foreach (var formdata in formDataList.ExpenseSelection)
                    {
                        Row? BrandstargetRow = sheet6.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == formdata.Id));
                        if (targetRow == null)
                        {

                            Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Expense"), Value = formdata.Expense });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "BTC/BTE"), Value = formdata.ExpenseType });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Amount"), Value = formdata.ExpenseAmountIncludingTax });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Amount Excluding Tax"), Value = formdata.ExpenseAmountExcludingTax });


                            IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet6.Id.Value, new Row[] { updateRow });
                           
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

            return Ok();
        }
    }

}





