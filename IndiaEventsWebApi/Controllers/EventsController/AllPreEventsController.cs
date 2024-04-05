﻿using IndiaEvents.Models.Models.EventTypeSheets;
using IndiaEvents.Models.Models.RequestSheets;
using IndiaEventsWebApi.Helper;
using IndiaEventsWebApi.Models;
using IndiaEventsWebApi.Models.EventTypeSheets;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using NPOI.Util;
using Org.BouncyCastle.Crypto.Tls;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Globalization;
using System.Text;
using static Org.BouncyCastle.Bcpg.Attr.ImageAttrib;

namespace IndiaEventsWebApi.Controllers.EventsController
{
    [Route("api/[controller]")]
    [ApiController]
    public class AllPreEventsController : ControllerBase
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

        public AllPreEventsController(IConfiguration configuration)
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

        [HttpPost("Class1PreEvent"), DisableRequestSizeLimit]
        public IActionResult Class1PreEvent(AllObjModels formDataList)
        {
            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
            Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
            Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
            Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
            Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
            Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);

            StringBuilder addedBrandsData = new();
            StringBuilder addedInviteesData = new();
            StringBuilder addedMEnariniInviteesData = new();
            StringBuilder addedHcpData = new();
            StringBuilder addedSlideKitData = new();
            StringBuilder addedExpences = new();

            int addedSlideKitDataNo = 1;
            int addedHcpDataNo = 1;
            int addedInviteesDataNo = 1;
            int addedInviteesDataNoforMenarini = 1;
            int addedBrandsDataNo = 1;
            int addedExpencesNo = 1;

            int TotalHonorariumAmount = 0;
            int TotalTravelAmount = 0;
            int TotalAccomodateAmount = 0;
            int TotalHCPLcAmount = 0;
            int TotalInviteesLcAmount = 0;
            int TotalExpenseAmount = 0;

            foreach (var formdata in formDataList.EventRequestExpenseSheet)
            {
                string rowData = $"{addedExpencesNo}. {formdata.Expense} | AmountExcludingTax: {formdata.AmountExcludingTax}| Amount: {formdata.Amount} | {formdata.BtcorBte}";
                addedExpences.AppendLine(rowData);
                addedExpencesNo++;
                var amount = SheetHelper.NumCheck(formdata.Amount);
                TotalExpenseAmount = TotalExpenseAmount + amount;
            }
            string Expense = addedExpences.ToString();
            foreach (var formdata in formDataList.EventRequestHCPSlideKits)
            {
                string rowData = $"{addedSlideKitDataNo}. {formdata.HcpName} | {formdata.SlideKitType}";
                addedSlideKitData.AppendLine(rowData);
                addedSlideKitDataNo++;
            }
            string slideKit = addedSlideKitData.ToString();
            foreach (var formdata in formDataList.RequestBrandsList)
            {
                string rowData = $"{addedBrandsDataNo}. {formdata.BrandName} | {formdata.ProjectId} | {formdata.PercentAllocation}";
                addedBrandsData.AppendLine(rowData);
                addedBrandsDataNo++;
            }
            string brand = addedBrandsData.ToString();
            foreach (var formdata in formDataList.EventRequestInvitees)
            {
                if (formdata.InviteedFrom == "Menarini Employees")
                {
                    string row = $"{addedInviteesDataNoforMenarini}. {formdata.InviteeName}";
                    addedMEnariniInviteesData.AppendLine(row);
                    addedInviteesDataNoforMenarini++;
                }
                else
                {
                    string rowData = $"{addedInviteesDataNo}. {formdata.InviteeName}";
                    addedInviteesData.AppendLine(rowData);
                    addedInviteesDataNo++;
                }
                TotalInviteesLcAmount = TotalInviteesLcAmount + SheetHelper.NumCheck(formdata.LcAmount);
            }
            string Invitees = addedInviteesData.ToString();
            string MenariniInvitees = addedMEnariniInviteesData.ToString();
            foreach (var formdata in formDataList.EventRequestHcpRole)
            {
                int HM = SheetHelper.NumCheck(formdata.HonarariumAmount);
                int t = SheetHelper.NumCheck(formdata.Travel) + SheetHelper.NumCheck(formdata.Accomdation);
                string rowData = $"{addedHcpDataNo}. {formdata.HcpRole} |{formdata.HcpName} | Honr.Amt: {HM} |Trav.&Acc.Amt: {t} ";
                addedHcpData.AppendLine(rowData);
                addedHcpDataNo++;
                TotalHonorariumAmount = TotalHonorariumAmount + SheetHelper.NumCheck(formdata.HonarariumAmount);
                TotalTravelAmount = TotalTravelAmount + SheetHelper.NumCheck(formdata.Travel);
                TotalAccomodateAmount = TotalAccomodateAmount + SheetHelper.NumCheck(formdata.Accomdation);
                TotalHCPLcAmount = TotalHCPLcAmount + SheetHelper.NumCheck(formdata.LocalConveyance);
            }
            string HCP = addedHcpData.ToString();
            int c = TotalHCPLcAmount + TotalInviteesLcAmount;
            int total = TotalHonorariumAmount + TotalTravelAmount + TotalAccomodateAmount + TotalHCPLcAmount + TotalInviteesLcAmount + TotalExpenseAmount;
            int s = TotalTravelAmount + TotalAccomodateAmount;
            try
            {
                Row newRow = new()
                {
                    Cells = new List<Cell>()
                };
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.class1.EventTopic });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "StartTime"), Value = formDataList.class1.StartTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EndTime"), Value = formDataList.class1.EndTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "VenueName"), Value = formDataList.class1.VenueName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "City"), Value = formDataList.class1.City });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "State"), Value = formDataList.class1.State });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Panelists"), Value = HCP });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Invitees"), Value = Invitees });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "MIPL Invitees"), Value = MenariniInvitees });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventType"), Value = formDataList.class1.EventType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventDate"), Value = formDataList.class1.EventDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Brands"), Value = brand });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Expenses"), Value = Expense });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "SlideKits"), Value = slideKit });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "IsAdvanceRequired"), Value = formDataList.class1.IsAdvanceRequired });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventOpen30days"), Value = formDataList.class1.EventOpen30days });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventWithin7days"), Value = formDataList.class1.EventWithin7days });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "InitiatorName"), Value = formDataList.class1.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Advance Amount"), Value = SheetHelper.NumCheck(formDataList.class1.AdvanceAmount) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, " Total Expense BTC"), Value = SheetHelper.NumCheck(formDataList.class1.TotalExpenseBTC) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense BTE"), Value = SheetHelper.NumCheck(formDataList.class1.TotalExpenseBTE) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Honorarium Amount"), Value = TotalHonorariumAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel Amount"), Value = TotalTravelAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel & Accommodation Amount"), Value = s });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Accomodation Amount"), Value = TotalAccomodateAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Budget"), Value = total });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Local Conveyance"), Value = c });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense"), Value = TotalExpenseAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Email"), Value = formDataList.class1.Initiator_Email });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "RBM/BM"), Value = formDataList.class1.RBMorBM });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Head"), Value = formDataList.class1.Sales_Head });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Coordinator"), Value = formDataList.class1.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Marketing Head"), Value = formDataList.class1.Marketing_Head });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Treasury"), Value = formDataList.class1.Finance });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Compliance"), Value = formDataList.class1.ComplianceEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Accounts"), Value = formDataList.class1.FinanceAccountsEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Reporting Manager"), Value = formDataList.class1.ReportingManagerEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "1 Up Manager"), Value = formDataList.class1.FirstLevelEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Affairs Head"), Value = formDataList.class1.MedicalAffairsEmail });

                IList<Row> addedRows = smartsheet.SheetResources.RowResources.AddRows(sheet1.Id.Value, new Row[] { newRow });
                long eventIdColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                Cell? eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                string val = eventIdCell.DisplayValue;
                int x = 1;
                foreach (var p in formDataList.class1.Files)
                {
                    string[] words = p.Split(':');
                    string r = words[0];
                    string q = words[1];
                    string name = r.Split(".")[0];
                    string filePath = SheetHelper.testingFile(q, val, name);
                    Row addedRow = addedRows[0];
                    Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                    x++;
                    if (System.IO.File.Exists(filePath))
                    {
                        SheetHelper.DeleteFile(filePath);
                    }
                }
                if (formDataList.class1.EventOpen30days == "Yes" || formDataList.class1.EventWithin7days == "Yes" || formDataList.class1.FB_Expense_Excluding_Tax == "Yes" || formDataList.class1.IsDeviationUpload == "Yes")
                {
                    List<string> DeviationNames = new List<string>();
                    foreach (var p in formDataList.class1.DeviationFiles)
                    {
                        string[] words = p.Split(':');
                        string r = words[0];
                        DeviationNames.Add(r);
                    }
                    foreach (var deviationname in DeviationNames)
                    {
                        string file = deviationname.Split(".")[0];
                        string eventId = val;
                        try
                        {
                            Row newRow7 = new()
                            {
                                Cells = new List<Cell>()
                            };
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.class1.EventTopic });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formDataList.class1.EventType });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formDataList.class1.EventDate });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "StartTime"), Value = formDataList.class1.StartTime });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EndTime"), Value = formDataList.class1.EndTime });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "VenueName"), Value = formDataList.class1.VenueName });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "City"), Value = formDataList.class1.City });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "State"), Value = formDataList.class1.State });

                            if (file == "30DaysDeviationFile")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Outstanding with initiator for more than 45 days" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = formDataList.class1.EventOpen30days });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = SheetHelper.NumCheck(formDataList.class1.EventOpen30dayscount) });
                            }
                            else if (file == "7DaysDeviationFile")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "5 days from the Event Date" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = formDataList.class1.EventWithin7days });

                            }
                            else if (file == "ExpenseExcludingTax")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Food and Beverages expense exceeds 1500" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "PRE-F&B Expense Excluding Tax"), Value = formDataList.class1.FB_Expense_Excluding_Tax });
                            }
                            else if (file.Contains("Travel_Accomodation3LExceededFile"))
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Travel/AccomodationAggregate Limit of 3,00,000 is Exceeded" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Travel/Accomodation 3,00,000 Exceeded Trigger"), Value = "Yes" });//formDataList.class1.FB_Expense_Excluding_Tax });
                            }
                            else if (file.Contains("TrainerHonorarium12LExceededFile"))
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Honorarium Aggregate Limit of 12,00,000 is Exceeded" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Trainer Honorarium 12,00,000 Exceeded Trigger"), Value = "Yes" }); //formDataList.class1.FB_Expense_Excluding_Tax });
                            }
                            else if (file.Contains("HCPHonorarium6LExceededFile"))
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Honorarium Aggregate Limit of 6,00,000 is Exceeded" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HCP Honorarium 6,00,000 Exceeded Trigger"), Value = "Yes" }); // formDataList.class1.FB_Expense_Excluding_Tax });
                            }
                            //else if (file == "Travel_Accomodation3LExceededFile")
                            //{
                            //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Travel/AccomodationAggregate Limit of 3,00,000 is Exceeded" });
                            //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Travel/Accomodation 3,00,000 Exceeded Trigger"), Value = "Yes" });//formDataList.class1.FB_Expense_Excluding_Tax });
                            //}
                            //else if (file == "TrainerHonorarium12LExceededFile")
                            //{
                            //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Honorarium Aggregate Limit of 12,00,000 is Exceeded" });
                            //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Trainer Honorarium 12,00,000 Exceeded Trigger"), Value = "Yes" }); //formDataList.class1.FB_Expense_Excluding_Tax });
                            //}
                            //else if (file == "HCPHonorarium6LExceededFile")
                            //{
                            //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Honorarium Aggregate Limit of 6,00,000 is Exceeded" });
                            //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HCP Honorarium 6,00,000 Exceeded Trigger"), Value = "Yes" }); // formDataList.class1.FB_Expense_Excluding_Tax });
                            //}
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.class1.Sales_Head });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formDataList.class1.FinanceHead });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formDataList.class1.InitiatorName });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.class1.Initiator_Email });

                            IList<Row> addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });

                            int j = 1;
                            foreach (var p in formDataList.class1.DeviationFiles)
                            {
                                string[] words = p.Split(':');
                                string r = words[0];
                                string q = words[1];
                                if (deviationname == r)
                                {
                                    string name = r.Split(".")[0];
                                    string filePath = SheetHelper.testingFile(q, val, name);
                                    Row addedRow = addeddeviationrow[0];
                                    Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                                    Attachment attachmentinmain = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet1.Id.Value, addedRows[0].Id.Value, filePath, "application/msword");
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

                foreach (var formData in formDataList.EventRequestHcpRole)
                {
                    Row newRow1 = new()
                    {
                        Cells = new List<Cell>()
                    };
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HcpRole"), Value = formData.HcpRole });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "MISCode"), Value = formData.MisCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Travel"), Value = SheetHelper.NumCheck(formData.Travel) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TotalSpend"), Value = SheetHelper.NumCheck(formData.FinalAmount) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Accomodation"), Value = SheetHelper.NumCheck(formData.Accomdation) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "LocalConveyance"), Value = SheetHelper.NumCheck(formData.LocalConveyance) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "SpeakerCode"), Value = formData.SpeakerCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TrainerCode"), Value = formData.TrainerCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HonorariumRequired"), Value = formData.HonorariumRequired });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "AgreementAmount"), Value = SheetHelper.NumCheck(formData.AgreementAmount) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Speciality"), Value = formData.Speciality });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Topic"), Value = formDataList.class1.EventTopic });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Type"), Value = formDataList.class1.EventType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Venue name"), Value = formDataList.class1.VenueName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Date Start"), Value = formDataList.class1.EventDate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event End Date"), Value = formDataList.class1.EventEndDate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCPName"), Value = formData.HcpName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PAN card name"), Value = formData.PanCardName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "ExpenseType"), Value = formData.ExpenseType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Bank Account Number"), Value = formData.BankAccountNumber });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Bank Name"), Value = formData.BankName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "IFSC Code"), Value = formData.IFSCCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "FCPA Date"), Value = formData.Fcpadate });
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
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PresentationDuration"), Value = SheetHelper.NumCheck(formData.PresentationDuration) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PanelSessionPreparationDuration"), Value = SheetHelper.NumCheck(formData.PanelSessionPreperationDuration) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PanelDiscussionDuration"), Value = SheetHelper.NumCheck(formData.PanelDisscussionDuration) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "QASessionDuration"), Value = SheetHelper.NumCheck(formData.QASessionDuration) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "BriefingSession"), Value = SheetHelper.NumCheck(formData.BriefingSession) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TotalSessionHours"), Value = SheetHelper.NumCheck(formData.TotalSessionHours) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Rationale"), Value = formData.Rationale });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "EventId/EventRequestId"), Value = val });


                    IList<Row> row = smartsheet.SheetResources.RowResources.AddRows(sheet4.Id.Value, new Row[] { newRow1 });
                    if (formData.IsUpload == "Yes")
                    {
                        foreach (string p in formData.FilesToUpload)
                        {
                            string[] words = p.Split(':');
                            string r = words[0];
                            string q = words[1];
                            string name = r.Split(".")[0];
                            string filePath = SheetHelper.testingFile(q, val, name);
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
                List<Row> newRows2 = new();
                foreach (var formdata in formDataList.RequestBrandsList)
                {
                    Row newRow2 = new()
                    {
                        Cells = new List<Cell>()
                        {
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "% Allocation"), Value = formdata.PercentAllocation },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Brands"), Value = formdata.BrandName },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Project ID"), Value = formdata.ProjectId },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "EventId/EventRequestId"), Value = val }
                        }
                    };

                    newRows2.Add(newRow2);
                }
                smartsheet.SheetResources.RowResources.AddRows(sheet2.Id.Value, newRows2.ToArray());
                List<Row> newRows3 = new();
                foreach (var formdata in formDataList.EventRequestInvitees)
                {
                    Row newRow3 = new()
                    {
                        Cells = new List<Cell>()
                        {
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "HCPName"), Value = formdata.InviteeName },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Designation"), Value = formdata.Designation },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Employee Code"), Value = formdata.EmployeeCode },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "LocalConveyance"), Value = formdata.LocalConveyance },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "BTC/BTE"), Value = formdata.BtcorBte },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "LcAmount"), Value = SheetHelper.NumCheck(formdata.LcAmount) },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Lc Amount Excluding Tax"), Value = formdata.LcAmountExcludingTax },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "EventId/EventRequestId"), Value = val },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Invitee Source"), Value = formdata.InviteedFrom },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "HCP Type"), Value = formdata.HCPType },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Speciality"), Value = formdata.Speciality },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "MISCode"), Value = formdata.MISCode },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Topic"), Value = formDataList.class1.EventTopic },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Type"), Value = formDataList.class1.EventType },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Venue name"), Value = formDataList.class1.VenueName },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Date Start"), Value = formDataList.class1.EventDate },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event End Date"), Value = formDataList.class1.EventDate }

                        }
                    };
                    newRows3.Add(newRow3);
                }
                smartsheet.SheetResources.RowResources.AddRows(sheet3.Id.Value, newRows3.ToArray());
                foreach (var formdata in formDataList.EventRequestHCPSlideKits)
                {
                    Row newRow5 = new()
                    {
                        Cells = new List<Cell>()
                    };

                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "MIS"), Value = formdata.MIS });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "Slide Kit Type"), Value = formdata.SlideKitType });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "SlideKit Document"), Value = formdata.SlideKitDocument });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "EventId/EventRequestId"), Value = val });


                    IList<Row> row = smartsheet.SheetResources.RowResources.AddRows(sheet5.Id.Value, new Row[] { newRow5 });
                    if (formdata.IsUpload == "Yes")
                    {
                        foreach (string p in formdata.FilesToUpload)
                        {
                            string[] words = p.Split(':');
                            string r = words[0];
                            string q = words[1];
                            string name = r.Split(".")[0];
                            string filePath = SheetHelper.testingFile(q, val, name);
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
                List<Row> newRows6 = new();
                foreach (var formdata in formDataList.EventRequestExpenseSheet)
                {
                    Row newRow6 = new()
                    {
                        Cells = new List<Cell>()
                        {
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Expense"), Value = formdata.Expense },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "EventId/EventRequestID"), Value = val },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "AmountExcludingTax?"), Value = formdata.AmountExcludingTax },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount Excluding Tax"), Value = formdata.ExcludingTaxAmount },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount"), Value = SheetHelper.NumCheck(formdata.Amount) },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formdata.BtcorBte },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BudgetAmount"), Value = SheetHelper.NumCheck(formdata.BudgetAmount) },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTCAmount"), Value = SheetHelper.NumCheck(formdata.BtcAmount) },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTEAmount"), Value = SheetHelper.NumCheck(formdata.BteAmount) },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Topic"), Value = formDataList.class1.EventTopic },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Type"), Value = formDataList.class1.EventType },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Venue name"), Value = formDataList.class1.VenueName },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Date Start"), Value = formDataList.class1.EventDate },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event End Date"), Value = formDataList.class1.EventDate }
                        }
                    };
                    newRows6.Add(newRow6);
                }
                smartsheet.SheetResources.RowResources.AddRows(sheet6.Id.Value, newRows6.ToArray());

                Row targetRow = addedRows[0];
                long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role");
                Cell cellToUpdateB = new() { ColumnId = honorariumSubmittedColumnId, Value = formDataList.class1.Role };
                Row updateRow = new() { Id = targetRow.Id, Cells = new Cell[] { cellToUpdateB } };
                Cell? cellToUpdate = targetRow.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                if (cellToUpdate != null) { cellToUpdate.Value = formDataList.class1.Role; }

                smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRow });

                return Ok(new
                { Message = " Success!" });

            }
            catch (Exception ex)
            {
                return BadRequest($"Could not find {ex.Message}");
            }

        }

        [HttpPost("WebinarPreEvent"), DisableRequestSizeLimit]
        public IActionResult WebinarPreEvent(WebinarPayload formDataList)
        {


            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
            Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
            Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
            Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
            Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
            Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);

            StringBuilder addedBrandsData = new();
            StringBuilder addedInviteesData = new();
            StringBuilder addedMEnariniInviteesData = new();
            StringBuilder addedHcpData = new();
            StringBuilder addedSlideKitData = new();
            StringBuilder addedExpences = new();
            int addedSlideKitDataNo = 1;
            int addedHcpDataNo = 1;
            int addedInviteesDataNo = 1;
            int addedInviteesDataNoforMenarini = 1;
            int addedBrandsDataNo = 1;
            int addedExpencesNo = 1;
            int TotalHonorariumAmount = 0;
            int TotalTravelAmount = 0;
            int TotalAccomodateAmount = 0;
            int TotalHCPLcAmount = 0;
            int TotalInviteesLcAmount = 0;
            int TotalExpenseAmount = 0;
            foreach (var formdata in formDataList.EventRequestExpenseSheet)
            {
                string rowData = $"{addedExpencesNo}. {formdata.Expense} | AmountExcludingTax: {formdata.AmountExcludingTax}| Amount: {formdata.Amount} | {formdata.BtcorBte}";
                addedExpences.AppendLine(rowData);
                addedExpencesNo++;
                var amount = SheetHelper.NumCheck(formdata.Amount);
                TotalExpenseAmount = TotalExpenseAmount + amount;
            }
            string Expense = addedExpences.ToString();

            foreach (var formdata in formDataList.EventRequestHCPSlideKits)
            {
                string rowData = $"{addedSlideKitDataNo}. {formdata.HcpName} | {formdata.SlideKitType}";
                addedSlideKitData.AppendLine(rowData);
                addedSlideKitDataNo++;
            }
            string slideKit = addedSlideKitData.ToString();

            foreach (var formdata in formDataList.RequestBrandsList)
            {
                string rowData = $"{addedBrandsDataNo}. {formdata.BrandName} | {formdata.ProjectId} | {formdata.PercentAllocation}";
                addedBrandsData.AppendLine(rowData);
                addedBrandsDataNo++;
            }
            string brand = addedBrandsData.ToString();

            foreach (var formdata in formDataList.EventRequestInvitees)
            {

                if (formdata.InviteedFrom == "Menarini Employees")
                {
                    string row = $"{addedInviteesDataNoforMenarini}. {formdata.InviteeName}";
                    addedMEnariniInviteesData.AppendLine(row);
                    addedInviteesDataNoforMenarini++;
                }
                else
                {
                    string rowData = $"{addedInviteesDataNo}. {formdata.InviteeName}";
                    addedInviteesData.AppendLine(rowData);
                    addedInviteesDataNo++;
                }

                TotalInviteesLcAmount = TotalInviteesLcAmount + SheetHelper.NumCheck(formdata.LcAmount);
            }
            string Invitees = addedInviteesData.ToString();
            string MenariniInvitees = addedMEnariniInviteesData.ToString();

            foreach (var formdata in formDataList.EventRequestHcpRole)
            {
                int HM = SheetHelper.NumCheck(formdata.HonarariumAmount);
                int t = SheetHelper.NumCheck(formdata.Travel) + SheetHelper.NumCheck(formdata.Accomdation);

                string rowData = $"{addedHcpDataNo}. {formdata.HcpRole} |{formdata.HcpName} | Honr.Amt: {HM} |Trav.&Acc.Amt: {t} ";
                addedHcpData.AppendLine(rowData);
                addedHcpDataNo++;
                TotalHonorariumAmount = TotalHonorariumAmount + SheetHelper.NumCheck(formdata.HonarariumAmount);
                TotalTravelAmount = TotalTravelAmount + SheetHelper.NumCheck(formdata.Travel);
                TotalAccomodateAmount = TotalAccomodateAmount + SheetHelper.NumCheck(formdata.Accomdation);
                TotalHCPLcAmount = TotalHCPLcAmount + SheetHelper.NumCheck(formdata.LocalConveyance);
            }
            string HCP = addedHcpData.ToString();

            var c = TotalHCPLcAmount + TotalInviteesLcAmount;


            var total = TotalHonorariumAmount + TotalTravelAmount + TotalAccomodateAmount + TotalHCPLcAmount + TotalInviteesLcAmount + TotalExpenseAmount;

            var s = (TotalTravelAmount + TotalAccomodateAmount);
            try
            {
                Row newRow = new()
                {
                    Cells = new List<Cell>()
                };
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.Webinar.EventTopic });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventType"), Value = formDataList.Webinar.EventType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventDate"), Value = formDataList.Webinar.EventDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "StartTime"), Value = formDataList.Webinar.StartTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EndTime"), Value = formDataList.Webinar.EndTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Meeting Type"), Value = formDataList.Webinar.Meeting_Type });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Brands"), Value = brand });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Expenses"), Value = Expense });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Panelists"), Value = HCP });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Invitees"), Value = Invitees });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "SlideKits"), Value = slideKit });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "IsAdvanceRequired"), Value = formDataList.Webinar.IsAdvanceRequired });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventOpen30days"), Value = formDataList.Webinar.EventOpen30days });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventWithin7days"), Value = formDataList.Webinar.EventWithin7days });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "InitiatorName"), Value = formDataList.Webinar.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Advance Amount"), Value = SheetHelper.NumCheck(formDataList.Webinar.AdvanceAmount) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, " Total Expense BTC"), Value = SheetHelper.NumCheck(formDataList.Webinar.TotalExpenseBTC) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense BTE"), Value = SheetHelper.NumCheck(formDataList.Webinar.TotalExpenseBTE) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Honorarium Amount"), Value = TotalHonorariumAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel Amount"), Value = TotalTravelAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel & Accommodation Amount"), Value = s });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Accomodation Amount"), Value = TotalAccomodateAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Budget"), Value = total });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Local Conveyance"), Value = c });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense"), Value = TotalExpenseAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Email"), Value = formDataList.Webinar.Initiator_Email });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "RBM/BM"), Value = formDataList.Webinar.RBMorBM });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Head"), Value = formDataList.Webinar.Sales_Head });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Coordinator"), Value = formDataList.Webinar.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Marketing Head"), Value = formDataList.Webinar.Marketing_Head });
                //newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance Treasury"), Value = formData.RequestHonorariumList.MarketingHeadEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Compliance"), Value = formDataList.Webinar.ComplianceEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Accounts"), Value = formDataList.Webinar.FinanceAccountsEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Treasury"), Value = formDataList.Webinar.Finance });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Reporting Manager"), Value = formDataList.Webinar.ReportingManagerEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "1 Up Manager"), Value = formDataList.Webinar.FirstLevelEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Affairs Head"), Value = formDataList.Webinar.MedicalAffairsEmail });
                var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheet1.Id.Value, new Row[] { newRow });
                var eventIdColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                var eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                var val = eventIdCell.DisplayValue;
                var x = 1;

                foreach (var p in formDataList.Webinar.Files)
                {
                    string[] words = p.Split(':');
                    var r = words[0];
                    var q = words[1];
                    var name = r.Split(".")[0];
                    var filePath = SheetHelper.testingFile(q, val, name);
                    var addedRow = addedRows[0];

                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                            sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                    x++;
                    if (System.IO.File.Exists(filePath))
                    {
                        SheetHelper.DeleteFile(filePath);
                    }
                }
                if (formDataList.Webinar.EventOpen30days == "Yes" || formDataList.Webinar.EventWithin7days == "Yes" || formDataList.Webinar.FB_Expense_Excluding_Tax == "Yes" || formDataList.Webinar.IsDeviationUpload == "Yes")
                {
                    List<string> DeviationNames = new List<string>();
                    foreach (var p in formDataList.Webinar.DeviationFiles)
                    {
                        string[] words = p.Split(':');
                        var r = words[0];

                        DeviationNames.Add(r);
                    }
                    foreach (var deviationname in DeviationNames)
                    {
                        var file = deviationname.Split(".")[0];
                        var eventId = val;
                        try
                        {
                            var newRow7 = new Row();
                            newRow7.Cells = new List<Cell>();
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.Webinar.EventTopic });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formDataList.Webinar.EventType });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formDataList.Webinar.EventDate });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "StartTime"), Value = formDataList.Webinar.StartTime });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EndTime"), Value = formDataList.Webinar.EndTime });
                            if (file == "30DaysDeviationFile")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = formDataList.Webinar.EventOpen30days });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Outstanding with initiator for more than 45 days" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = SheetHelper.NumCheck(formDataList.Webinar.EventOpen30dayscount) });
                            }
                            else if (file == "7DaysDeviationFile")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = formDataList.Webinar.EventWithin7days });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "5 days from the Event Date" });
                            }
                            else if (file == "ExpenseExcludingTax")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "PRE-F&B Expense Excluding Tax"), Value = formDataList.Webinar.FB_Expense_Excluding_Tax });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Food and Beverages expense exceeds 1500" });
                            }


                            else if (file.Contains("Travel_Accomodation3LExceededFile"))
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Travel/AccomodationAggregate Limit of 3,00,000 is Exceeded" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Travel/Accomodation 3,00,000 Exceeded Trigger"), Value = "Yes" });//formDataList.class1.FB_Expense_Excluding_Tax });
                            }
                            else if (file.Contains("TrainerHonorarium12LExceededFile"))
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Honorarium Aggregate Limit of 12,00,000 is Exceeded" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Trainer Honorarium 12,00,000 Exceeded Trigger"), Value = "Yes" }); //formDataList.class1.FB_Expense_Excluding_Tax });
                            }
                            else if (file.Contains("HCPHonorarium6LExceededFile"))
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Honorarium Aggregate Limit of 6,00,000 is Exceeded" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HCP Honorarium 6,00,000 Exceeded Trigger"), Value = "Yes" }); // formDataList.class1.FB_Expense_Excluding_Tax });
                            }
                            //else if (file == "Travel_Accomodation3LExceededFile")
                            //{
                            //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Travel/AccomodationAggregate Limit of 3,00,000 is Exceeded" });
                            //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Travel/Accomodation 3,00,000 Exceeded Trigger"), Value = "Yes" });//formDataList.class1.FB_Expense_Excluding_Tax });
                            //}
                            //else if (file == "TrainerHonorarium12LExceededFile")
                            //{
                            //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Honorarium Aggregate Limit of 12,00,000 is Exceeded" });
                            //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Trainer Honorarium 12,00,000 Exceeded Trigger"), Value = "Yes" }); //formDataList.class1.FB_Expense_Excluding_Tax });
                            //}
                            //else if (file == "HCPHonorarium6LExceededFile")
                            //{
                            //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Honorarium Aggregate Limit of 6,00,000 is Exceeded" });
                            //    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HCP Honorarium 6,00,000 Exceeded Trigger"), Value = "Yes" }); // formDataList.class1.FB_Expense_Excluding_Tax });
                            //}
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.Webinar.Sales_Head });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formDataList.Webinar.FinanceHead });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formDataList.Webinar.InitiatorName });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.Webinar.Initiator_Email });

                            var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });
                            var j = 1;
                            foreach (var p in formDataList.Webinar.DeviationFiles)
                            {
                                string[] words = p.Split(':');
                                var r = words[0];
                                var q = words[1];
                                if (deviationname == r)
                                {
                                    var name = r.Split(".")[0];

                                    var filePath = SheetHelper.testingFile(q, val, name);

                                    var addedRow = addeddeviationrow[0];

                                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                            sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                                    var attachmentintoMain = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                             sheet1.Id.Value, addedRows[0].Id.Value, filePath, "application/msword");
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


                foreach (var formData in formDataList.EventRequestHcpRole)
                {
                    Row newRow1 = new()
                    {
                        Cells = new List<Cell>()
                    };
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HcpRole"), Value = formData.HcpRole });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "MISCode"), Value = formData.MisCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Travel"), Value = SheetHelper.NumCheck(formData.Travel) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TotalSpend"), Value = SheetHelper.NumCheck(formData.FinalAmount) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Accomodation"), Value = SheetHelper.NumCheck(formData.Accomdation) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "LocalConveyance"), Value = SheetHelper.NumCheck(formData.LocalConveyance) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "SpeakerCode"), Value = formData.SpeakerCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TrainerCode"), Value = formData.TrainerCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HonorariumRequired"), Value = formData.HonorariumRequired });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "AgreementAmount"), Value = SheetHelper.NumCheck(formData.AgreementAmount) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Speciality"), Value = formData.Speciality });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Topic"), Value = formDataList.Webinar.EventTopic });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Type"), Value = formDataList.Webinar.EventType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Date Start"), Value = formDataList.Webinar.EventDate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event End Date"), Value = formDataList.Webinar.EventEndDate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCPName"), Value = formData.HcpName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PAN card name"), Value = formData.PanCardName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "ExpenseType"), Value = formData.ExpenseType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Bank Account Number"), Value = formData.BankAccountNumber });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Bank Name"), Value = formData.BankName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "IFSC Code"), Value = formData.IFSCCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "FCPA Date"), Value = formData.Fcpadate });
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
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PresentationDuration"), Value = SheetHelper.NumCheck(formData.PresentationDuration) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PanelSessionPreparationDuration"), Value = SheetHelper.NumCheck(formData.PanelSessionPreperationDuration) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PanelDiscussionDuration"), Value = SheetHelper.NumCheck(formData.PanelDisscussionDuration) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "QASessionDuration"), Value = SheetHelper.NumCheck(formData.QASessionDuration) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "BriefingSession"), Value = SheetHelper.NumCheck(formData.BriefingSession) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TotalSessionHours"), Value = SheetHelper.NumCheck(formData.TotalSessionHours) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Rationale"), Value = formData.Rationale });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "EventId/EventRequestId"), Value = val });


                    IList<Row> row = smartsheet.SheetResources.RowResources.AddRows(sheet4.Id.Value, new Row[] { newRow1 });
                    if (formData.IsUpload == "Yes")
                    {
                        foreach (string p in formData.FilesToUpload)
                        {
                            string[] words = p.Split(':');
                            string r = words[0];
                            string q = words[1];
                            string name = r.Split(".")[0];
                            string filePath = SheetHelper.testingFile(q, val, name);
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
                List<Row> newRows2 = new();
                foreach (var formdata in formDataList.RequestBrandsList)
                {
                    Row newRow2 = new()
                    {
                        Cells = new List<Cell>()
                        {
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "% Allocation"), Value = formdata.PercentAllocation },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Brands"), Value = formdata.BrandName },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Project ID"), Value = formdata.ProjectId },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "EventId/EventRequestId"), Value = val }
                        }
                    };

                    newRows2.Add(newRow2);
                }
                smartsheet.SheetResources.RowResources.AddRows(sheet2.Id.Value, newRows2.ToArray());
                List<Row> newRows3 = new();
                foreach (var formdata in formDataList.EventRequestInvitees)
                {
                    Row newRow3 = new()
                    {
                        Cells = new List<Cell>()
                        {
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "HCPName"), Value = formdata.InviteeName },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Designation"), Value = formdata.Designation },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Employee Code"), Value = formdata.EmployeeCode },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "LocalConveyance"), Value = formdata.LocalConveyance },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "BTC/BTE"), Value = formdata.BtcorBte },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "LcAmount"), Value = SheetHelper.NumCheck(formdata.LcAmount) },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Lc Amount Excluding Tax"), Value = formdata.LcAmountExcludingTax },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "EventId/EventRequestId"), Value = val },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Invitee Source"), Value = formdata.InviteedFrom },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "HCP Type"), Value = formdata.HCPType },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Speciality"), Value = formdata.Speciality },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "MISCode"), Value = formdata.MISCode },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Topic"), Value = formDataList.Webinar.EventTopic },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Type"), Value = formDataList.Webinar.EventType },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Date Start"), Value = formDataList.Webinar.EventDate },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event End Date"), Value = formDataList.Webinar.EventDate }

                        }
                    };
                    newRows3.Add(newRow3);
                }
                smartsheet.SheetResources.RowResources.AddRows(sheet3.Id.Value, newRows3.ToArray());
                foreach (var formdata in formDataList.EventRequestHCPSlideKits)
                {
                    Row newRow5 = new()
                    {
                        Cells = new List<Cell>()
                    };

                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "MIS"), Value = formdata.MIS });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "Slide Kit Type"), Value = formdata.SlideKitType });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "SlideKit Document"), Value = formdata.SlideKitDocument });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "EventId/EventRequestId"), Value = val });


                    IList<Row> row = smartsheet.SheetResources.RowResources.AddRows(sheet5.Id.Value, new Row[] { newRow5 });
                    if (formdata.IsUpload == "Yes")
                    {
                        foreach (string p in formdata.FilesToUpload)
                        {
                            string[] words = p.Split(':');
                            string r = words[0];
                            string q = words[1];
                            string name = r.Split(".")[0];
                            string filePath = SheetHelper.testingFile(q, val, name);
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
                List<Row> newRows6 = new();
                foreach (var formdata in formDataList.EventRequestExpenseSheet)
                {
                    Row newRow6 = new()
                    {
                        Cells = new List<Cell>()
                        {
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Expense"), Value = formdata.Expense },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "EventId/EventRequestID"), Value = val },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "AmountExcludingTax?"), Value = formdata.AmountExcludingTax },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount Excluding Tax"), Value = formdata.ExcludingTaxAmount },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount"), Value = SheetHelper.NumCheck(formdata.Amount) },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formdata.BtcorBte },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BudgetAmount"), Value = SheetHelper.NumCheck(formdata.BudgetAmount) },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTCAmount"), Value = SheetHelper.NumCheck(formdata.BtcAmount) },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTEAmount"), Value = SheetHelper.NumCheck(formdata.BteAmount) },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Topic"), Value = formDataList.Webinar.EventTopic },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Type"), Value = formDataList.Webinar.EventType },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Date Start"), Value = formDataList.Webinar.EventDate },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event End Date"), Value = formDataList.Webinar.EventDate }
                        }
                    };
                    newRows6.Add(newRow6);
                }
                smartsheet.SheetResources.RowResources.AddRows(sheet6.Id.Value, newRows6.ToArray());



                var addedrow = addedRows[0];
                long ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role");
                var UpdateB = new Cell { ColumnId = ColumnId, Value = formDataList.Webinar.Role };
                Row updateRows = new Row { Id = addedrow.Id, Cells = new Cell[] { UpdateB } };
                var cellsToUpdate = addedrow.Cells.FirstOrDefault(c => c.ColumnId == ColumnId);
                if (cellsToUpdate != null) { cellsToUpdate.Value = formDataList.Webinar.Role; }

                smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRows });


                return Ok(new
                { Message = " Success!" });
            }
            catch (Exception ex)
            {
                return BadRequest($"Could not find {ex.Message}");
            }
        }

        [HttpPost("HCPConsultantPreEvent"), DisableRequestSizeLimit]
        public IActionResult HCPConsultantPreEvent(HCPConsultantPayload formDataList)
        {

            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
            Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
            Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
            Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);

            StringBuilder addedBrandsData = new();
            StringBuilder addedHcpData = new();
            StringBuilder addedExpences = new();

            int addedSlideKitDataNo = 1;
            int addedHcpDataNo = 1;
            int addedInviteesDataNo = 1;
            int addedBrandsDataNo = 1;
            int addedExpencesNo = 1;

            int TotalHonorariumAmount = 0;
            int TotalTravelAmount = 0;
            int TotalAccomodateAmount = 0;
            int TotalHCPLcAmount = 0;
            int TotalInviteesLcAmount = 0;
            int TotalExpenseAmount = 0;
            int TotalRegstAmount = 0;



            string EventOpen30Days = !string.IsNullOrEmpty(formDataList.HcpConsultant.EventOpen30days) ? "Yes" : "No";
            string EventWithin7Days = !string.IsNullOrEmpty(formDataList.HcpConsultant.EventWithin7days) ? "Yes" : "No";
            string BrouchereUpload = !string.IsNullOrEmpty(formDataList.HcpConsultant.BrochureFile) ? "Yes" : "No";
            string FCPA = !string.IsNullOrEmpty(formDataList.HcpConsultant.FcpaFile) ? "Yes" : "No";


            foreach (var formdata in formDataList.ExpenseSheet)
            {
                string rowData = $"{addedExpencesNo}. {formdata.Expense} | RegstAmount: {formdata.RegstAmount}| {formdata.BTC_BTE}";
                addedExpences.AppendLine(rowData);
                addedExpencesNo++;

                int amount = SheetHelper.NumCheck(formdata.ExpenseAmount);
                int regst = SheetHelper.NumCheck(formdata.RegstAmount);
                TotalExpenseAmount = TotalExpenseAmount + amount;
                TotalRegstAmount = TotalRegstAmount + regst;
            }
            string Expense = addedExpences.ToString();

            foreach (var formdata in formDataList.BrandsList)
            {
                string rowData = $"{addedBrandsDataNo}. {formdata.BrandName} | {formdata.ProjectId} | {formdata.PercentAllocation}";
                addedBrandsData.AppendLine(rowData);
                addedBrandsDataNo++;
            }
            string brand = addedBrandsData.ToString();

            foreach (var formdata in formDataList.HcpList)
            {
                int HM = SheetHelper.NumCheck(formdata.RegistrationAmount);
                int t = SheetHelper.NumCheck(formdata.TravelAmount) + SheetHelper.NumCheck(formdata.AccomAmount);
                string rowData = $"{addedHcpDataNo}. {formdata.MisCode} |{formdata.HcpName} | Regst.Amt: {HM} |Trav.&Acc.Amt: {t} ";

                addedHcpData.AppendLine(rowData);
                addedHcpDataNo++;
                TotalHonorariumAmount = TotalHonorariumAmount + SheetHelper.NumCheck(formdata.RegistrationAmount);
                TotalTravelAmount = TotalTravelAmount + SheetHelper.NumCheck(formdata.TravelAmount);
                TotalAccomodateAmount = TotalAccomodateAmount + SheetHelper.NumCheck(formdata.AccomAmount);
                TotalHCPLcAmount = TotalHCPLcAmount + SheetHelper.NumCheck(formdata.LcAmount);
            }
            string HCP = addedHcpData.ToString();
            var c = TotalHCPLcAmount + TotalInviteesLcAmount;
            var total = TotalHonorariumAmount + TotalTravelAmount + TotalAccomodateAmount + TotalHCPLcAmount + TotalInviteesLcAmount;

            var s = (TotalTravelAmount + TotalAccomodateAmount);

            try
            {
                var newRow = new Row();
                newRow.Cells = new List<Cell>();
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.HcpConsultant.EventTopic });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventType"), Value = formDataList.HcpConsultant.EventType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventDate"), Value = formDataList.HcpConsultant.EventDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event End Date"), Value = formDataList.HcpConsultant.EventEndDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "VenueName"), Value = formDataList.HcpConsultant.VenueName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "StartTime"), Value = formDataList.HcpConsultant.StartTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EndTime"), Value = formDataList.HcpConsultant.EndTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sponsorship Society Name"), Value = formDataList.HcpConsultant.SponsorshipSocietyName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Venue Country"), Value = formDataList.HcpConsultant.Country });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "IsAdvanceRequired"), Value = formDataList.HcpConsultant.IsAdvanceRequired });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Advance Amount"), Value = SheetHelper.NumCheck(formDataList.HcpConsultant.AdvanceAmount) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Brands"), Value = brand });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Expenses"), Value = Expense });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Panelists"), Value = HCP });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "InitiatorName"), Value = formDataList.HcpConsultant.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total HCP Registration Amount"), Value = TotalHonorariumAmount });
                //newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total HCP Registration Amount"), Value = TotalExpenseAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel Amount"), Value = TotalTravelAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel & Accommodation Amount"), Value = s });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Accomodation Amount"), Value = TotalAccomodateAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Local Conveyance"), Value = c });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense"), Value = TotalExpenseAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Budget"), Value = total });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Email"), Value = formDataList.HcpConsultant.Initiator_Email });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "RBM/BM"), Value = formDataList.HcpConsultant.RBMorBM });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Head"), Value = formDataList.HcpConsultant.Sales_Head });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Coordinator"), Value = formDataList.HcpConsultant.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Marketing Head"), Value = formDataList.HcpConsultant.Marketing_Head });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Compliance"), Value = formDataList.HcpConsultant.ComplianceEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Accounts"), Value = formDataList.HcpConsultant.FinanceAccountsEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Treasury"), Value = formDataList.HcpConsultant.Finance });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Reporting Manager"), Value = formDataList.HcpConsultant.ReportingManagerEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "1 Up Manager"), Value = formDataList.HcpConsultant.FirstLevelEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Affairs Head"), Value = formDataList.HcpConsultant.MedicalAffairsEmail });
                //newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role"), Value = formDataList.HcpConsultant.Role });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, " Total Expense BTC"), Value = SheetHelper.NumCheck(formDataList.HcpConsultant.TotalExpenseBTC) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense BTE"), Value = SheetHelper.NumCheck(formDataList.HcpConsultant.TotalExpenseBTE) });

                var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheet1.Id.Value, new Row[] { newRow });

                var eventIdColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                var eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                var val = eventIdCell.DisplayValue;

                if (BrouchereUpload == "Yes")
                {


                    var filename = "Brochure";
                    var filePath = SheetHelper.testingFile(formDataList.HcpConsultant.BrochureFile, val, filename);


                    var addedRow = addedRows[0];
                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                            sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");


                    if (System.IO.File.Exists(filePath))
                    {
                        System.IO.File.Delete(filePath);
                    }
                }



                if (EventOpen30Days == "Yes" || EventWithin7Days == "Yes" || formDataList.HcpConsultant.AggregateDeviation == "Yes")
                {
                    List<string> DeviationNames = new List<string>();

                    DeviationNames.Add($"30DaysDeviationFile:{formDataList.HcpConsultant.EventOpen30days}");
                    DeviationNames.Add($"7DaysDeviationFile:{formDataList.HcpConsultant.EventWithin7days}");
                    DeviationNames.Add($"AgregateSpendDeviationFile:{formDataList.HcpConsultant.AggregateDeviationFiles}");
                    var eventId = val;
                    foreach (var name in DeviationNames)
                    {
                        var y = name.Split(':');
                        var fn = y[0];
                        var bs = y[1];

                        if (bs != "")
                        {
                            try
                            {

                                var newRow7 = new Row();
                                newRow7.Cells = new List<Cell>();

                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.HcpConsultant.EventTopic });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formDataList.HcpConsultant.EventType });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formDataList.HcpConsultant.EventDate });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "StartTime"), Value = formDataList.HcpConsultant.StartTime });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EndTime"), Value = formDataList.HcpConsultant.EndTime });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "VenueName"), Value = formDataList.HcpConsultant.VenueName });
                                if (fn == "30DaysDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = EventOpen30Days });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Outstanding with initiator for more than 45 days" });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = SheetHelper.NumCheck(formDataList.HcpConsultant.EventOpen30dayscount) });

                                }
                                else if (fn == "7DaysDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = EventWithin7Days });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "5 days from the Event Date" });
                                }
                                else if (fn == "AgregateSpendDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HCP exceeds 5,00,000 Trigger"), Value = formDataList.HcpConsultant.AggregateDeviation });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Aggregate Limit of 5,00,000 is Exceeded" });
                                }
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.HcpConsultant.Sales_Head });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formDataList.HcpConsultant.FinanceHead });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formDataList.HcpConsultant.InitiatorName });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.HcpConsultant.Initiator_Email });

                                var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });

                                if (fn == "AgregateSpendDeviationFile")
                                {
                                    foreach (var file in formDataList.HcpConsultant.AggregateDeviationFiles)
                                    {
                                        string filename = fn;
                                        string filePath = SheetHelper.testingFile(file, val, filename);
                                        Row addedRow = addeddeviationrow[0];
                                        Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                                sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                                        Attachment attachmentintoMain = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                                sheet1.Id.Value, addedRows[0].Id.Value, filePath, "application/msword");
                                        if (System.IO.File.Exists(filePath))
                                        {
                                            SheetHelper.DeleteFile(filePath);
                                        }
                                    }
                                }
                                else
                                {
                                    string filename = fn;
                                    string filePath = SheetHelper.testingFile(bs, val, filename);

                                    Row addedRow = addeddeviationrow[0];

                                    Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                            sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                                    Attachment attachmentintoMain = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                               sheet1.Id.Value, addedRows[0].Id.Value, filePath, "application/msword");
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
                    }


                }







                foreach (var formData in formDataList.HcpList)
                {
                    Row newRow1 = new()
                    {
                        Cells = new List<Cell>()
                    };
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCPName"), Value = formData.HcpName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "MISCode"), Value = formData.MisCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCP Type"), Value = formData.HcpType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Travel"), Value = SheetHelper.NumCheck(formData.TravelAmount) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Topic"), Value = formDataList.HcpConsultant.EventTopic });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Type"), Value = formDataList.HcpConsultant.EventType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Date Start"), Value = formDataList.HcpConsultant.EventDate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event End Date"), Value = formDataList.HcpConsultant.EventEndDate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Venue name"), Value = formDataList.HcpConsultant.VenueName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Accomodation"), Value = SheetHelper.NumCheck(formData.AccomAmount) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "LocalConveyance"), Value = SheetHelper.NumCheck(formData.LcAmount) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Registration Amount"), Value = SheetHelper.NumCheck(formData.RegistrationAmount) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TotalSpend"), Value = SheetHelper.NumCheck(formData.BudgetAmount) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Legitimate Need"), Value = formData.Legitimate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Objective Criteria"), Value = formData.Objective });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Rationale"), Value = formData.Rationale });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "FCPA Date"), Value = formData.Fcpadate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "ExpenseType"), Value = formData.ExpenseType });



                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "EventId/EventRequestId"), Value = val });







                    IList<Row> addeddatarow = smartsheet.SheetResources.RowResources.AddRows(sheet4.Id.Value, new Row[] { newRow1 });

                    if (formData.IsUpload == "Yes")
                    {
                        foreach (string p in formData.FilesToUpload)
                        {
                            string[] words = p.Split(':');
                            string r = words[0];
                            string q = words[1];
                            string name = r.Split(".")[0];
                            string filePath = SheetHelper.testingFile(q, val, name);
                            Row addedRow = addeddatarow[0];
                            Row websheetrow = addedRows[0];
                            Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet4.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                            Attachment attachmentInWeb = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet1.Id.Value, websheetrow.Id.Value, filePath, "application/msword");

                            if (System.IO.File.Exists(filePath))
                            {
                                SheetHelper.DeleteFile(filePath);
                            }
                        }
                    }



                    //long columnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                    //Cell? Cell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == columnId);
                    //string value = Cell.DisplayValue;
                    //if (FCPA == "Yes")
                    //{
                    //    string filename = " FCPA";
                    //    string filePath = SheetHelper.testingFile(formDataList.HcpConsultant.FcpaFile, value, filename);
                    //    Row addedRow = addeddatarow[0];
                    //    Row websheetrow = addedRows[0];
                    //    Attachment attachmentInWeb = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet1.Id.Value, websheetrow.Id.Value, filePath, "application/msword");
                    //    Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet4.Id.Value, addedRow.Id.Value, filePath, "application/msword");

                    //    if (System.IO.File.Exists(filePath))
                    //    {
                    //        SheetHelper.DeleteFile(filePath);
                    //    }
                    //}
                }



                List<Row> newRows2 = new();

                foreach (var formdata in formDataList.BrandsList)
                {
                    Row newRow2 = new()
                    {
                        Cells = new List<Cell>()
                        {
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "% Allocation"), Value = formdata.PercentAllocation },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Brands"), Value = formdata.BrandName },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Project ID"), Value = formdata.ProjectId },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "EventId/EventRequestId"), Value = val }
                        }
                    };

                    newRows2.Add(newRow2);
                }

                smartsheet.SheetResources.RowResources.AddRows(sheet2.Id.Value, newRows2.ToArray());

                List<Row> newRows6 = new();
                foreach (var formdata in formDataList.ExpenseSheet)
                {
                    Row newRow6 = new()
                    {
                        Cells = new List<Cell>()
                        {

                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Expense"), Value = formdata.Expense },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "EventId/EventRequestID"), Value = val },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formdata.BTC_BTE },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTCAmount"), Value = SheetHelper.NumCheck(formdata.BtcAmount) },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTEAmount"), Value = SheetHelper.NumCheck(formdata.BteAmount) },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Registration Amount"), Value = SheetHelper.NumCheck(formdata.RegstAmount) },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Registration Amount Excluding Tax"), Value = formdata.RegstAmountExcludingTax },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount"), Value = SheetHelper.NumCheck(formdata.ExpenseAmount) },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount Excluding Tax"), Value = formdata.ExpenseAmountExcludingTax },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Topic"), Value = formDataList.HcpConsultant.EventTopic },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Type"), Value = formDataList.HcpConsultant.EventType },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Date Start"), Value = formDataList.HcpConsultant.EventDate },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event End Date"), Value = formDataList.HcpConsultant.EventEndDate },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Venue name"), Value = formDataList.HcpConsultant.VenueName }

                        }
                    };
                    newRows6.Add(newRow6);
                }
                smartsheet.SheetResources.RowResources.AddRows(sheet6.Id.Value, newRows6.ToArray());





                Row addedrow = addedRows[0];
                long ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role");
                Cell UpdateB = new() { ColumnId = ColumnId, Value = formDataList.HcpConsultant.Role };
                Row updateRows = new() { Id = addedrow.Id, Cells = new Cell[] { UpdateB } };
                Cell? cellsToUpdate = addedrow.Cells.FirstOrDefault(c => c.ColumnId == ColumnId);
                if (cellsToUpdate != null) { cellsToUpdate.Value = formDataList.HcpConsultant.Role; }

                smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRows });

                return Ok(new
                { Message = " Success!" });
            }
            catch (Exception ex)
            {
                return BadRequest($"Could not find {ex.Message}");
            }
        }

        [HttpPost("StallFabricationPreEvent"), DisableRequestSizeLimit]
        public IActionResult StallFabricationPreEvent(AllStallFabrication formDataList)
        {
            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
            Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
            Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);

            StringBuilder addedBrandsData = new StringBuilder();
            StringBuilder addedExpences = new StringBuilder();
            int addedBrandsDataNo = 1;
            int addedExpencesNo = 1;
            int TotalExpenseAmount = 0;
            var uploadDeviationForTableContainsData = !string.IsNullOrEmpty(formDataList.StallFabrication.TableContainsDataUpload) ? "Yes" : "No";
            var EventWithin7Days = !string.IsNullOrEmpty(formDataList.StallFabrication.EventWithin7daysUpload) ? "Yes" : "No";
            var BrouchereUpload = !string.IsNullOrEmpty(formDataList.StallFabrication.EventBrouchereUpload) ? "Yes" : "No";
            var InvoiceUpload = !string.IsNullOrEmpty(formDataList.StallFabrication.Invoice_QuotationUpload) ? "Yes" : "No";

            foreach (var formdata in formDataList.ExpenseSheets)
            {
                string rowData = $"{addedExpencesNo}. {formdata.Expense} | AmountExcludingTax: {formdata.AmountExcludingTax}| Amount: {formdata.Amount} | {formdata.BtcorBte}";
                addedExpences.AppendLine(rowData);
                addedExpencesNo++;
                var amount = SheetHelper.NumCheck(formdata.Amount);
                TotalExpenseAmount = TotalExpenseAmount + amount;
            }
            string Expense = addedExpences.ToString();
            foreach (var formdata in formDataList.EventBrands)
            {
                string rowData = $"{addedBrandsDataNo}. {formdata.BrandName} | {formdata.ProjectId} | {formdata.PercentAllocation}";
                addedBrandsData.AppendLine(rowData);
                addedBrandsDataNo++;
            }
            string brand = addedBrandsData.ToString();
            var total = TotalExpenseAmount;

            try
            {
                var newRow = new Row();
                newRow.Cells = new List<Cell>();
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.StallFabrication.EventName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventType"), Value = formDataList.StallFabrication.EventType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventDate"), Value = formDataList.StallFabrication.StartDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event End Date"), Value = formDataList.StallFabrication.EndDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Class III Event Code"), Value = formDataList.StallFabrication.Class_III_EventCode });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Brands"), Value = brand });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Expenses"), Value = Expense });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "InitiatorName"), Value = formDataList.StallFabrication.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense"), Value = TotalExpenseAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Budget"), Value = total });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "IsAdvanceRequired"), Value = formDataList.StallFabrication.IsAdvanceRequired });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Advance Amount"), Value = SheetHelper.NumCheck(formDataList.StallFabrication.AdvanceAmount) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Email"), Value = formDataList.StallFabrication.Initiator_Email });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "RBM/BM"), Value = formDataList.StallFabrication.RBMorBM });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Head"), Value = formDataList.StallFabrication.Sales_Head });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Coordinator"), Value = formDataList.StallFabrication.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Marketing Head"), Value = formDataList.StallFabrication.Marketing_Head });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Compliance"), Value = formDataList.StallFabrication.ComplianceEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Accounts"), Value = formDataList.StallFabrication.FinanceAccountsEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Treasury"), Value = formDataList.StallFabrication.Finance });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Reporting Manager"), Value = formDataList.StallFabrication.ReportingManagerEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "1 Up Manager"), Value = formDataList.StallFabrication.FirstLevelEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Affairs Head"), Value = formDataList.StallFabrication.MedicalAffairsEmail });
                //newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role"), Value = formDataList.StallFabrication.Role });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, " Total Expense BTC"), Value = SheetHelper.NumCheck(formDataList.StallFabrication.TotalExpenseBTC) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense BTE"), Value = SheetHelper.NumCheck(formDataList.StallFabrication.TotalExpenseBTE) });

                var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheet1.Id.Value, new Row[] { newRow });
                var eventIdColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                var eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                var val = eventIdCell.DisplayValue;



                if (BrouchereUpload == "Yes")
                {
                    var name = " Brochure";
                    var filePath = SheetHelper.testingFile(formDataList.StallFabrication.Invoice_QuotationUpload, val, name);

                    var addedRow = addedRows[0];

                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                            sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");


                    if (System.IO.File.Exists(filePath))
                    {
                        SheetHelper.DeleteFile(filePath);
                    }
                }
                if (InvoiceUpload == "Yes")
                {
                    var name = " Invoice_Quotation";

                    var filePath = SheetHelper.testingFile(formDataList.StallFabrication.Invoice_QuotationUpload, val, name);

                    var addedRow = addedRows[0];

                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                            sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");


                    if (System.IO.File.Exists(filePath))
                    {
                        SheetHelper.DeleteFile(filePath);
                    }
                }


                if (formDataList.StallFabrication.IsDeviationUpload == "Yes")
                {
                    var eventId = val;
                    List<string> list = new List<string>
                    {
                        "30days","7days"
                    };
                    foreach (var item in list)
                    {
                        if (uploadDeviationForTableContainsData == "Yes" && item == "30days")
                        {
                            try
                            {
                                var newRow7 = new Row();
                                newRow7.Cells = new List<Cell>();
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.StallFabrication.EventName });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formDataList.StallFabrication.EventType });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formDataList.StallFabrication.StartDate });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event End Date"), Value = formDataList.StallFabrication.EndDate });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = uploadDeviationForTableContainsData });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.StallFabrication.Sales_Head });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formDataList.StallFabrication.InitiatorName });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.StallFabrication.Initiator_Email });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Outstanding with initiator for more than 45 days" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = SheetHelper.NumCheck(formDataList.StallFabrication.EventOpen30dayscount) });

                                var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });

                                var name = "30DaysDeviationFile";
                                var filePath = SheetHelper.testingFile(formDataList.StallFabrication.Invoice_QuotationUpload, val, name);

                                var addedRow = addeddeviationrow[0];

                                var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                        sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                                var attachmentintoMain = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                        sheet1.Id.Value, addedRows[0].Id.Value, filePath, "application/msword");

                                if (System.IO.File.Exists(filePath))
                                {
                                    SheetHelper.DeleteFile(filePath);
                                }

                            }
                            catch (Exception ex)
                            {
                                return BadRequest(ex.Message);
                            }
                        }
                        else if (EventWithin7Days == "Yes" && item == "7days")
                        {
                            try
                            {

                                Row newRow7 = new()
                                {
                                    Cells = new List<Cell>()
                                };
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.StallFabrication.EventName });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formDataList.StallFabrication.EventType });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formDataList.StallFabrication.StartDate });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event End Date"), Value = formDataList.StallFabrication.EndDate });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = EventWithin7Days });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.StallFabrication.Sales_Head });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formDataList.StallFabrication.InitiatorName });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.StallFabrication.Initiator_Email });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "5 days from the Event Date" });


                                var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });

                                var name = "7DaysDeviationFile";
                                var filePath = SheetHelper.testingFile(formDataList.StallFabrication.Invoice_QuotationUpload, val, name);



                                var addedRow = addeddeviationrow[0];
                                var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                        sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                                var attachmentintoMain = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                        sheet1.Id.Value, addedRows[0].Id.Value, filePath, "application/msword");


                                if (System.IO.File.Exists(filePath))
                                {
                                    SheetHelper.DeleteFile(filePath);
                                }

                            }
                            catch (Exception ex)
                            {
                                return BadRequest(ex.Message);
                            }
                        }
                    }

                }



                List<Row> newRows2 = new();

                foreach (var formdata in formDataList.EventBrands)
                {
                    Row newRow2 = new()
                    {
                        Cells = new List<Cell>()
                        {
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "% Allocation"), Value = formdata.PercentAllocation },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Brands"), Value = formdata.BrandName },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Project ID"), Value = formdata.ProjectId },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "EventId/EventRequestId"), Value = val }
                        }
                    };

                    newRows2.Add(newRow2);
                }

                smartsheet.SheetResources.RowResources.AddRows(sheet2.Id.Value, newRows2.ToArray());





                foreach (var formdata in formDataList.ExpenseSheets)
                {
                    Row newRow6 = new()
                    {
                        Cells = new List<Cell>()
                    };

                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Expense"), Value = formdata.Expense });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "EventId/EventRequestID"), Value = val });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "AmountExcludingTax?"), Value = formdata.AmountExcludingTax });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount"), Value = SheetHelper.NumCheck(formdata.Amount) });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount Excluding Tax"), Value = formdata.ExcludingTaxAmount });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formdata.BtcorBte });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTCAmount"), Value = SheetHelper.NumCheck(formdata.BtcAmount) });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTEAmount"), Value = SheetHelper.NumCheck(formdata.BteAmount) });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BudgetAmount"), Value = SheetHelper.NumCheck(formdata.BudgetAmount) });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Topic"), Value = formDataList.StallFabrication.EventName });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Type"), Value = formDataList.StallFabrication.EventType });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Date Start"), Value = formDataList.StallFabrication.StartDate });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event End Date"), Value = formDataList.StallFabrication.EndDate });

                    smartsheet.SheetResources.RowResources.AddRows(sheet6.Id.Value, new Row[] { newRow6 });
                }
                Row addedrow = addedRows[0];
                long ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role");
                Cell UpdateB = new Cell { ColumnId = ColumnId, Value = formDataList.StallFabrication.Role };
                Row updateRows = new Row { Id = addedrow.Id, Cells = new Cell[] { UpdateB } };
                Cell? cellsToUpdate = addedrow.Cells.FirstOrDefault(c => c.ColumnId == ColumnId);
                if (cellsToUpdate != null) { cellsToUpdate.Value = formDataList.StallFabrication.Role; }

                smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRows });

                return Ok(new
                { Message = " Success!" });



            }



            catch (Exception ex)
            {
                return BadRequest($"Could not find {ex.Message}");
            }












        }

        [HttpPost("MedicalUtilityPreEvent"), DisableRequestSizeLimit]
        public IActionResult MedicalUtilityPreEvent(MedicalUtilityPreEventPayload formDataList)
        {


            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
            Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
            Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
            Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);


            StringBuilder addedBrandsData = new StringBuilder();
            StringBuilder addedHcpData = new StringBuilder();
            StringBuilder addedExpences = new StringBuilder();

            int addedHcpDataNo = 1;
            int addedBrandsDataNo = 1;
            int addedExpencesNo = 1;


            var TotalExpenseAmount = 0;
            var EventOpen30Days = !string.IsNullOrEmpty(formDataList.MedicalUtilityData.EventOpen30daysFile) ? "Yes" : "No";
            var EventWithin7Days = !string.IsNullOrEmpty(formDataList.MedicalUtilityData.EventWithin7daysFile) ? "Yes" : "No";
            var UploadDeviationFile = !string.IsNullOrEmpty(formDataList.MedicalUtilityData.UploadDeviationFile) ? "Yes" : "No";
            var FCPA = "";




            foreach (var formdata in formDataList.ExpenseSheet)
            {
                string rowData = $"{addedExpencesNo}. {formdata.Expense} | TotalAmount: {formdata.TotalExpenseAmount}| {formdata.BTC_BTE}";
                addedExpences.AppendLine(rowData);
                addedExpencesNo++;
                var amount = SheetHelper.NumCheck(formdata.TotalExpenseAmount);
                TotalExpenseAmount = TotalExpenseAmount + amount;
            }
            string Expense = addedExpences.ToString();
            foreach (var formdata in formDataList.BrandsList)
            {
                string rowData = $"{addedBrandsDataNo}. {formdata.BrandName} | {formdata.ProjectId} | {formdata.PercentAllocation}";
                addedBrandsData.AppendLine(rowData);
                addedBrandsDataNo++;
            }
            string brand = addedBrandsData.ToString();
            foreach (var formdata in formDataList.HcpList)
            {
                string rowData = $"{addedHcpDataNo}. {formdata.MisCode} |{formdata.HcpName} |Speciality: {formdata.Speciality} |Tier: {formdata.Tier} ";
                addedHcpData.AppendLine(rowData);
                addedHcpDataNo++;
            }
            string HCP = addedHcpData.ToString();

            var total = TotalExpenseAmount;


            try
            {

                var newRow = new Row();
                newRow.Cells = new List<Cell>();
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.MedicalUtilityData.EventTopic });

                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventType"), Value = formDataList.MedicalUtilityData.EventType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventDate"), Value = formDataList.MedicalUtilityData.EventDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Valid From"), Value = formDataList.MedicalUtilityData.ValidFrom });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Valid To"), Value = formDataList.MedicalUtilityData.ValidTill });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Utility Type"), Value = formDataList.MedicalUtilityData.MedicalUtilityType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Utility Description"), Value = formDataList.MedicalUtilityData.MedicalUtilityDescription });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "IsAdvanceRequired"), Value = formDataList.MedicalUtilityData.IsAdvanceRequired });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Advance Amount"), Value = SheetHelper.NumCheck(formDataList.MedicalUtilityData.AdvanceAmount) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Brands"), Value = brand });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Expenses"), Value = Expense });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Panelists"), Value = HCP });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "InitiatorName"), Value = formDataList.MedicalUtilityData.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense"), Value = TotalExpenseAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Budget"), Value = total });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Email"), Value = formDataList.MedicalUtilityData.Initiator_Email });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "RBM/BM"), Value = formDataList.MedicalUtilityData.RBMorBM });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Head"), Value = formDataList.MedicalUtilityData.Sales_Head });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Coordinator"), Value = formDataList.MedicalUtilityData.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Marketing Head"), Value = formDataList.MedicalUtilityData.Marketing_Head });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Compliance"), Value = formDataList.MedicalUtilityData.ComplianceEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Accounts"), Value = formDataList.MedicalUtilityData.FinanceAccountsEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Treasury"), Value = formDataList.MedicalUtilityData.Finance });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Reporting Manager"), Value = formDataList.MedicalUtilityData.ReportingManagerEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "1 Up Manager"), Value = formDataList.MedicalUtilityData.FirstLevelEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Affairs Head"), Value = formDataList.MedicalUtilityData.MedicalAffairsEmail });
                //newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role"), Value = formDataList.MedicalUtilityData.Role });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, " Total Expense BTC"), Value = SheetHelper.NumCheck(formDataList.MedicalUtilityData.TotalExpenseBTC) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense BTE"), Value = SheetHelper.NumCheck(formDataList.MedicalUtilityData.TotalExpenseBTE) });


                var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheet1.Id.Value, new Row[] { newRow });

                var eventIdColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                var eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                var val = eventIdCell.DisplayValue;


                List<string> DeviationNames = new List<string>();
                DeviationNames.Add($"30DaysDeviationFile:{formDataList.MedicalUtilityData.EventOpen30daysFile}");
                DeviationNames.Add($"7DaysDeviationFile:{formDataList.MedicalUtilityData.EventWithin7daysFile}");
                DeviationNames.Add($"AgregateSpendDeviationFile:{formDataList.MedicalUtilityData.UploadDeviationFile}");

                if (EventOpen30Days == "Yes" || EventWithin7Days == "Yes" || UploadDeviationFile == "Yes")
                {
                    var eventId = val;
                    foreach (var name in DeviationNames)
                    {
                        var y = name.Split(':');
                        var fn = y[0];
                        var bs = y[1];

                        if (bs != "")
                        {
                            try
                            {
                                var newRow7 = new Row();
                                newRow7.Cells = new List<Cell>();
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.MedicalUtilityData.EventTopic });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formDataList.MedicalUtilityData.EventType });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formDataList.MedicalUtilityData.EventDate });
                                if (fn == "30DaysDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = EventOpen30Days });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = formDataList.MedicalUtilityData.EventOpen30dayscount });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Outstanding with initiator for more than 45 days" });

                                }
                                else if (fn == "7DaysDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = EventWithin7Days });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "5 days from the Event Date" });

                                }
                                else if (fn == "AgregateSpendDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HCP exceeds 1,00,000 Trigger"), Value = UploadDeviationFile });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Aggregate Limit of 1,00,000 is Exceeded" });
                                }
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.MedicalUtilityData.Sales_Head });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = /*"sybil.oommen@menariniapac.com"*/formDataList.MedicalUtilityData.FinanceHead });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formDataList.MedicalUtilityData.InitiatorName });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.MedicalUtilityData.Initiator_Email });


                                var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });

                                var filename = fn;
                                var filePath = SheetHelper.testingFile(bs, val, filename);



                                var addedRow = addeddeviationrow[0];

                                var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                        sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                                var attachmentintoMain = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                        sheet1.Id.Value, addedRows[0].Id.Value, filePath, "application/msword");

                                if (System.IO.File.Exists(filePath))
                                {
                                    System.IO.File.Delete(filePath);
                                }

                            }
                            catch (Exception ex)
                            {
                                return BadRequest(ex.Message);
                            }
                        }
                    }


                }





                foreach (var formData in formDataList.HcpList)
                {
                    var newRow1 = new Row();
                    newRow1.Cells = new List<Cell>();
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCPName"), Value = formData.HcpName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "MISCode"), Value = formData.MisCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCP Type"), Value = formData.HcpType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Speciality"), Value = formData.Speciality });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Tier"), Value = formData.Tier });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Medical Utility Cost"), Value = SheetHelper.NumCheck(formData.MedicalUtilityCostAmount) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Medical Utility Type"), Value = formDataList.MedicalUtilityData.MedicalUtilityType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Medical Utility description"), Value = formDataList.MedicalUtilityData.MedicalUtilityDescription });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Legitimate Need"), Value = formData.Legitimate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Objective Criteria"), Value = formData.Objective });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Rationale"), Value = formData.Rationale });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Topic"), Value = formDataList.MedicalUtilityData.EventTopic });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Type"), Value = formDataList.MedicalUtilityData.EventType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Date Start"), Value = formDataList.MedicalUtilityData.ValidFrom });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event End Date"), Value = formDataList.MedicalUtilityData.ValidTill });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "EventId/EventRequestId"), Value = val });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "ExpenseType"), Value = formData.ExpenseType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "FCPA Date"), Value = formData.Fcpadate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Request Date"), Value = formData.HCPRequestDate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Valid From"), Value = formDataList.MedicalUtilityData.ValidFrom });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Valid To"), Value = formDataList.MedicalUtilityData.ValidTill });

                    //Request Date,fcpa date,Expense type
                    var addeddatarows = smartsheet.SheetResources.RowResources.AddRows(sheet4.Id.Value, new Row[] { newRow1 });

                    var FCPAFile = !string.IsNullOrEmpty(formData.UploadFCPA) ? "Yes" : "No";
                    var UploadWrittenRequestDate = !string.IsNullOrEmpty(formData.UploadWrittenRequestDate) ? "Yes" : "No";
                    var Invoice_Brouchere_Quotation = !string.IsNullOrEmpty(formData.Invoice_Brouchere_Quotation) ? "Yes" : "No";

                    var columnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                    var Cell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == columnId);
                    var value = Cell.DisplayValue;
                    if (FCPAFile == "Yes")
                    {
                        var filename = " FCPA";
                        var filePath = SheetHelper.testingFile(formData.UploadFCPA, val, filename);
                        var addedRow = addeddatarows[0];
                        var webRow = addedRows[0];
                        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet4.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                        var webattachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet1.Id.Value, webRow.Id.Value, filePath, "application/msword");
                        if (System.IO.File.Exists(filePath))
                        {
                            SheetHelper.DeleteFile(filePath);
                        }
                    }
                    if (UploadWrittenRequestDate == "Yes")
                    {
                        var filename = " UploadWrittenRequestDate";
                        var filePath = SheetHelper.testingFile(formData.UploadWrittenRequestDate, val, filename);
                        var addedRow = addeddatarows[0];
                        var webRow = addedRows[0];
                        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet4.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                        var webattachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet1.Id.Value, webRow.Id.Value, filePath, "application/msword");
                        if (System.IO.File.Exists(filePath))
                        {
                            SheetHelper.DeleteFile(filePath);
                        }
                    }

                    if (Invoice_Brouchere_Quotation == "Yes")
                    {
                        var filename = " Invoice_Brouchere_Quotation";
                        var filePath = SheetHelper.testingFile(formData.Invoice_Brouchere_Quotation, val, filename);
                        var addedRow = addeddatarows[0];
                        var webRow = addedRows[0];
                        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet4.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                        var webattachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet1.Id.Value, webRow.Id.Value, filePath, "application/msword");
                        if (System.IO.File.Exists(filePath))
                        {
                            System.IO.File.Delete(filePath);
                        }
                    }

                }


                List<Row> newRows2 = new();

                foreach (var formdata in formDataList.BrandsList)
                {
                    Row newRow2 = new()
                    {
                        Cells = new List<Cell>()
                        {
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "% Allocation"), Value = formdata.PercentAllocation },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Brands"), Value = formdata.BrandName },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Project ID"), Value = formdata.ProjectId },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "EventId/EventRequestId"), Value = val }
                        }
                    };

                    newRows2.Add(newRow2);
                }

                smartsheet.SheetResources.RowResources.AddRows(sheet2.Id.Value, newRows2.ToArray());



                List<Row> newRows6 = new();
                foreach (var formdata in formDataList.ExpenseSheet)
                {
                    Row newRow6 = new()
                    {
                        Cells = new List<Cell>()
                        {

                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "MisCode"), Value = formdata.MisCode },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Expense"), Value = formdata.Expense },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "EventId/EventRequestID"), Value = val },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formdata.BTC_BTE },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount"), Value = SheetHelper.NumCheck(formdata.TotalExpenseAmount) },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount Excluding Tax"), Value = formdata.TotalExpenseAmountExcludingTax },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTCAmount"), Value = SheetHelper.NumCheck(formdata.BtcAmount) },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTEAmount"), Value = SheetHelper.NumCheck(formdata.BteAmount) },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Topic"), Value = formDataList.MedicalUtilityData.EventTopic },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Type"), Value = formDataList.MedicalUtilityData.EventType },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Date Start"), Value = formDataList.MedicalUtilityData.ValidFrom },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event End Date"), Value = formDataList.MedicalUtilityData.ValidTill },
                        }
                    };
                    newRows6.Add(newRow6);
                }
                smartsheet.SheetResources.RowResources.AddRows(sheet6.Id.Value, newRows6.ToArray());


                var s = addedRows[0];
                long ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role");
                var UpdateB = new Cell { ColumnId = ColumnId, Value = formDataList.MedicalUtilityData.Role };
                Row updateRows = new Row { Id = s.Id, Cells = new Cell[] { UpdateB } };
                var cellsToUpdate = s.Cells.FirstOrDefault(c => c.ColumnId == ColumnId);
                if (cellsToUpdate != null) { cellsToUpdate.Value = formDataList.MedicalUtilityData.Role; }

                smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRows });


                return Ok(new
                { Message = " Success!" });
            }

            catch (Exception ex)
            {
                return BadRequest($"Could not find {ex.Message}");
            }
        }





        #region
        //[HttpPost("ClassIPreEvent"), DisableRequestSizeLimit]
        //public IActionResult ClassIPreEvent(AllObjModels formDataList)
        //{

        //    string sheetId1 = configuration.GetSection("SmartsheetSettings:Class1").Value;
        //    string sheetId2 = configuration.GetSection("SmartsheetSettings:EventRequestBrandsList").Value;
        //    string sheetId3 = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
        //    string sheetId4 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
        //    string sheetId5 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpSlideKit").Value;
        //    string sheetId6 = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
        //    string sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;


        //    Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
        //    Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
        //    Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
        //    Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
        //    Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
        //    Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
        //    Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);

        //    StringBuilder addedBrandsData = new StringBuilder();
        //    StringBuilder addedInviteesData = new StringBuilder();
        //    StringBuilder addedHcpData = new StringBuilder();
        //    StringBuilder addedSlideKitData = new StringBuilder();
        //    StringBuilder addedExpences = new StringBuilder();

        //    int addedSlideKitDataNo = 1;
        //    int addedHcpDataNo = 1;
        //    int addedInviteesDataNo = 1;
        //    int addedBrandsDataNo = 1;
        //    int addedExpencesNo = 1;

        //    var TotalHonorariumAmount = 0;
        //    var TotalTravelAmount = 0;
        //    var TotalAccomodateAmount = 0;
        //    var TotalHCPLcAmount = 0;
        //    var TotalInviteesLcAmount = 0;
        //    var TotalExpenseAmount = 0;

        //    CultureInfo hindi = new CultureInfo("hi-IN");

        //    foreach (var formdata in formDataList.EventRequestExpenseSheet)
        //    {
        //        string rowData = $"{addedExpencesNo}. {formdata.Expense} | AmountExcludingTax: {formdata.AmountExcludingTax}| Amount: {formdata.Amount} | {formdata.BtcorBte}";
        //        addedExpences.AppendLine(rowData);
        //        addedExpencesNo++;
        //        var amount = int.Parse(formdata.Amount);
        //        TotalExpenseAmount = TotalExpenseAmount + amount;
        //    }
        //    string Expense = addedExpences.ToString();
        //    foreach (var formdata in formDataList.EventRequestHCPSlideKits)
        //    {

        //        string rowData = $"{addedSlideKitDataNo}. {formdata.MIS} | {formdata.SlideKitType}";
        //        addedSlideKitData.AppendLine(rowData);
        //        addedSlideKitDataNo++;
        //    }
        //    string slideKit = addedSlideKitData.ToString();

        //    foreach (var formdata in formDataList.RequestBrandsList)
        //    {

        //        string rowData = $"{addedBrandsDataNo}. {formdata.BrandName} | {formdata.ProjectId} | {formdata.PercentAllocation}";
        //        addedBrandsData.AppendLine(rowData);
        //        addedBrandsDataNo++;
        //    }
        //    string brand = addedBrandsData.ToString();

        //    foreach (var formdata in formDataList.EventRequestInvitees)
        //    {
        //        // string rowData = $"{addedInviteesDataNo}. Name: {formdata.InviteeName} | MIS Code: {formdata.MISCode} | LocalConveyance: {formdata.LocalConveyance} ";
        //        string rowData = $"{addedInviteesDataNo}. {formdata.InviteeName}";
        //        addedInviteesData.AppendLine(rowData);
        //        addedInviteesDataNo++;
        //        TotalInviteesLcAmount = TotalInviteesLcAmount + int.Parse(formdata.LcAmount);
        //    }
        //    string Invitees = addedInviteesData.ToString();

        //    foreach (var formdata in formDataList.EventRequestHcpRole)
        //    {
        //        var HM = int.Parse(formdata.HonarariumAmount);
        //        var x = string.Format(hindi, "{0:#,#}", HM);
        //        var t = int.Parse(formdata.Travel) + int.Parse(formdata.Accomdation);
        //        var y = string.Format(hindi, "{0:#,#}", t);
        //        //string rowData = $"{addedHcpDataNo}. {formdata.HcpRole} |Name: {formdata.HcpName} | Honr.Amt: {formdata.HonarariumAmount} |Trav.Amt: {formdata.Travel} |Acco.Amt: {formdata.Accomdation} ";
        //        string rowData = $"{addedHcpDataNo}. {formdata.HcpRole} |{formdata.HcpName} | Honr.Amt: {x} |Trav.&Acc.Amt: {y} ";

        //        addedHcpData.AppendLine(rowData);
        //        addedHcpDataNo++;
        //        TotalHonorariumAmount = TotalHonorariumAmount + int.Parse(formdata.HonarariumAmount);
        //        TotalTravelAmount = TotalTravelAmount + int.Parse(formdata.Travel);
        //        TotalAccomodateAmount = TotalAccomodateAmount + int.Parse(formdata.Accomdation);
        //        TotalHCPLcAmount = TotalHCPLcAmount + int.Parse(formdata.LocalConveyance);
        //    }
        //    string HCP = addedHcpData.ToString();
        //    var FormattedTotalHonorariumAmount = string.Format(hindi, "{0:#,#}", TotalHonorariumAmount);
        //    var FormattedTotalTravelAmount = string.Format(hindi, "{0:#,#}", TotalTravelAmount);
        //    var FormattedTotalAccomodateAmount = string.Format(hindi, "{0:#,#}", TotalAccomodateAmount);
        //    var FormattedTotalHCPLcAmount = string.Format(hindi, "{0:#,#}", TotalHCPLcAmount);
        //    var FornattedTotalInviteesLcAmount = string.Format(hindi, "{0:#,#}", TotalInviteesLcAmount);
        //    var FormattedTotalExpenseAmount = string.Format(hindi, "{0:#,#}", TotalExpenseAmount);
        //    var c = TotalHCPLcAmount + TotalInviteesLcAmount;
        //    var FormattedTotalLC = string.Format(hindi, "{0:#,#}", c);
        //    var total = TotalHonorariumAmount + TotalTravelAmount + TotalAccomodateAmount + TotalHCPLcAmount + TotalInviteesLcAmount + TotalExpenseAmount;

        //    var FormattedTotal = string.Format(hindi, "{0:#,#}", total);
        //    var s = TotalTravelAmount + TotalAccomodateAmount;
        //    var FormattedTotalTAAmount = string.Format(hindi, "{0:#,#}", s);


        //    try
        //    {

        //        var newRow = new Row();
        //        newRow.Cells = new List<Cell>();
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.class1.EventTopic });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "StartTime"), Value = formDataList.class1.StartTime });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EndTime"), Value = formDataList.class1.EndTime });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "VenueName"), Value = formDataList.class1.VenueName });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "City"), Value = formDataList.class1.City });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "State"), Value = formDataList.class1.State });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Panelists"), Value = HCP });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Invitees"), Value = Invitees });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventType"), Value = formDataList.class1.EventType });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventDate"), Value = formDataList.class1.EventDate });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Brands"), Value = brand });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Expenses"), Value = Expense });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "SlideKits"), Value = slideKit });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "IsAdvanceRequired"), Value = formDataList.class1.IsAdvanceRequired });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventOpen30days"), Value = formDataList.class1.EventOpen30days });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventWithin7days"), Value = formDataList.class1.EventWithin7days });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "InitiatorName"), Value = formDataList.class1.InitiatorName });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Advance Amount"), Value = int.Parse(formDataList.class1.AdvanceAmount )});
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, " Total Expense BTC"), Value = int.Parse(formDataList.class1.TotalExpenseBTC )});
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense BTE"), Value = int.Parse(formDataList.class1.TotalExpenseBTE )});
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Honorarium Amount"), Value = TotalHonorariumAmount });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel Amount"), Value = TotalTravelAmount });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel & Accommodation Amount"), Value = s });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Accomodation Amount"), Value = TotalAccomodateAmount });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Budget"), Value = total });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Local Conveyance"), Value = c });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense"), Value = TotalExpenseAmount });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Email"), Value = formDataList.class1.Initiator_Email });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "RBM/BM"), Value = formDataList.class1.RBMorBM });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Head"), Value = formDataList.class1.Sales_Head });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Coordinator"), Value = formDataList.class1.SalesCoordinatorEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Marketing Head"), Value = formDataList.class1.Marketing_Head });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Treasury"), Value = formDataList.class1.Finance });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Compliance"), Value = formDataList.class1.ComplianceEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Accounts"), Value = formDataList.class1.FinanceAccountsEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Reporting Manager"), Value = formDataList.class1.ReportingManagerEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "1 Up Manager"), Value = formDataList.class1.FirstLevelEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Affairs Head"), Value = formDataList.class1.MedicalAffairsEmail });
        //        //newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role"), Value = formDataList.class1.Role });


        //        var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheet1.Id.Value, new Row[] { newRow });

        //        var eventIdColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
        //        var eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
        //        var val = eventIdCell.DisplayValue;

        //        var x = 1;
        //        foreach (var p in formDataList.class1.Files)
        //        {
        //            string[] words = p.Split(':');
        //            var r = words[0];
        //            var q = words[1];
        //            var name = r.Split(".")[0];
        //            var filePath = SheetHelper.testingFile(q, val, name);
        //            var addedRow = addedRows[0];
        //            var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
        //                   sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");
        //            x++;

        //            // ////////////////////////////////////////////////
        //            if (System.IO.File.Exists(filePath))
        //            {
        //                SheetHelper.DeleteFile(filePath);
        //            }
        //        }

        //        //var targetRow = addedRows[0];
        //        //long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role");
        //        //var cellToUpdateB = new Cell { ColumnId = honorariumSubmittedColumnId, Value = formDataList.class1.Role };
        //        //Row updateRow = new Row { Id = targetRow.Id, Cells = new Cell[] { cellToUpdateB } };
        //        //var cellToUpdate = targetRow.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
        //        //if (cellToUpdate != null) { cellToUpdate.Value = formDataList.class1.Role; }

        //        //smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRow });

        //        if (formDataList.class1.EventOpen30days == "Yes" || formDataList.class1.EventWithin7days == "Yes" || formDataList.class1.FB_Expense_Excluding_Tax == "Yes")
        //        {
        //            List<string> DeviationNames = new List<string>();
        //            foreach (var p in formDataList.class1.DeviationFiles)
        //            {
        //                string[] words = p.Split(':');
        //                var r = words[0];

        //                DeviationNames.Add(r);
        //            }
        //            foreach (var deviationname in DeviationNames)
        //            {
        //                var file = deviationname.Split(".")[0];
        //                var eventId = val;
        //                try
        //                {
        //                    var newRow7 = new Row();
        //                    newRow7.Cells = new List<Cell>();

        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.class1.EventTopic });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formDataList.class1.EventType });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formDataList.class1.EventDate });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "StartTime"), Value = formDataList.class1.StartTime });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EndTime"), Value = formDataList.class1.EndTime });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "VenueName"), Value = formDataList.class1.VenueName });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "City"), Value = formDataList.class1.City });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "State"), Value = formDataList.class1.State });


        //                    if (file == "30DaysDeviationFile")
        //                    {
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Outstanding with initiator for more than 45 days" });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = formDataList.class1.EventOpen30days });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = formDataList.class1.EventOpen30dayscount });
        //                    }
        //                    else if (file == "7DaysDeviationFile")
        //                    {
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "5 days from the Event Date" });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = formDataList.class1.EventWithin7days });

        //                    }
        //                    else if (file == "ExpenseExcludingTax")
        //                    {
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Food and Beverages expense exceeds 1500" });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "PRE-F&B Expense Excluding Tax"), Value = formDataList.class1.FB_Expense_Excluding_Tax });
        //                    }
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.class1.Sales_Head });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formDataList.class1.FinanceHead });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formDataList.class1.InitiatorName });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.class1.Initiator_Email });

        //                    var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });

        //                    var j = 1;
        //                    foreach (var p in formDataList.class1.DeviationFiles)
        //                    {
        //                        string[] words = p.Split(':');
        //                        var r = words[0];
        //                        var q = words[1];
        //                        if (deviationname == r)
        //                        {
        //                            var name = r.Split(".")[0];
        //                            var filePath = SheetHelper.testingFile(q, val, name);
        //                            var addedRow = addeddeviationrow[0];
        //                            var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
        //                                    sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");
        //                            j++;
        //                            if (System.IO.File.Exists(filePath))
        //                            {
        //                                SheetHelper.DeleteFile(filePath);
        //                            }
        //                        }
        //                    }
        //                }
        //                catch (Exception ex)
        //                {
        //                    return BadRequest(ex.Message);
        //                }
        //            }
        //        }


        //        foreach (var formData in formDataList.EventRequestHcpRole)
        //        {
        //            var newRow1 = new Row();
        //            newRow1.Cells = new List<Cell>();
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HcpRole"), Value = formData.HcpRole });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "MISCode"), Value = formData.MisCode });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Travel"), Value = formData.Travel });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TotalSpend"), Value = formData.FinalAmount });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Accomodation"), Value = formData.Accomdation });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "LocalConveyance"), Value = formData.LocalConveyance });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "SpeakerCode"), Value = formData.SpeakerCode });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TrainerCode"), Value = formData.TrainerCode });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HonorariumRequired"), Value = formData.HonorariumRequired });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "AgreementAmount"), Value = formData.HonarariumAmount });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Speciality"), Value = formData.Speciality });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Topic"), Value = formDataList.class1.EventTopic });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Type"), Value = formDataList.class1.EventType });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Venue name"), Value = formDataList.class1.VenueName });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Date Start"), Value = formDataList.class1.EventDate });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event End Date"), Value = formDataList.class1.EventEndDate });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCPName"), Value = formData.HcpName });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PAN card name"), Value = formData.PanCardName });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "ExpenseType"), Value = formData.ExpenseType });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Bank Account Number"), Value = formData.BankAccountNumber });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Bank Name"), Value = formData.BankName });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "IFSC Code"), Value = formData.IFSCCode });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "FCPA Date"), Value = formData.Fcpadate });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Currency"), Value = formData.Currency });
        //            if (formData.Currency == "Others")
        //            {
        //                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Other Currency"), Value = formData.OtherCurrencyType });
        //            }
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Beneficiary Name"), Value = formData.BeneficiaryName });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Pan Number"), Value = formData.PanNumber });
        //            if (formData.HcpRole == "Others")
        //            {
        //                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Other Type"), Value = formData.OthersType });
        //            }
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Tier"), Value = formData.Tier });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCP Type"), Value = formData.GOorNGO });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PresentationDuration"), Value = formData.PresentationDuration });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PanelSessionPreparationDuration"), Value = formData.PanelSessionPreperationDuration });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PanelDiscussionDuration"), Value = formData.PanelDisscussionDuration });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "QASessionDuration"), Value = formData.QASessionDuration });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "BriefingSession"), Value = formData.BriefingSession });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TotalSessionHours"), Value = formData.TotalSessionHours });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Rationale"), Value = formData.Rationale });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "EventId/EventRequestId"), Value = val });

        //            smartsheet.SheetResources.RowResources.AddRows(sheet4.Id.Value, new Row[] { newRow1 });
        //        }

        //        foreach (var formdata in formDataList.RequestBrandsList)
        //        {
        //            var newRow2 = new Row();
        //            newRow2.Cells = new List<Cell>();
        //            newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "% Allocation"), Value = formdata.PercentAllocation });
        //            newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Brands"), Value = formdata.BrandName });
        //            newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Project ID"), Value = formdata.ProjectId });
        //            newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "EventId/EventRequestId"), Value = val });
        //            smartsheet.SheetResources.RowResources.AddRows(sheet2.Id.Value, new Row[] { newRow2 });
        //        }

        //        foreach (var formdata in formDataList.EventRequestInvitees)
        //        {
        //            var newRow3 = new Row();
        //            newRow3.Cells = new List<Cell>();
        //            newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Designation"), Value = formdata.Designation });
        //            newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Employee Code"), Value = formdata.EmployeeCode });
        //            newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "HCPName"), Value = formdata.InviteeName });
        //            newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "LocalConveyance"), Value = formdata.LocalConveyance });
        //            newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "BTC/BTE"), Value = formdata.BtcorBte });
        //            newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "LcAmount"), Value = formdata.LcAmount });
        //            newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "EventId/EventRequestId"), Value = val });
        //            newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Invitee Source"), Value = formdata.InviteedFrom });
        //            if (formdata.InviteedFrom == "Others")
        //            {
        //                newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Speciality"), Value = formdata.Speciality });
        //                newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "HCP Type"), Value = formdata.HCPType });
        //            }
        //            else
        //            {
        //                newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "MISCode"), Value = formdata.MISCode });
        //            }
        //            newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Topic"), Value = formDataList.class1.EventTopic });
        //            newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Type"), Value = formDataList.class1.EventType });
        //            newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Venue name"), Value = formDataList.class1.VenueName });
        //            newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Date Start"), Value = formDataList.class1.EventDate });
        //            newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event End Date"), Value = formDataList.class1.EventDate });

        //            smartsheet.SheetResources.RowResources.AddRows(sheet3.Id.Value, new Row[] { newRow3 });
        //        }

        //        foreach (var formdata in formDataList.EventRequestHCPSlideKits)
        //        {
        //            var newRow5 = new Row();
        //            newRow5.Cells = new List<Cell>();
        //            newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "MIS"), Value = formdata.MIS });
        //            newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "Slide Kit Type"), Value = formdata.SlideKitType });
        //            newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "SlideKit Document"), Value = formdata.SlideKitDocument });
        //            newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "EventId/EventRequestId"), Value = val });

        //            smartsheet.SheetResources.RowResources.AddRows(sheet5.Id.Value, new Row[] { newRow5 });
        //        }

        //        foreach (var formdata in formDataList.EventRequestExpenseSheet)
        //        {
        //            var newRow6 = new Row();
        //            newRow6.Cells = new List<Cell>();

        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Expense"), Value = formdata.Expense });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "EventId/EventRequestID"), Value = val });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "AmountExcludingTax?"), Value = formdata.AmountExcludingTax });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount"), Value = formdata.Amount });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formdata.BtcorBte });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BudgetAmount"), Value = formdata.BudgetAmount });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTCAmount"), Value = formdata.BtcAmount });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTEAmount"), Value = formdata.BteAmount });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Topic"), Value = formDataList.class1.EventTopic });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Type"), Value = formDataList.class1.EventType });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Venue name"), Value = formDataList.class1.VenueName });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Date Start"), Value = formDataList.class1.EventDate });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event End Date"), Value = formDataList.class1.EventDate });

        //            smartsheet.SheetResources.RowResources.AddRows(sheet6.Id.Value, new Row[] { newRow6 });
        //        }



        //        var targetRow = addedRows[0];
        //        long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role");
        //        var cellToUpdateB = new Cell { ColumnId = honorariumSubmittedColumnId, Value = formDataList.class1.Role };
        //        Row updateRow = new Row { Id = targetRow.Id, Cells = new Cell[] { cellToUpdateB } };
        //        var cellToUpdate = targetRow.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
        //        if (cellToUpdate != null) { cellToUpdate.Value = formDataList.class1.Role; }

        //        smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRow });

        //        return Ok(new
        //        { Message = " Success!" });

        //    }

        //    catch (Exception ex)
        //    {
        //        return BadRequest($"Could not find {ex.Message}");
        //    }
        //}


        //[HttpPost("WebinarPreEvent"), DisableRequestSizeLimit]
        //public IActionResult WebinarPreEvent(WebinarPayload formDataList)
        //{


        //    string sheetId1 = configuration.GetSection("SmartsheetSettings:Class1").Value;
        //    string sheetId2 = configuration.GetSection("SmartsheetSettings:EventRequestBrandsList").Value;
        //    string sheetId3 = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
        //    string sheetId4 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
        //    string sheetId5 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpSlideKit").Value;
        //    string sheetId6 = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
        //    string sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;


        //    Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
        //    Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
        //    Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
        //    Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
        //    Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
        //    Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
        //    Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);

        //    StringBuilder addedBrandsData = new StringBuilder();
        //    StringBuilder addedInviteesData = new StringBuilder();
        //    StringBuilder addedHcpData = new StringBuilder();
        //    StringBuilder addedSlideKitData = new StringBuilder();
        //    StringBuilder addedExpences = new StringBuilder();
        //    int addedSlideKitDataNo = 1;
        //    int addedHcpDataNo = 1;
        //    int addedInviteesDataNo = 1;
        //    int addedBrandsDataNo = 1;
        //    int addedExpencesNo = 1;
        //    var TotalHonorariumAmount = 0;
        //    var TotalTravelAmount = 0;
        //    var TotalAccomodateAmount = 0;
        //    var TotalHCPLcAmount = 0;
        //    var TotalInviteesLcAmount = 0;
        //    var TotalExpenseAmount = 0;
        //    CultureInfo hindi = new CultureInfo("hi-IN");
        //    foreach (var formdata in formDataList.EventRequestExpenseSheet)
        //    {
        //        string rowData = $"{addedExpencesNo}. {formdata.Expense} | AmountExcludingTax: {formdata.AmountExcludingTax}| Amount: {formdata.Amount} | {formdata.BtcorBte}";
        //        addedExpences.AppendLine(rowData);
        //        addedExpencesNo++;
        //        var amount = int.Parse(formdata.Amount);
        //        TotalExpenseAmount = TotalExpenseAmount + amount;
        //    }
        //    string Expense = addedExpences.ToString();

        //    foreach (var formdata in formDataList.EventRequestHCPSlideKits)
        //    {
        //        string rowData = $"{addedSlideKitDataNo}. {formdata.MIS} | {formdata.SlideKitType}";
        //        addedSlideKitData.AppendLine(rowData);
        //        addedSlideKitDataNo++;
        //    }
        //    string slideKit = addedSlideKitData.ToString();

        //    foreach (var formdata in formDataList.RequestBrandsList)
        //    {
        //        string rowData = $"{addedBrandsDataNo}. {formdata.BrandName} | {formdata.ProjectId} | {formdata.PercentAllocation}";
        //        addedBrandsData.AppendLine(rowData);
        //        addedBrandsDataNo++;
        //    }
        //    string brand = addedBrandsData.ToString();

        //    foreach (var formdata in formDataList.EventRequestInvitees)
        //    {
        //        string rowData = $"{addedInviteesDataNo}. {formdata.InviteeName}";
        //        addedInviteesData.AppendLine(rowData);
        //        addedInviteesDataNo++;
        //        TotalInviteesLcAmount = TotalInviteesLcAmount + int.Parse(formdata.LcAmount);
        //    }
        //    string Invitees = addedInviteesData.ToString();


        //    foreach (var formdata in formDataList.EventRequestHcpRole)
        //    {
        //        var HM = int.Parse(formdata.HonarariumAmount);
        //        var x = string.Format(hindi, "{0:#,#}", HM);
        //        var t = int.Parse(formdata.Travel) + int.Parse(formdata.Accomdation);
        //        var y = string.Format(hindi, "{0:#,#}", t);
        //        //string rowData = $"{addedHcpDataNo}. {formdata.HcpRole} |Name: {formdata.HcpName} | Honr.Amt: {formdata.HonarariumAmount} |Trav.Amt: {formdata.Travel} |Acco.Amt: {formdata.Accomdation} ";
        //        string rowData = $"{addedHcpDataNo}. {formdata.HcpRole} |{formdata.HcpName} | Honr.Amt: {x} |Trav.&Acc.Amt: {y} ";
        //        addedHcpData.AppendLine(rowData);
        //        addedHcpDataNo++;
        //        TotalHonorariumAmount = TotalHonorariumAmount + int.Parse(formdata.HonarariumAmount);
        //        TotalTravelAmount = TotalTravelAmount + int.Parse(formdata.Travel);
        //        TotalAccomodateAmount = TotalAccomodateAmount + int.Parse(formdata.Accomdation);
        //        TotalHCPLcAmount = TotalHCPLcAmount + int.Parse(formdata.LocalConveyance);
        //    }
        //    string HCP = addedHcpData.ToString();
        //    var FormattedTotalHonorariumAmount = string.Format(hindi, "{0:#,#}", TotalHonorariumAmount);
        //    var FormattedTotalTravelAmount = string.Format(hindi, "{0:#,#}", TotalTravelAmount);
        //    var FormattedTotalAccomodateAmount = string.Format(hindi, "{0:#,#}", TotalAccomodateAmount);
        //    var FormattedTotalHCPLcAmount = string.Format(hindi, "{0:#,#}", TotalHCPLcAmount);
        //    var FornattedTotalInviteesLcAmount = string.Format(hindi, "{0:#,#}", TotalInviteesLcAmount);
        //    var FormattedTotalExpenseAmount = string.Format(hindi, "{0:#,#}", TotalExpenseAmount);
        //    var c = TotalHCPLcAmount + TotalInviteesLcAmount;
        //    var FormattedTotalLC = string.Format(hindi, "{0:#,#}", c);
        //    var total = TotalHonorariumAmount + TotalTravelAmount + TotalAccomodateAmount + TotalHCPLcAmount + TotalInviteesLcAmount + TotalExpenseAmount;
        //    var FormattedTotal = string.Format(hindi, "{0:#,#}", total);
        //    var s = TotalTravelAmount + TotalAccomodateAmount;
        //    var FormattedTotalTAAmount = string.Format(hindi, "{0:#,#}", s);
        //    try
        //    {

        //        var newRow = new Row();
        //        newRow.Cells = new List<Cell>();
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.Webinar.EventTopic });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventType"), Value = formDataList.Webinar.EventType });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventDate"), Value = formDataList.Webinar.EventDate });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "StartTime"), Value = formDataList.Webinar.StartTime });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EndTime"), Value = formDataList.Webinar.EndTime });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Meeting Type"), Value = formDataList.Webinar.Meeting_Type });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Brands"), Value = brand });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Expenses"), Value = Expense });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Panelists"), Value = HCP });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Invitees"), Value = Invitees });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "SlideKits"), Value = slideKit });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "IsAdvanceRequired"), Value = formDataList.Webinar.IsAdvanceRequired });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventOpen30days"), Value = formDataList.Webinar.EventOpen30days });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventWithin7days"), Value = formDataList.Webinar.EventWithin7days });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "InitiatorName"), Value = formDataList.Webinar.InitiatorName });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Advance Amount"), Value = int.Parse(formDataList.Webinar.AdvanceAmount )});
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, " Total Expense BTC"), Value = int.Parse(formDataList.Webinar.TotalExpenseBTC )});
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense BTE"), Value = int.Parse(formDataList.Webinar.TotalExpenseBTE )});
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Honorarium Amount"), Value = TotalHonorariumAmount });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel Amount"), Value = TotalTravelAmount });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel & Accommodation Amount"), Value = s });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Accomodation Amount"), Value = TotalAccomodateAmount });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Budget"), Value = total });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Local Conveyance"), Value = c });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense"), Value = TotalExpenseAmount });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Email"), Value = formDataList.Webinar.Initiator_Email });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "RBM/BM"), Value = formDataList.Webinar.RBMorBM });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Head"), Value = formDataList.Webinar.Sales_Head });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Coordinator"), Value = formDataList.Webinar.SalesCoordinatorEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Marketing Head"), Value = formDataList.Webinar.Marketing_Head });
        //        //newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance Treasury"), Value = formData.RequestHonorariumList.MarketingHeadEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Compliance"), Value = formDataList.Webinar.ComplianceEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Accounts"), Value = formDataList.Webinar.FinanceAccountsEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Treasury"), Value = formDataList.Webinar.Finance });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Reporting Manager"), Value = formDataList.Webinar.ReportingManagerEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "1 Up Manager"), Value = formDataList.Webinar.FirstLevelEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Affairs Head"), Value = formDataList.Webinar.MedicalAffairsEmail });
        //        //newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role"), Value = formDataList.Webinar.Role });
        //        var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheet1.Id.Value, new Row[] { newRow });
        //        var eventIdColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
        //        var eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
        //        var val = eventIdCell.DisplayValue;
        //        var x = 1;
        //        foreach (var p in formDataList.Webinar.Files)
        //        {
        //            string[] words = p.Split(':');
        //            var r = words[0];
        //            var q = words[1];
        //            var name = r.Split(".")[0];
        //            var filePath = SheetHelper.testingFile(q, val, name);
        //            var addedRow = addedRows[0];

        //            var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
        //                    sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");
        //            x++;
        //            if (System.IO.File.Exists(filePath))
        //            {
        //                SheetHelper.DeleteFile(filePath);
        //            }
        //        }


        //        if (formDataList.Webinar.EventOpen30days == "Yes" || formDataList.Webinar.EventWithin7days == "Yes" || formDataList.Webinar.FB_Expense_Excluding_Tax == "Yes")
        //        {
        //            List<string> DeviationNames = new List<string>();
        //            foreach (var p in formDataList.Webinar.DeviationFiles)
        //            {
        //                string[] words = p.Split(':');
        //                var r = words[0];

        //                DeviationNames.Add(r);
        //            }
        //            foreach (var deviationname in DeviationNames)
        //            {
        //                var file = deviationname.Split(".")[0];
        //                var eventId = val;
        //                try
        //                {
        //                    var newRow7 = new Row();
        //                    newRow7.Cells = new List<Cell>();
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.Webinar.EventTopic });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formDataList.Webinar.EventType });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formDataList.Webinar.EventDate });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "StartTime"), Value = formDataList.Webinar.StartTime });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EndTime"), Value = formDataList.Webinar.EndTime });
        //                    if (file == "30DaysDeviationFile")
        //                    {
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = formDataList.Webinar.EventOpen30days });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Outstanding with initiator for more than 45 days" });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = formDataList.Webinar.EventOpen30dayscount });
        //                    }
        //                    else if (file == "7DaysDeviationFile")
        //                    {
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = formDataList.Webinar.EventWithin7days });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "5 days from the Event Date" });
        //                    }
        //                    else if (file == "ExpenseExcludingTax")
        //                    {
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "PRE-F&B Expense Excluding Tax"), Value = formDataList.Webinar.FB_Expense_Excluding_Tax });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Food and Beverages expense exceeds 1500" });
        //                    }
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.Webinar.Sales_Head });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formDataList.Webinar.FinanceHead });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formDataList.Webinar.InitiatorName });
        //                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.Webinar.Initiator_Email });

        //                    var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });
        //                    var j = 1;
        //                    foreach (var p in formDataList.Webinar.DeviationFiles)
        //                    {
        //                        string[] words = p.Split(':');
        //                        var r = words[0];
        //                        var q = words[1];
        //                        if (deviationname == r)
        //                        {
        //                            var name = r.Split(".")[0];

        //                            var filePath = SheetHelper.testingFile(q, val, name);

        //                            var addedRow = addeddeviationrow[0];

        //                            var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
        //                                    sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");
        //                            j++;
        //                            if (System.IO.File.Exists(filePath))
        //                            {
        //                                SheetHelper.DeleteFile(filePath);
        //                            }
        //                        }
        //                    }
        //                }
        //                catch (Exception ex)
        //                {
        //                    return BadRequest(ex.Message);
        //                }
        //            }

        //        }



        //        foreach (var formData in formDataList.EventRequestHcpRole)
        //        {
        //            var newRow1 = new Row();
        //            newRow1.Cells = new List<Cell>();
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HcpRole"), Value = formData.HcpRole });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "MISCode"), Value = formData.MisCode });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Travel"), Value = formData.Travel });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TotalSpend"), Value = formData.FinalAmount });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Accomodation"), Value = formData.Accomdation });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "LocalConveyance"), Value = formData.LocalConveyance });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "SpeakerCode"), Value = formData.SpeakerCode });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TrainerCode"), Value = formData.TrainerCode });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HonorariumRequired"), Value = formData.HonorariumRequired });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "AgreementAmount"), Value = formData.HonarariumAmount });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Speciality"), Value = formData.Speciality });

        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Topic"), Value = formDataList.Webinar.EventTopic });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Type"), Value = formDataList.Webinar.EventType });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Date Start"), Value = formDataList.Webinar.EventDate });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event End Date"), Value = formDataList.Webinar.EventEndDate });

        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCPName"), Value = formData.HcpName });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PAN card name"), Value = formData.PanCardName });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "ExpenseType"), Value = formData.ExpenseType });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Bank Account Number"), Value = formData.BankAccountNumber });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Bank Name"), Value = formData.BankName });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "IFSC Code"), Value = formData.IFSCCode });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "FCPA Date"), Value = formData.Fcpadate });

        //            if (formData.HcpRole == "Others")
        //            {

        //                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Other Type"), Value = formData.OthersType });
        //            }

        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Tier"), Value = formData.Tier });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCP Type"), Value = formData.GOorNGO });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PresentationDuration"), Value = formData.PresentationDuration });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PanelSessionPreparationDuration"), Value = formData.PanelSessionPreperationDuration });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PanelDiscussionDuration"), Value = formData.PanelDisscussionDuration });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "QASessionDuration"), Value = formData.QASessionDuration });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "BriefingSession"), Value = formData.BriefingSession });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TotalSessionHours"), Value = formData.TotalSessionHours });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Rationale"), Value = formData.Rationale });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "EventId/EventRequestId"), Value = val });

        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Currency"), Value = formData.Currency });
        //            if (formData.Currency == "Others")
        //            {
        //                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Other Currency"), Value = formData.OtherCurrencyType });
        //            }
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Beneficiary Name"), Value = formData.BeneficiaryName });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Pan Number"), Value = formData.PanNumber });


        //            smartsheet.SheetResources.RowResources.AddRows(sheet4.Id.Value, new Row[] { newRow1 });







        //        }







        //        foreach (var formdata in formDataList.RequestBrandsList)
        //        {
        //            var newRow2 = new Row();
        //            newRow2.Cells = new List<Cell>();
        //            newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "% Allocation"), Value = formdata.PercentAllocation });
        //            newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Brands"), Value = formdata.BrandName });
        //            newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Project ID"), Value = formdata.ProjectId });
        //            newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "EventId/EventRequestId"), Value = val });


        //            smartsheet.SheetResources.RowResources.AddRows(sheet2.Id.Value, new Row[] { newRow2 });

        //        }
        //        foreach (var formdata in formDataList.EventRequestInvitees)
        //        {
        //            var newRow3 = new Row();
        //            newRow3.Cells = new List<Cell>();
        //            newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Designation"), Value = formdata.Designation });
        //            newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Employee Code"), Value = formdata.EmployeeCode });
        //            newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "HCPName"), Value = formdata.InviteeName });
        //            newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "LocalConveyance"), Value = formdata.LocalConveyance });
        //            newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "BTC/BTE"), Value = formdata.BtcorBte });
        //            newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "LcAmount"), Value = formdata.LcAmount });
        //            newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "EventId/EventRequestId"), Value = val });
        //            newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Invitee Source"), Value = formdata.InviteedFrom });
        //            newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Topic"), Value = formDataList.Webinar.EventTopic });
        //            newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Type"), Value = formDataList.Webinar.EventType });
        //            newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Date Start"), Value = formDataList.Webinar.EventDate });
        //            newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event End Date"), Value = formDataList.Webinar.EventDate });
        //            // newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Start Time"), Value = formDataList.Webinar.StartTime });

        //            if (formdata.InviteedFrom == "Others")
        //            {
        //                newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Speciality"), Value = formdata.Speciality });
        //                newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "HCP Type"), Value = formdata.HCPType });
        //            }
        //            else
        //            {
        //                newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "MISCode"), Value = formdata.MISCode });
        //            }
        //            smartsheet.SheetResources.RowResources.AddRows(sheet3.Id.Value, new Row[] { newRow3 });
        //        }
        //        foreach (var formdata in formDataList.EventRequestHCPSlideKits)
        //        {
        //            var newRow5 = new Row();
        //            newRow5.Cells = new List<Cell>();

        //            newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "MIS"), Value = formdata.MIS });
        //            newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "Slide Kit Type"), Value = formdata.SlideKitType });
        //            newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "SlideKit Document"), Value = formdata.SlideKitDocument });
        //            newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "EventId/EventRequestId"), Value = val });


        //            smartsheet.SheetResources.RowResources.AddRows(sheet5.Id.Value, new Row[] { newRow5 });
        //        }
        //        foreach (var formdata in formDataList.EventRequestExpenseSheet)
        //        {
        //            var newRow6 = new Row();
        //            newRow6.Cells = new List<Cell>();

        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Expense"), Value = formdata.Expense });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "EventId/EventRequestID"), Value = val });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "AmountExcludingTax?"), Value = formdata.AmountExcludingTax });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount"), Value = formdata.Amount });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formdata.BtcorBte });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BudgetAmount"), Value = formdata.BudgetAmount });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTCAmount"), Value = formdata.BtcAmount });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTEAmount"), Value = formdata.BteAmount });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Topic"), Value = formDataList.Webinar.EventTopic });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Type"), Value = formDataList.Webinar.EventType });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Date Start"), Value = formDataList.Webinar.EventDate });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event End Date"), Value = formDataList.Webinar.EventDate });

        //            smartsheet.SheetResources.RowResources.AddRows(sheet6.Id.Value, new Row[] { newRow6 });
        //        }


        //        var targetRow = addedRows[0];
        //        long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role");
        //        var cellToUpdateB = new Cell { ColumnId = honorariumSubmittedColumnId, Value = formDataList.Webinar.Role };
        //        Row updateRow = new Row { Id = targetRow.Id, Cells = new Cell[] { cellToUpdateB } };
        //        var cellToUpdate = targetRow.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
        //        if (cellToUpdate != null) { cellToUpdate.Value = formDataList.Webinar.Role; }

        //        smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRow });




        //        return Ok(new
        //        { Message = " Success!" });
        //    }
        //    catch (Exception ex)
        //    {
        //        return BadRequest($"Could not find {ex.Message}");
        //    }
        //}


        //[HttpPost("StallFabricationPreEvent"), DisableRequestSizeLimit]
        //public IActionResult StallFabricationPreEvent(AllStallFabrication formDataList)
        //{
        //    string sheetId1 = configuration.GetSection("SmartsheetSettings:Class1").Value;
        //    string sheetId2 = configuration.GetSection("SmartsheetSettings:EventRequestBrandsList").Value;
        //    // string sheetId3 = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
        //    // sheetId4 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
        //    // sheetId5 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpSlideKit").Value;
        //    string sheetId6 = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
        //    string sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;


        //    Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
        //    Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
        //    // Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
        //    // Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
        //    // Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
        //    Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
        //    Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);
        //    StringBuilder addedBrandsData = new StringBuilder();
        //    StringBuilder addedExpences = new StringBuilder();
        //    int addedBrandsDataNo = 1;
        //    int addedExpencesNo = 1;
        //    var TotalExpenseAmount = 0;
        //    CultureInfo hindi = new CultureInfo("hi-IN");

        //    var uploadDeviationForTableContainsData = !string.IsNullOrEmpty(formDataList.StallFabrication.TableContainsDataUpload) ? "Yes" : "No";
        //    var EventWithin7Days = !string.IsNullOrEmpty(formDataList.StallFabrication.EventWithin7daysUpload) ? "Yes" : "No";
        //    var BrouchereUpload = !string.IsNullOrEmpty(formDataList.StallFabrication.EventBrouchereUpload) ? "Yes" : "No";
        //    var InvoiceUpload = !string.IsNullOrEmpty(formDataList.StallFabrication.Invoice_QuotationUpload) ? "Yes" : "No";


        //    foreach (var formdata in formDataList.ExpenseSheets)
        //    {
        //        string rowData = $"{addedExpencesNo}. {formdata.Expense} | AmountExcludingTax: {formdata.AmountExcludingTax}| Amount: {formdata.Amount} | {formdata.BtcorBte}";
        //        addedExpences.AppendLine(rowData);
        //        addedExpencesNo++;
        //        var amount = int.Parse(formdata.Amount);
        //        TotalExpenseAmount = TotalExpenseAmount + amount;

        //    }
        //    string Expense = addedExpences.ToString();


        //    foreach (var formdata in formDataList.EventBrands)
        //    {

        //        string rowData = $"{addedBrandsDataNo}. {formdata.BrandName} | {formdata.ProjectId} | {formdata.PercentAllocation}";
        //        addedBrandsData.AppendLine(rowData);
        //        addedBrandsDataNo++;
        //    }
        //    string brand = addedBrandsData.ToString();

        //    var FormattedTotalExpenseAmount = string.Format(hindi, "{0:#,#}", TotalExpenseAmount);

        //    var total = TotalExpenseAmount;

        //    var FormattedTotal = string.Format(hindi, "{0:#,#}", total);


        //    try
        //    {

        //        var newRow = new Row();
        //        newRow.Cells = new List<Cell>();
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.StallFabrication.EventName });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventType"), Value = formDataList.StallFabrication.EventType });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventDate"), Value = formDataList.StallFabrication.StartDate });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event End Date"), Value = formDataList.StallFabrication.EndDate });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Class III Event Code"), Value = formDataList.StallFabrication.Class_III_EventCode });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Brands"), Value = brand });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Expenses"), Value = Expense });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "InitiatorName"), Value = formDataList.StallFabrication.InitiatorName });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense"), Value = TotalExpenseAmount });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Budget"), Value = total });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "IsAdvanceRequired"), Value = formDataList.StallFabrication.IsAdvanceRequired });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Advance Amount"), Value = int.Parse(formDataList.StallFabrication.AdvanceAmount )});

        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Email"), Value = formDataList.StallFabrication.Initiator_Email });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "RBM/BM"), Value = formDataList.StallFabrication.RBMorBM });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Head"), Value = formDataList.StallFabrication.Sales_Head });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Coordinator"), Value = formDataList.StallFabrication.SalesCoordinatorEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Marketing Head"), Value = formDataList.StallFabrication.Marketing_Head });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Compliance"), Value = formDataList.StallFabrication.ComplianceEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Accounts"), Value = formDataList.StallFabrication.FinanceAccountsEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Treasury"), Value = formDataList.StallFabrication.Finance });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Reporting Manager"), Value = formDataList.StallFabrication.ReportingManagerEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "1 Up Manager"), Value = formDataList.StallFabrication.FirstLevelEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Affairs Head"), Value = formDataList.StallFabrication.MedicalAffairsEmail });
        //        //newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role"), Value = formDataList.StallFabrication.Role });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, " Total Expense BTC"), Value = int.Parse(formDataList.StallFabrication.TotalExpenseBTC )});
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense BTE"), Value = int.Parse(formDataList.StallFabrication.TotalExpenseBTE )});

        //        var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheet1.Id.Value, new Row[] { newRow });
        //        var eventIdColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
        //        var eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
        //        var val = eventIdCell.DisplayValue;



        //        if (BrouchereUpload == "Yes")
        //        {
        //            var name = " Brochure";
        //            var filePath = SheetHelper.testingFile(formDataList.StallFabrication.Invoice_QuotationUpload, val, name);

        //            var addedRow = addedRows[0];

        //            var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
        //                    sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");


        //            if (System.IO.File.Exists(filePath))
        //            {
        //                SheetHelper.DeleteFile(filePath);
        //            }
        //        }
        //        if (InvoiceUpload == "Yes")
        //        {
        //            var name = " Invoice_Quotation";

        //            var filePath = SheetHelper.testingFile(formDataList.StallFabrication.Invoice_QuotationUpload, val, name);

        //            var addedRow = addedRows[0];

        //            var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
        //                    sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");


        //            if (System.IO.File.Exists(filePath))
        //            {
        //                SheetHelper.DeleteFile(filePath);
        //            }
        //        }


        //        if (formDataList.StallFabrication.IsDeviationUpload == "Yes")
        //        {
        //            var eventId = val;
        //            List<string> list = new List<string>
        //            {
        //                "30days","7days"
        //            };
        //            foreach (var item in list)
        //            {
        //                if (uploadDeviationForTableContainsData == "Yes" && item == "30days")
        //                {
        //                    try
        //                    {
        //                        var newRow7 = new Row();
        //                        newRow7.Cells = new List<Cell>();
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.StallFabrication.EventName });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formDataList.StallFabrication.EventType });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formDataList.StallFabrication.StartDate });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event End Date"), Value = formDataList.StallFabrication.EndDate });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = uploadDeviationForTableContainsData });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.StallFabrication.Sales_Head });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formDataList.StallFabrication.InitiatorName });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.StallFabrication.Initiator_Email });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Outstanding with initiator for more than 45 days" });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = formDataList.StallFabrication.EventOpen30dayscount });

        //                        var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });

        //                        var name = "30DaysDeviationFile";
        //                        var filePath = SheetHelper.testingFile(formDataList.StallFabrication.Invoice_QuotationUpload, val, name);

        //                        var addedRow = addeddeviationrow[0];

        //                        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
        //                                sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");

        //                        if (System.IO.File.Exists(filePath))
        //                        {
        //                            SheetHelper.DeleteFile(filePath);
        //                        }

        //                    }
        //                    catch (Exception ex)
        //                    {
        //                        return BadRequest(ex.Message);
        //                    }
        //                }
        //                else if (EventWithin7Days == "Yes" && item == "7days")
        //                {
        //                    try
        //                    {

        //                        var newRow7 = new Row();
        //                        newRow7.Cells = new List<Cell>();
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.StallFabrication.EventName });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formDataList.StallFabrication.EventType });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formDataList.StallFabrication.StartDate });
        //                        //newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "StartTime"), Value = formDataList.StallFabrication.StartTime });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event End Date"), Value = formDataList.StallFabrication.EndDate });
        //                        //newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen30days"), Value = uploadDeviationForTableContainsData });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = EventWithin7Days });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.StallFabrication.Sales_Head });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formDataList.StallFabrication.InitiatorName });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.StallFabrication.Initiator_Email });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "5 days from the Event Date" });


        //                        var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });

        //                        var name = "7DaysDeviationFile";
        //                        var filePath = SheetHelper.testingFile(formDataList.StallFabrication.Invoice_QuotationUpload, val, name);



        //                        var addedRow = addeddeviationrow[0];
        //                        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
        //                                sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");


        //                        if (System.IO.File.Exists(filePath))
        //                        {
        //                            SheetHelper.DeleteFile(filePath);
        //                        }

        //                    }
        //                    catch (Exception ex)
        //                    {
        //                        return BadRequest(ex.Message);
        //                    }
        //                }
        //            }

        //        }


        //        foreach (var formdata in formDataList.EventBrands)
        //        {
        //            var newRow2 = new Row();
        //            newRow2.Cells = new List<Cell>();
        //            newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "% Allocation"), Value = formdata.PercentAllocation });
        //            newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Brands"), Value = formdata.BrandName });
        //            newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Project ID"), Value = formdata.ProjectId });
        //            newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "EventId/EventRequestId"), Value = val });

        //            smartsheet.SheetResources.RowResources.AddRows(sheet2.Id.Value, new Row[] { newRow2 });

        //        }


        //        foreach (var formdata in formDataList.ExpenseSheets)
        //        {
        //            var newRow6 = new Row();
        //            newRow6.Cells = new List<Cell>();

        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Expense"), Value = formdata.Expense });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "EventId/EventRequestID"), Value = val });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "AmountExcludingTax?"), Value = formdata.AmountExcludingTax });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount"), Value = formdata.Amount });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formdata.BtcorBte });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTCAmount"), Value = formdata.BtcAmount });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTEAmount"), Value = formdata.BteAmount });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BudgetAmount"), Value = formdata.BudgetAmount });
        //            //newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTCAmount"), Value = formdata.BtcAmount });
        //            //newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTEAmount"), Value = formdata.BteAmount });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Topic"), Value = formDataList.StallFabrication.EventName });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Type"), Value = formDataList.StallFabrication.EventType });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Date Start"), Value = formDataList.StallFabrication.StartDate });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event End Date"), Value = formDataList.StallFabrication.EndDate });

        //            smartsheet.SheetResources.RowResources.AddRows(sheet6.Id.Value, new Row[] { newRow6 });
        //        }


        //        var targetRow = addedRows[0];
        //        long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role");
        //        var cellToUpdateB = new Cell { ColumnId = honorariumSubmittedColumnId, Value = formDataList.StallFabrication.Role };
        //        Row updateRow = new Row { Id = targetRow.Id, Cells = new Cell[] { cellToUpdateB } };
        //        var cellToUpdate = targetRow.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
        //        if (cellToUpdate != null) { cellToUpdate.Value = formDataList.StallFabrication.Role; }

        //        smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRow });



        //        return Ok(new
        //        { Message = " Success!" });



        //    }



        //    catch (Exception ex)
        //    {
        //        return BadRequest($"Could not find {ex.Message}");
        //    }












        //}


        //[HttpPost("HCPConsultantPreEvent"), DisableRequestSizeLimit]
        //public IActionResult HCPConsultantPreEvent(HCPConsultantPayload formDataList)
        //{

        //    string sheetId1 = configuration.GetSection("SmartsheetSettings:Class1").Value;
        //    string sheetId2 = configuration.GetSection("SmartsheetSettings:EventRequestBrandsList").Value;
        //    //string sheetId3 = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
        //    string sheetId4 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
        //    //string sheetId5 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpSlideKit").Value;
        //    string sheetId6 = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
        //    string sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;


        //    Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
        //    Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
        //    //Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
        //    Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
        //    //Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
        //    Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
        //    Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);


        //    //Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId1, null, null, null, null, null, null, null);
        //    //Sheet sheet2 = smartsheet.SheetResources.GetSheet(parsedSheetId2, null, null, null, null, null, null, null);
        //    //Sheet sheet4 = smartsheet.SheetResources.GetSheet(parsedSheetId4, null, null, null, null, null, null, null);
        //    //Sheet sheet6 = smartsheet.SheetResources.GetSheet(parsedSheetId6, null, null, null, null, null, null, null);
        //    //Sheet sheet7 = smartsheet.SheetResources.GetSheet(parsedSheetId7, null, null, null, null, null, null, null);

        //    StringBuilder addedBrandsData = new StringBuilder();
        //    StringBuilder addedHcpData = new StringBuilder();
        //    StringBuilder addedExpences = new StringBuilder();

        //    int addedSlideKitDataNo = 1;
        //    int addedHcpDataNo = 1;
        //    int addedInviteesDataNo = 1;
        //    int addedBrandsDataNo = 1;
        //    int addedExpencesNo = 1;

        //    var TotalHonorariumAmount = 0;
        //    var TotalTravelAmount = 0;
        //    var TotalAccomodateAmount = 0;
        //    var TotalHCPLcAmount = 0;
        //    var TotalInviteesLcAmount = 0;
        //    var TotalExpenseAmount = 0;
        //    var TotalRegstAmount = 0;


        //    CultureInfo hindi = new CultureInfo("hi-IN");


        //    var EventOpen30Days = !string.IsNullOrEmpty(formDataList.HcpConsultant.EventOpen30days) ? "Yes" : "No";
        //    var EventWithin7Days = !string.IsNullOrEmpty(formDataList.HcpConsultant.EventWithin7days) ? "Yes" : "No";
        //    var BrouchereUpload = !string.IsNullOrEmpty(formDataList.HcpConsultant.BrochureFile) ? "Yes" : "No";
        //    var FCPA = !string.IsNullOrEmpty(formDataList.HcpConsultant.FcpaFile) ? "Yes" : "No";


        //    foreach (var formdata in formDataList.ExpenseSheet)
        //    {
        //        string rowData = $"{addedExpencesNo}. {formdata.Expense} | RegstAmount: {formdata.RegstAmount}| {formdata.BTC_BTE}";
        //        addedExpences.AppendLine(rowData);
        //        addedExpencesNo++;

        //        var amount = int.Parse(formdata.ExpenseAmount);
        //        var regst = int.Parse(formdata.RegstAmount);
        //        TotalExpenseAmount = TotalExpenseAmount + amount;
        //        TotalRegstAmount = TotalRegstAmount + regst;
        //    }
        //    string Expense = addedExpences.ToString();

        //    foreach (var formdata in formDataList.BrandsList)
        //    {
        //        string rowData = $"{addedBrandsDataNo}. {formdata.BrandName} | {formdata.ProjectId} | {formdata.PercentAllocation}";
        //        addedBrandsData.AppendLine(rowData);
        //        addedBrandsDataNo++;
        //    }
        //    string brand = addedBrandsData.ToString();

        //    foreach (var formdata in formDataList.HcpList)
        //    {
        //        var HM = int.Parse(formdata.RegistrationAmount);
        //        var x = string.Format(hindi, "{0:#,#}", HM);
        //        var t = int.Parse(formdata.TravelAmount) + int.Parse(formdata.AccomAmount);
        //        var y = string.Format(hindi, "{0:#,#}", t);
        //        string rowData = $"{addedHcpDataNo}. {formdata.MisCode} |{formdata.HcpName} | Regst.Amt: {x} |Trav.&Acc.Amt: {y} ";



        //        addedHcpData.AppendLine(rowData);
        //        addedHcpDataNo++;
        //        TotalHonorariumAmount = TotalHonorariumAmount + int.Parse(formdata.RegistrationAmount);
        //        TotalTravelAmount = TotalTravelAmount + int.Parse(formdata.TravelAmount);
        //        TotalAccomodateAmount = TotalAccomodateAmount + int.Parse(formdata.AccomAmount);
        //        TotalHCPLcAmount = TotalHCPLcAmount + int.Parse(formdata.LcAmount);
        //    }
        //    string HCP = addedHcpData.ToString();
        //    var c = TotalHCPLcAmount + TotalInviteesLcAmount;
        //    var total = TotalHonorariumAmount + TotalTravelAmount + TotalAccomodateAmount + TotalHCPLcAmount + TotalInviteesLcAmount;
        //    //var BTE = int.Parse(formDataList.HcpConsultant.TotalExpenseBTE);
        //    //var BTC = int.Parse(formDataList.HcpConsultant.TotalExpenseBTC);

        //    //var total = BTC + BTE;

        //    var s = TotalTravelAmount + TotalAccomodateAmount;

        //    try
        //    {
        //        var newRow = new Row();
        //        newRow.Cells = new List<Cell>();
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.HcpConsultant.EventTopic });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventType"), Value = formDataList.HcpConsultant.EventType });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventDate"), Value = formDataList.HcpConsultant.EventDate });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event End Date"), Value = formDataList.HcpConsultant.EventEndDate });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "VenueName"), Value = formDataList.HcpConsultant.VenueName });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "StartTime"), Value = formDataList.HcpConsultant.StartTime });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EndTime"), Value = formDataList.HcpConsultant.EndTime });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sponsorship Society Name"), Value = formDataList.HcpConsultant.SponsorshipSocietyName });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Venue Country"), Value = formDataList.HcpConsultant.Country });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "IsAdvanceRequired"), Value = formDataList.HcpConsultant.IsAdvanceRequired });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Advance Amount"), Value = int.Parse(formDataList.HcpConsultant.AdvanceAmount )});
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Brands"), Value = brand });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Expenses"), Value = Expense });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Panelists"), Value = HCP });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "InitiatorName"), Value = formDataList.HcpConsultant.InitiatorName });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total HCP Registration Amount"), Value = TotalHonorariumAmount });
        //        //newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total HCP Registration Amount"), Value = TotalExpenseAmount });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel Amount"), Value = TotalTravelAmount });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel & Accommodation Amount"), Value = s });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Accomodation Amount"), Value = TotalAccomodateAmount });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Local Conveyance"), Value = c });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense"), Value = TotalExpenseAmount });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Budget"), Value = total });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Email"), Value = formDataList.HcpConsultant.Initiator_Email });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "RBM/BM"), Value = formDataList.HcpConsultant.RBMorBM });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Head"), Value = formDataList.HcpConsultant.Sales_Head });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Coordinator"), Value = formDataList.HcpConsultant.SalesCoordinatorEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Marketing Head"), Value = formDataList.HcpConsultant.Marketing_Head });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Compliance"), Value = formDataList.HcpConsultant.ComplianceEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Accounts"), Value = formDataList.HcpConsultant.FinanceAccountsEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Treasury"), Value = formDataList.HcpConsultant.Finance });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Reporting Manager"), Value = formDataList.HcpConsultant.ReportingManagerEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "1 Up Manager"), Value = formDataList.HcpConsultant.FirstLevelEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Affairs Head"), Value = formDataList.HcpConsultant.MedicalAffairsEmail });
        //        //newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role"), Value = formDataList.HcpConsultant.Role });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, " Total Expense BTC"), Value = int.Parse(formDataList.HcpConsultant.TotalExpenseBTC )});
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense BTE"), Value = int.Parse(formDataList.HcpConsultant.TotalExpenseBTE )});

        //        var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheet1.Id.Value, new Row[] { newRow });

        //        var eventIdColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
        //        var eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
        //        var val = eventIdCell.DisplayValue;

        //        if (BrouchereUpload == "Yes")
        //        {


        //            var filename = "Brochure";
        //            var filePath = SheetHelper.testingFile(formDataList.HcpConsultant.BrochureFile, val, filename);


        //            var addedRow = addedRows[0];
        //            var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
        //                    sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");


        //            if (System.IO.File.Exists(filePath))
        //            {
        //                System.IO.File.Delete(filePath);
        //            }
        //        }



        //        if (EventOpen30Days == "Yes" || EventWithin7Days == "Yes" || formDataList.HcpConsultant.AggregateDeviation == "Yes")
        //        {
        //            List<string> DeviationNames = new List<string>();

        //            DeviationNames.Add($"30DaysDeviationFile:{formDataList.HcpConsultant.EventOpen30days}");
        //            DeviationNames.Add($"7DaysDeviationFile:{formDataList.HcpConsultant.EventWithin7days}");
        //            DeviationNames.Add($"AgregateSpendDeviationFile:{formDataList.HcpConsultant.AggregateDeviationFiles}");
        //            var eventId = val;
        //            foreach (var name in DeviationNames)
        //            {
        //                var y = name.Split(':');
        //                var fn = y[0];
        //                var bs = y[1];

        //                if (bs != "")
        //                {
        //                    try
        //                    {

        //                        var newRow7 = new Row();
        //                        newRow7.Cells = new List<Cell>();

        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.HcpConsultant.EventTopic });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formDataList.HcpConsultant.EventType });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formDataList.HcpConsultant.EventDate });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "StartTime"), Value = formDataList.HcpConsultant.StartTime });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EndTime"), Value = formDataList.HcpConsultant.EndTime });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "VenueName"), Value = formDataList.HcpConsultant.VenueName });
        //                        if (fn == "30DaysDeviationFile")
        //                        {
        //                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = EventOpen30Days });
        //                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Outstanding with initiator for more than 45 days" });
        //                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = formDataList.HcpConsultant.EventOpen30dayscount });

        //                        }
        //                        else if (fn == "7DaysDeviationFile")
        //                        {
        //                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = EventWithin7Days });
        //                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "5 days from the Event Date" });
        //                        }
        //                        else if (fn == "AgregateSpendDeviationFile")
        //                        {
        //                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HCP exceeds 5,00,000 Trigger"), Value = formDataList.HcpConsultant.AggregateDeviation });
        //                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Aggregate Limit of 5,00,000 is Exceeded" });
        //                        }
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.HcpConsultant.Sales_Head });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formDataList.HcpConsultant.FinanceHead });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formDataList.HcpConsultant.InitiatorName });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.HcpConsultant.Initiator_Email });

        //                        var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });

        //                        if (fn == "AgregateSpendDeviationFile")
        //                        {
        //                            foreach (var file in formDataList.HcpConsultant.AggregateDeviationFiles)
        //                            {
        //                                var filename = fn;
        //                                var filePath = SheetHelper.testingFile(file, val, filename);
        //                                var addedRow = addeddeviationrow[0];
        //                                var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
        //                                        sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");
        //                                if (System.IO.File.Exists(filePath))
        //                                {
        //                                    SheetHelper.DeleteFile(filePath);
        //                                }
        //                            }
        //                        }
        //                        else
        //                        {
        //                            var filename = fn;
        //                            var filePath = SheetHelper.testingFile(bs, val, filename);

        //                            var addedRow = addeddeviationrow[0];

        //                            var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
        //                                    sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");
        //                            if (System.IO.File.Exists(filePath))
        //                            {
        //                                SheetHelper.DeleteFile(filePath);
        //                            }
        //                        }

        //                    }
        //                    catch (Exception ex)
        //                    {
        //                        return BadRequest(ex.Message);
        //                    }
        //                }
        //            }


        //        }







        //        foreach (var formData in formDataList.HcpList)
        //        {
        //            var newRow1 = new Row();
        //            newRow1.Cells = new List<Cell>();
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCPName"), Value = formData.HcpName });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "MISCode"), Value = formData.MisCode });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCP Type"), Value = formData.HcpType });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Travel"), Value = formData.TravelAmount });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Topic"), Value = formDataList.HcpConsultant.EventTopic });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Type"), Value = formDataList.HcpConsultant.EventType });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Date Start"), Value = formDataList.HcpConsultant.EventDate });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event End Date"), Value = formDataList.HcpConsultant.EventEndDate });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Venue name"), Value = formDataList.HcpConsultant.VenueName });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Accomodation"), Value = formData.AccomAmount });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "LocalConveyance"), Value = formData.LcAmount });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Registration Amount"), Value = formData.RegistrationAmount });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TotalSpend"), Value = formData.BudgetAmount });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Legitimate Need"), Value = formData.Legitimate });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Objective Criteria"), Value = formData.Objective });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Rationale"), Value = formData.Rationale });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "FCPA Date"), Value = formData.Fcpadate });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "ExpenseType"), Value = formData.ExpenseType });



        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "EventId/EventRequestId"), Value = val });







        //            var addeddatarow = smartsheet.SheetResources.RowResources.AddRows(sheet4.Id.Value, new Row[] { newRow1 });

        //            var columnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
        //            var Cell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == columnId);
        //            var value = Cell.DisplayValue;
        //            if (FCPA == "Yes")
        //            {
        //                var filename = " FCPA";
        //                var filePath = SheetHelper.testingFile(formDataList.HcpConsultant.FcpaFile, value, filename);
        //                var addedRow = addedRows[0];
        //                var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");
        //            }
        //        }

        //        foreach (var formdata in formDataList.BrandsList)
        //        {
        //            var newRow2 = new Row();
        //            newRow2.Cells = new List<Cell>();
        //            newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "% Allocation"), Value = formdata.PercentAllocation });
        //            newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Brands"), Value = formdata.BrandName });
        //            newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Project ID"), Value = formdata.ProjectId });
        //            newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "EventId/EventRequestId"), Value = val });


        //            smartsheet.SheetResources.RowResources.AddRows(sheet2.Id.Value, new Row[] { newRow2 });

        //        }


        //        foreach (var formdata in formDataList.ExpenseSheet)
        //        {
        //            var newRow6 = new Row();
        //            newRow6.Cells = new List<Cell>();

        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Expense"), Value = formdata.Expense });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "EventId/EventRequestID"), Value = val });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formdata.BTC_BTE });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTCAmount"), Value = formdata.BtcAmount });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTEAmount"), Value = formdata.BteAmount });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Registration Amount"), Value = formdata.RegstAmount });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount"), Value = formdata.ExpenseAmount });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Topic"), Value = formDataList.HcpConsultant.EventTopic });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Type"), Value = formDataList.HcpConsultant.EventType });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Date Start"), Value = formDataList.HcpConsultant.EventDate });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event End Date"), Value = formDataList.HcpConsultant.EventEndDate });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Venue name"), Value = formDataList.HcpConsultant.VenueName });

        //            smartsheet.SheetResources.RowResources.AddRows(sheet6.Id.Value, new Row[] { newRow6 });
        //        }




        //        var targetRow = addedRows[0];
        //        long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role");
        //        var cellToUpdateB = new Cell { ColumnId = honorariumSubmittedColumnId, Value = formDataList.HcpConsultant.Role };
        //        Row updateRow = new Row { Id = targetRow.Id, Cells = new Cell[] { cellToUpdateB } };
        //        var cellToUpdate = targetRow.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
        //        if (cellToUpdate != null) { cellToUpdate.Value = formDataList.HcpConsultant.Role; }

        //        smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRow });

        //        return Ok(new
        //        { Message = " Success!" });
        //    }
        //    catch (Exception ex)
        //    {
        //        return BadRequest($"Could not find {ex.Message}");
        //    }
        //}


        //[HttpPost("MedicalUtilityPreEvent"), DisableRequestSizeLimit]
        //public IActionResult MedicalUtilityPreEvent(MedicalUtilityPreEventPayload formDataList)
        //{


        //    string sheetId1 = configuration.GetSection("SmartsheetSettings:Class1").Value;
        //    string sheetId2 = configuration.GetSection("SmartsheetSettings:EventRequestBrandsList").Value;
        //    //string sheetId3 = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
        //    string sheetId4 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
        //    //string sheetId5 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpSlideKit").Value;
        //    string sheetId6 = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
        //    string sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;


        //    Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
        //    Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
        //    //Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
        //    Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
        //    //Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
        //    Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
        //    Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);



        //    //string sheetId1 = configuration.GetSection("SmartsheetSettings:Class1").Value;
        //    //string sheetId2 = configuration.GetSection("SmartsheetSettings:EventRequestBrandsList").Value;
        //    //string sheetId4 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
        //    //string sheetId6 = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
        //    //string sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;



        //    //long.TryParse(sheetId1, out long parsedSheetId1);
        //    //long.TryParse(sheetId2, out long parsedSheetId2);
        //    //long.TryParse(sheetId4, out long parsedSheetId4);
        //    //long.TryParse(sheetId6, out long parsedSheetId6);
        //    //long.TryParse(sheetId7, out long parsedSheetId7);

        //    //Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId1, null, null, null, null, null, null, null);
        //    //Sheet sheet2 = smartsheet.SheetResources.GetSheet(parsedSheetId2, null, null, null, null, null, null, null);
        //    //Sheet sheet4 = smartsheet.SheetResources.GetSheet(parsedSheetId4, null, null, null, null, null, null, null);
        //    //Sheet sheet6 = smartsheet.SheetResources.GetSheet(parsedSheetId6, null, null, null, null, null, null, null);
        //    //Sheet sheet7 = smartsheet.SheetResources.GetSheet(parsedSheetId7, null, null, null, null, null, null, null);

        //    StringBuilder addedBrandsData = new StringBuilder();
        //    StringBuilder addedHcpData = new StringBuilder();
        //    StringBuilder addedExpences = new StringBuilder();

        //    int addedHcpDataNo = 1;
        //    int addedBrandsDataNo = 1;
        //    int addedExpencesNo = 1;


        //    var TotalExpenseAmount = 0;

        //    CultureInfo hindi = new CultureInfo("hi-IN");

        //    //var EventOpen30Days = !string.IsNullOrEmpty(formDataList.HcpConsultant.EventOpen30days) ? "Yes" : "No";
        //    var EventOpen30Days = !string.IsNullOrEmpty(formDataList.MedicalUtilityData.EventOpen30daysFile) ? "Yes" : "No";
        //    var EventWithin7Days = !string.IsNullOrEmpty(formDataList.MedicalUtilityData.EventWithin7daysFile) ? "Yes" : "No";
        //    var UploadDeviationFile = !string.IsNullOrEmpty(formDataList.MedicalUtilityData.UploadDeviationFile) ? "Yes" : "No";
        //    var FCPA = "";


        //    //if (formDataList.MedicalUtilityData.EventWithin7daysFile != "")
        //    //{
        //    //    EventWithin7Days = "Yes";
        //    //}
        //    //else
        //    //{
        //    //    EventWithin7Days = "No";
        //    //}
        //    //if (formDataList.MedicalUtilityData.EventOpen30daysFile != "")
        //    //{
        //    //    EventOpen30Days = "Yes";
        //    //}
        //    //else
        //    //{
        //    //    EventOpen30Days = "No";
        //    //}
        //    //if (formDataList.MedicalUtilityData.UploadDeviationFile != "")
        //    //{
        //    //    UploadDeviationFile = "Yes";
        //    //}
        //    //else
        //    //{
        //    //    UploadDeviationFile = "No";
        //    //}




        //    foreach (var formdata in formDataList.ExpenseSheet)
        //    {
        //        string rowData = $"{addedExpencesNo}. {formdata.Expense} | TotalAmount: {formdata.TotalExpenseAmount}| {formdata.BTC_BTE}";
        //        addedExpences.AppendLine(rowData);
        //        addedExpencesNo++;
        //        var amount = int.Parse(formdata.TotalExpenseAmount);
        //        TotalExpenseAmount = TotalExpenseAmount + amount;

        //    }
        //    string Expense = addedExpences.ToString();



        //    foreach (var formdata in formDataList.BrandsList)
        //    {

        //        string rowData = $"{addedBrandsDataNo}. {formdata.BrandName} | {formdata.ProjectId} | {formdata.PercentAllocation}";
        //        addedBrandsData.AppendLine(rowData);
        //        addedBrandsDataNo++;
        //    }
        //    string brand = addedBrandsData.ToString();




        //    foreach (var formdata in formDataList.HcpList)
        //    {
        //        string rowData = $"{addedHcpDataNo}. {formdata.MisCode} |{formdata.HcpName} |Speciality: {formdata.Speciality} |Tier: {formdata.Tier} ";

        //        addedHcpData.AppendLine(rowData);
        //        addedHcpDataNo++;

        //    }
        //    string HCP = addedHcpData.ToString();



        //    var total = TotalExpenseAmount;






        //    try
        //    {

        //        var newRow = new Row();
        //        newRow.Cells = new List<Cell>();
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.MedicalUtilityData.EventTopic });

        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventType"), Value = formDataList.MedicalUtilityData.EventType });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventDate"), Value = formDataList.MedicalUtilityData.EventDate });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Valid From"), Value = formDataList.MedicalUtilityData.ValidFrom });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Valid To"), Value = formDataList.MedicalUtilityData.ValidTill });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Utility Type"), Value = formDataList.MedicalUtilityData.MedicalUtilityType });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Utility Description"), Value = formDataList.MedicalUtilityData.MedicalUtilityDescription });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "IsAdvanceRequired"), Value = formDataList.MedicalUtilityData.IsAdvanceRequired });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Advance Amount"), Value = int.Parse(formDataList.MedicalUtilityData.AdvanceAmount )});
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Brands"), Value = brand });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Expenses"), Value = Expense });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Panelists"), Value = HCP });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "InitiatorName"), Value = formDataList.MedicalUtilityData.InitiatorName });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense"), Value = TotalExpenseAmount });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Budget"), Value = total });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Email"), Value = formDataList.MedicalUtilityData.Initiator_Email });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "RBM/BM"), Value = formDataList.MedicalUtilityData.RBMorBM });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Head"), Value = formDataList.MedicalUtilityData.Sales_Head });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Coordinator"), Value = formDataList.MedicalUtilityData.SalesCoordinatorEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Marketing Head"), Value = formDataList.MedicalUtilityData.Marketing_Head });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Compliance"), Value = formDataList.MedicalUtilityData.ComplianceEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Accounts"), Value = formDataList.MedicalUtilityData.FinanceAccountsEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Treasury"), Value = formDataList.MedicalUtilityData.Finance });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Reporting Manager"), Value = formDataList.MedicalUtilityData.ReportingManagerEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "1 Up Manager"), Value = formDataList.MedicalUtilityData.FirstLevelEmail });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Affairs Head"), Value = formDataList.MedicalUtilityData.MedicalAffairsEmail });
        //        //newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role"), Value = formDataList.MedicalUtilityData.Role });
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, " Total Expense BTC"), Value = int.Parse(formDataList.MedicalUtilityData.TotalExpenseBTC )});
        //        newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense BTE"), Value = int.Parse(formDataList.MedicalUtilityData.TotalExpenseBTE )});


        //        var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheet1.Id.Value, new Row[] { newRow });

        //        var eventIdColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
        //        var eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
        //        var val = eventIdCell.DisplayValue;


        //        List<string> DeviationNames = new List<string>();
        //        DeviationNames.Add($"30DaysDeviationFile:{formDataList.MedicalUtilityData.EventOpen30daysFile}");
        //        DeviationNames.Add($"7DaysDeviationFile:{formDataList.MedicalUtilityData.EventWithin7daysFile}");
        //        DeviationNames.Add($"AgregateSpendDeviationFile:{formDataList.MedicalUtilityData.UploadDeviationFile}");

        //        if (EventOpen30Days == "Yes" || EventWithin7Days == "Yes" || UploadDeviationFile == "Yes")
        //        {
        //            var eventId = val;
        //            foreach (var name in DeviationNames)
        //            {
        //                var y = name.Split(':');
        //                var fn = y[0];
        //                var bs = y[1];

        //                if (bs != "")
        //                {
        //                    try
        //                    {
        //                        var newRow7 = new Row();
        //                        newRow7.Cells = new List<Cell>();
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.MedicalUtilityData.EventTopic });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formDataList.MedicalUtilityData.EventType });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formDataList.MedicalUtilityData.EventDate });
        //                        if (fn == "30DaysDeviationFile")
        //                        {
        //                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = EventOpen30Days });
        //                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = formDataList.MedicalUtilityData.EventOpen30dayscount });
        //                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Outstanding with initiator for more than 45 days" });

        //                        }
        //                        else if (fn == "7DaysDeviationFile")
        //                        {
        //                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = EventWithin7Days });
        //                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "5 days from the Event Date" });

        //                        }
        //                        else if (fn == "AgregateSpendDeviationFile")
        //                        {
        //                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HCP exceeds 1,00,000 Trigger"), Value = UploadDeviationFile });
        //                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Aggregate Limit of 1,00,000 is Exceeded" });
        //                        }
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.MedicalUtilityData.Sales_Head });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formDataList.MedicalUtilityData.FinanceHead });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formDataList.MedicalUtilityData.InitiatorName });
        //                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.MedicalUtilityData.Initiator_Email });


        //                        var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });

        //                        var filename = fn;
        //                        var filePath = SheetHelper.testingFile(bs, val, filename);



        //                        var addedRow = addeddeviationrow[0];

        //                        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
        //                                sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");

        //                        if (System.IO.File.Exists(filePath))
        //                        {
        //                            System.IO.File.Delete(filePath);
        //                        }

        //                    }
        //                    catch (Exception ex)
        //                    {
        //                        return BadRequest(ex.Message);
        //                    }
        //                }
        //            }


        //        }





        //        foreach (var formData in formDataList.HcpList)
        //        {
        //            var newRow1 = new Row();
        //            newRow1.Cells = new List<Cell>();
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCPName"), Value = formData.HcpName });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "MISCode"), Value = formData.MisCode });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCP Type"), Value = formData.HcpType });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Speciality"), Value = formData.Speciality });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Tier"), Value = formData.Tier });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Medical Utility Cost"), Value = int.Parse(formData.MedicalUtilityCostAmount) });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Medical Utility Type"), Value = formDataList.MedicalUtilityData.MedicalUtilityType });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Medical Utility description"), Value = formDataList.MedicalUtilityData.MedicalUtilityDescription });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Legitimate Need"), Value = formData.Legitimate });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Objective Criteria"), Value = formData.Objective });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Rationale"), Value = formData.Rationale });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Topic"), Value = formDataList.MedicalUtilityData.EventTopic });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Type"), Value = formDataList.MedicalUtilityData.EventType });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Date Start"), Value = formDataList.MedicalUtilityData.ValidFrom });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event End Date"), Value = formDataList.MedicalUtilityData.ValidTill });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "EventId/EventRequestId"), Value = val });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "ExpenseType"), Value = formData.ExpenseType });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "FCPA Date"), Value = formData.Fcpadate });
        //            newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Request Date"), Value = formData.HCPRequestDate });

        //            //Request Date,fcpa date,Expense type
        //            var addeddatarows = smartsheet.SheetResources.RowResources.AddRows(sheet4.Id.Value, new Row[] { newRow1 });


        //            var FCPAFile = !string.IsNullOrEmpty(formData.UploadFCPA) ? "Yes" : "No";
        //            var UploadWrittenRequestDate = !string.IsNullOrEmpty(formData.UploadWrittenRequestDate) ? "Yes" : "No";
        //            var Invoice_Brouchere_Quotation = !string.IsNullOrEmpty(formData.Invoice_Brouchere_Quotation) ? "Yes" : "No";



        //            var columnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
        //            var Cell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == columnId);
        //            var value = Cell.DisplayValue;

        //            if (FCPAFile == "Yes")
        //            {

        //                var filename = " FCPA";
        //                var filePath = SheetHelper.testingFile(formData.UploadFCPA, val, filename);
        //                var addedRow = addedRows[0];
        //                var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
        //                        sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");
        //                if (System.IO.File.Exists(filePath))
        //                {
        //                    SheetHelper.DeleteFile(filePath);
        //                }
        //            }
        //            if (UploadWrittenRequestDate == "Yes")
        //            {
        //                var filename = " UploadWrittenRequestDate";
        //                var filePath = SheetHelper.testingFile(formData.UploadWrittenRequestDate, val, filename);
        //                var addedRow = addedRows[0];
        //                var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
        //                        sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");
        //                if (System.IO.File.Exists(filePath))
        //                {
        //                    SheetHelper.DeleteFile(filePath);
        //                }
        //            }

        //            if (Invoice_Brouchere_Quotation == "Yes")
        //            {
        //                var filename = " Invoice_Brouchere_Quotation";
        //                var filePath = SheetHelper.testingFile(formData.Invoice_Brouchere_Quotation, val, filename);
        //                var addedRow = addedRows[0];
        //                var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
        //                        sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");
        //                if (System.IO.File.Exists(filePath))
        //                {
        //                    System.IO.File.Delete(filePath);
        //                }
        //            }





        //        }

        //        foreach (var formdata in formDataList.BrandsList)
        //        {
        //            var newRow2 = new Row();
        //            newRow2.Cells = new List<Cell>();
        //            newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "% Allocation"), Value = formdata.PercentAllocation });
        //            newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Brands"), Value = formdata.BrandName });
        //            newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Project ID"), Value = formdata.ProjectId });
        //            newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "EventId/EventRequestId"), Value = val });

        //            smartsheet.SheetResources.RowResources.AddRows(sheet2.Id.Value, new Row[] { newRow2 });

        //        }


        //        foreach (var formdata in formDataList.ExpenseSheet)
        //        {
        //            var newRow6 = new Row();
        //            newRow6.Cells = new List<Cell>();
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "MisCode"), Value = formdata.MisCode });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Expense"), Value = formdata.Expense });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "EventId/EventRequestID"), Value = val });

        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formdata.BTC_BTE });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount"), Value = formdata.TotalExpenseAmount });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTCAmount"), Value = formdata.BtcAmount });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTEAmount"), Value = formdata.BteAmount });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Topic"), Value = formDataList.MedicalUtilityData.EventTopic });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Type"), Value = formDataList.MedicalUtilityData.EventType });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Date Start"), Value = formDataList.MedicalUtilityData.ValidFrom });
        //            newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event End Date"), Value = formDataList.MedicalUtilityData.ValidTill });

        //            smartsheet.SheetResources.RowResources.AddRows(sheet6.Id.Value, new Row[] { newRow6 });
        //        }



        //        var targetRow = addedRows[0];
        //        long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role");
        //        var cellToUpdateB = new Cell { ColumnId = honorariumSubmittedColumnId, Value = formDataList.MedicalUtilityData.Role };
        //        Row updateRow = new Row { Id = targetRow.Id, Cells = new Cell[] { cellToUpdateB } };
        //        var cellToUpdate = targetRow.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
        //        if (cellToUpdate != null) { cellToUpdate.Value = formDataList.MedicalUtilityData.Role; }

        //        smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRow });

        //        return Ok(new
        //        { Message = " Success!" });
        //    }

        //    catch (Exception ex)
        //    {
        //        return BadRequest($"Could not find {ex.Message}");
        //    }
        //}

        #endregion

        [HttpPost("HandsOnPreEvent"), DisableRequestSizeLimit]
        public IActionResult HandsOnPreEvent(HandsOnTrainingPreEvet formDataList)
        {

            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
            Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
            Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
            Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
            Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
            Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);
            Sheet sheet8 = SheetHelper.GetSheetById(smartsheet, sheetId8);
            Sheet sheet9 = SheetHelper.GetSheetById(smartsheet, sheetId9);

            StringBuilder addedBrandsData = new StringBuilder();
            StringBuilder addedInviteesData = new StringBuilder();
            StringBuilder addedHcpData = new StringBuilder();
            StringBuilder addedSlideKitData = new StringBuilder();
            StringBuilder addedExpences = new StringBuilder();
            StringBuilder SelectedProduct = new StringBuilder();
            StringBuilder addedMEnariniInviteesData = new StringBuilder();
            int addedSlideKitDataNo = 1;
            int addedHcpDataNo = 1;
            int addedInviteesDataNo = 1;
            int addedBrandsDataNo = 1;
            int addedExpencesNo = 1;
            int SelectedProductNo = 1;
            int addedInviteesDataNoforMenarini = 1;

            foreach (var formdata in formDataList.AttenderSelections)
            {
                if (formdata.InviteedFrom == "Menarini Employees")
                {
                    string row = $"{addedInviteesDataNoforMenarini}. {formdata.AttenderName}";
                    addedMEnariniInviteesData.AppendLine(row);
                    addedInviteesDataNoforMenarini++;
                }
                else
                {
                    string rowData = $"{addedInviteesDataNo}. {formdata.AttenderName}";
                    addedInviteesData.AppendLine(rowData);
                    addedInviteesDataNo++;
                }
            }
            string Invitees = addedInviteesData.ToString();
            string MenariniInvitees = addedMEnariniInviteesData.ToString();
            foreach (var formdata in formDataList.ExpenseData)
            {
                string rowData = $"{addedExpencesNo}. {formdata.ExpenseType} | AmountExcludingTax: {formdata.ExpenseAmountExcludingTax}| Amount: {formdata.ExpenseAmountIncludingTax} | {formdata.IsBtcorBte}";
                addedExpences.AppendLine(rowData);
                addedExpencesNo++;
            }
            string Expense = addedExpences.ToString();
            foreach (var formdata in formDataList.ProductSelections)
            {
                string rowData = $"{SelectedProductNo}. Product Selection Type {formdata.ProductSelectionType} | Product Name: {formdata.ProductName}| Samples Requires: {formdata.SamplesRequires} ";
                SelectedProduct.AppendLine(rowData);
                SelectedProductNo++;
            }
            string SelectedProductData = SelectedProduct.ToString();
            foreach (var formdata in formDataList.SlideKitSelectionData)
            {
                string rowData = $"{addedSlideKitDataNo}. {formdata.MISCode} | {formdata.SlideKitSelection}";
                addedSlideKitData.AppendLine(rowData);
                addedSlideKitDataNo++;
            }
            string slideKit = addedSlideKitData.ToString();
            foreach (var formdata in formDataList.EventBrandsList)
            {
                string rowData = $"{addedBrandsDataNo}. {formdata.BrandName} | {formdata.ProjectId} | {formdata.PercentAllocation}";
                addedBrandsData.AppendLine(rowData);
                addedBrandsDataNo++;
            }
            string brand = addedBrandsData.ToString();
            foreach (var formdata in formDataList.TrainerDetails)
            {

                string rowData = $"{addedHcpDataNo}. {formdata.TrainerType} |{formdata.TrainerName} | Honr.Amt: {formdata.HonorariumAmountincludingTax} |Trav.&Acc.Amt: {formdata.TravelAmountIncludingTax + formdata.AccomodationAmountIncludingTax} ";
                addedHcpData.AppendLine(rowData);
                addedHcpDataNo++;

            }
            string HCP = addedHcpData.ToString();
            string beneficiaryDetails = $"Currency : {formDataList.HandsOnTraining.BenificiaryDetailsData.Currency} | BenificiaryName :  {formDataList.HandsOnTraining.BenificiaryDetailsData.BenificiaryName} | BankAccountNumber : {formDataList.HandsOnTraining.BenificiaryDetailsData.BenificiaryName}";
            try
            {

                Row newRow = new()
                {
                    Cells = new List<Cell>()
                };

                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventDate"), Value = formDataList.HandsOnTraining.EventDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventType"), Value = formDataList.HandsOnTraining.EventType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.HandsOnTraining.EventName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "StartTime"), Value = formDataList.HandsOnTraining.EventStartTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EndTime"), Value = formDataList.HandsOnTraining.EventEndTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Product Brand"), Value = formDataList.HandsOnTraining.ProductBrandSelection });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Mode of Training"), Value = formDataList.HandsOnTraining.ModeOfTraining });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "VenueName"), Value = formDataList.HandsOnTraining.VenueName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "City"), Value = formDataList.HandsOnTraining.City });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "State"), Value = formDataList.HandsOnTraining.State });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "HOT Webinar Vendor Name"), Value = formDataList.HandsOnTraining.VendorName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "HOT Webinar Type"), Value = formDataList.HandsOnTraining.WebinarType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Venue Selection Checklist"), Value = formDataList.HandsOnTraining.VenueSelectionCheckList });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Emergency Support"), Value = formDataList.HandsOnTraining.IsVenueHasAnyEmergancySupport });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Emergency Contact No"), Value = formDataList.HandsOnTraining.EmerganctContact });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Facility Charges"), Value = formDataList.HandsOnTraining.IsVenueFacilityCharges });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Facility Charges BTC/BTE"), Value = formDataList.HandsOnTraining.VenueFacilityChargesBtc_Bte });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Facility charges Excluding Tax"), Value = formDataList.HandsOnTraining.FacilityChargesExcludingTax });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Facility charges including Tax"), Value = formDataList.HandsOnTraining.FacilityChargesIncludingTax });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Anesthetist Required?"), Value = formDataList.HandsOnTraining.IsAnesthetistRequired });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Anesthetist BTC/BTE"), Value = formDataList.HandsOnTraining.AnesthetistRequiredBtc_Bte });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Anesthetist Excluding Tax"), Value = formDataList.HandsOnTraining.AnesthetistChargesExcludingTax });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Anesthetist including Tax"), Value = formDataList.HandsOnTraining.AnesthetistChargesIncludingTax });

                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "InitiatorName"), Value = formDataList.HandsOnTraining.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Advance Amount"), Value = formDataList.HandsOnTraining.AdvanceAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, " Total Expense BTC"), Value = formDataList.HandsOnTraining.TotalExpenseBTC });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense BTE"), Value = formDataList.HandsOnTraining.TotalExpenseBTE });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Honorarium Amount"), Value = formDataList.HandsOnTraining.TotalHonorariumAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel Amount"), Value = formDataList.HandsOnTraining.TotalTravelAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel & Accommodation Amount"), Value = formDataList.HandsOnTraining.TotalTravelAccommodationAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Accomodation Amount"), Value = formDataList.HandsOnTraining.TotalAccomodationAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Budget"), Value = formDataList.HandsOnTraining.TotalBudget });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Local Conveyance"), Value = formDataList.HandsOnTraining.TotalLocalConveyance });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense"), Value = formDataList.HandsOnTraining.TotalExpense });

                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Email"), Value = formDataList.HandsOnTraining.InitiatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "RBM/BM"), Value = formDataList.HandsOnTraining.RBMorBMEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Head"), Value = formDataList.HandsOnTraining.SalesHeadEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Coordinator"), Value = formDataList.HandsOnTraining.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Marketing Head"), Value = formDataList.HandsOnTraining.MarketingHeadEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Treasury"), Value = formDataList.HandsOnTraining.FinanceEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Compliance"), Value = formDataList.HandsOnTraining.ComplianceEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Accounts"), Value = formDataList.HandsOnTraining.FinanceAccountsEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Reporting Manager"), Value = formDataList.HandsOnTraining.ReportingManagerEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "1 Up Manager"), Value = formDataList.HandsOnTraining.FirstLevelEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Affairs Head"), Value = formDataList.HandsOnTraining.MedicalAffairsEmail });

                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Panelists"), Value = HCP });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Invitees"), Value = Invitees });

                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Brands"), Value = brand });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Expenses"), Value = Expense });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "SlideKits"), Value = slideKit });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Beneficiary Details"), Value = beneficiaryDetails });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Selected Products"), Value = SelectedProductData });

                IList<Row> addedRows = smartsheet.SheetResources.RowResources.AddRows(sheet1.Id.Value, new Row[] { newRow });

                long eventIdColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                Cell? eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                string val = eventIdCell.DisplayValue;

                int x = 1;
                foreach (string p in formDataList.HandsOnTraining.Files)
                {
                    string[] words = p.Split(':');
                    string r = words[0];
                    string q = words[1];
                    string name = r.Split(".")[0];
                    string filePath = SheetHelper.testingFile(q, val, name);
                    Row addedRow = addedRows[0];
                    Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                           sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                    x++;


                    if (System.IO.File.Exists(filePath))
                    {
                        SheetHelper.DeleteFile(filePath);
                    }
                }

                if (formDataList.HandsOnTraining.IsDeviationUpload == "Yes")
                {
                    List<string> DeviationNames = new();
                    foreach (string p in formDataList.HandsOnTraining.DeviationFiles)
                    {
                        string[] words = p.Split(':');
                        string r = words[0];

                        DeviationNames.Add(r);
                    }
                    foreach (string deviationname in DeviationNames)
                    {
                        string file = deviationname.Split(".")[0];
                        string eventId = val;
                        try
                        {
                            Row newRow7 = new()
                            {
                                Cells = new List<Cell>()
                            };

                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.HandsOnTraining.EventName });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formDataList.HandsOnTraining.EventType });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formDataList.HandsOnTraining.EventDate });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "StartTime"), Value = formDataList.HandsOnTraining.EventStartTime });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EndTime"), Value = formDataList.HandsOnTraining.EventEndTime });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "VenueName"), Value = formDataList.HandsOnTraining.VenueName });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "City"), Value = formDataList.HandsOnTraining.City });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "State"), Value = formDataList.HandsOnTraining.State });


                            if (file == "30DaysDeviationFile")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Outstanding with initiator for more than 45 days" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = "Yes" });//formDataList.HandsOnTraining.EventOpen30days });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = formDataList.HandsOnTraining.EventOpen30dayscount });
                            }
                            else if (file == "7DaysDeviationFile")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "5 days from the Event Date" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = "Yes" });//formDataList.HandsOnTraining.EventWithin7days });

                            }
                            else if (file == "ExpenseExcludingTax")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Food and Beverages expense exceeds 1500" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "PRE-F&B Expense Excluding Tax"), Value = "Yes" });
                            }
                            else if (file.Contains("Travel_Accomodation3LExceededFile"))
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Travel/AccomodationAggregate Limit of 3,00,000 is Exceeded" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Travel/Accomodation 3,00,000 Exceeded Trigger"), Value = "Yes" });//formDataList.class1.FB_Expense_Excluding_Tax });
                            }
                            else if (file.Contains("TrainerHonorarium12LExceededFile"))
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Honorarium Aggregate Limit of 12,00,000 is Exceeded" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Trainer Honorarium 12,00,000 Exceeded Trigger"), Value = "Yes" }); //formDataList.class1.FB_Expense_Excluding_Tax });
                            }
                            else if (file.Contains("HCPHonorarium6LExceededFile"))
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Honorarium Aggregate Limit of 6,00,000 is Exceeded" });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HCP Honorarium 6,00,000 Exceeded Trigger"), Value = "Yes" }); // formDataList.class1.FB_Expense_Excluding_Tax });
                            }
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.HandsOnTraining.SalesHeadEmail });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formDataList.HandsOnTraining.FinanceEmail });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formDataList.HandsOnTraining.InitiatorName });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.HandsOnTraining.InitiatorEmail });

                            IList<Row> addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });

                            int j = 1;
                            foreach (var p in formDataList.HandsOnTraining.DeviationFiles)
                            {
                                string[] words = p.Split(':');
                                string r = words[0];
                                string q = words[1];
                                if (deviationname == r)
                                {
                                    string name = r.Split(".")[0];
                                    string filePath = SheetHelper.testingFile(q, val, name);
                                    Row addedRow = addeddeviationrow[0];
                                    Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                            sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");
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

                //List<Row> newRows = new();
                //foreach (var formdata in formDataList.RequestBrandsList)
                //{
                //    Row newRow2 = new()
                //    {
                //        Cells = new List<Cell>()
                //        {
                //            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "% Allocation"), Value = formdata.PercentAllocation },
                //            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Brands"), Value = formdata.BrandName },
                //            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Project ID"), Value = formdata.ProjectId },
                //            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "EventId/EventRequestId"), Value = val }
                //        }
                //    };

                //    newRows.Add(newRow2);
                //}

                //smartsheet.SheetResources.RowResources.AddRows(sheet2.Id.Value, newRows.ToArray());


                foreach (var formData in formDataList.TrainerDetails)
                {
                    Row newRow1 = new()
                    {
                        Cells = new List<Cell>()
                    };

                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "MISCode"), Value = formData.MISCode });
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
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Topic"), Value = formDataList.HandsOnTraining.EventName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Type"), Value = formDataList.HandsOnTraining.EventType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Venue name"), Value = formDataList.HandsOnTraining.VenueName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Date Start"), Value = formDataList.HandsOnTraining.EventDate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event End Date"), Value = formDataList.HandsOnTraining.EventDate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TotalSpend"), Value = formData.FinalAmount });

                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "EventId/EventRequestId"), Value = val });

                    IList<Row> row = smartsheet.SheetResources.RowResources.AddRows(sheet4.Id.Value, new Row[] { newRow1 });

                    foreach (string p in formData.TrainerFiles)
                    {
                        string[] words = p.Split(':');
                        string r = words[0];
                        string q = words[1];
                        string name = r.Split(".")[0];
                        string filePath = SheetHelper.testingFile(q, val, name);
                        Row addedRow = row[0];
                        Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                               sheet4.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                        x++;


                        if (System.IO.File.Exists(filePath))
                        {
                            SheetHelper.DeleteFile(filePath);
                        }
                    }
                }

                List<Row> newRows2 = new();
                foreach (var formdata in formDataList.EventBrandsList)
                {
                    Row newRow2 = new()
                    {
                        Cells = new List<Cell>()
                        {
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "% Allocation"), Value = formdata.PercentAllocation },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Brands"), Value = formdata.BrandName },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Project ID"), Value = formdata.ProjectId },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "EventId/EventRequestId"), Value = val }
                        }
                    };

                    newRows2.Add(newRow2);
                }

                smartsheet.SheetResources.RowResources.AddRows(sheet2.Id.Value, newRows2.ToArray());


                //foreach (var formdata in formDataList.EventBrandsList)
                //{
                //    Row newRow2 = new()
                //    {
                //        Cells = new List<Cell>()
                //    };
                //    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "% Allocation"), Value = formdata.PercentAllocation });
                //    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Brands"), Value = formdata.BrandName });
                //    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "Project ID"), Value = formdata.ProjectId });
                //    newRow2.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet2, "EventId/EventRequestId"), Value = val });
                //    smartsheet.SheetResources.RowResources.AddRows(sheet2.Id.Value, new Row[] { newRow2 });
                //}
                foreach (var formdata in formDataList.AttenderSelections)
                {
                    Row newRow3 = new()
                    {
                        Cells = new List<Cell>()
                    };
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "MISCode"), Value = formdata.MisCode });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "HCPName"), Value = formdata.AttenderName });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "HCP Type"), Value = formdata.AttenderType });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Designation"), Value = formdata.Designation });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Employee Code"), Value = formdata.EmployeeCode });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "LocalConveyance"), Value = formdata.IsLocalConveyance });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "BTC/BTE"), Value = formdata.IsBtcorBte });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "LcAmount"), Value = formdata.LocalConveyanceAmountIncludingTax });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Lc Amount Excluding Tax"), Value = formdata.LocalConveyanceAmountExcludingTax });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "EventId/EventRequestId"), Value = val });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Invitee Source"), Value = formdata.InviteedFrom });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Qualification"), Value = formdata.Qualification });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Experience"), Value = formdata.Experiance });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Topic"), Value = formDataList.HandsOnTraining.EventName });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Type"), Value = formDataList.HandsOnTraining.EventType });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Venue name"), Value = formDataList.HandsOnTraining.VenueName });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Date Start"), Value = formDataList.HandsOnTraining.EventDate });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event End Date"), Value = formDataList.HandsOnTraining.EventDate });

                    IList<Row> row = smartsheet.SheetResources.RowResources.AddRows(sheet3.Id.Value, new Row[] { newRow3 });

                    foreach (string p in formdata.AttenderFiles)
                    {
                        string[] words = p.Split(':');
                        string r = words[0];
                        string q = words[1];
                        string name = r.Split(".")[0];
                        string filePath = SheetHelper.testingFile(q, val, name);
                        Row addedRow = row[0];
                        Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                               sheet3.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                        x++;


                        if (System.IO.File.Exists(filePath))
                        {
                            SheetHelper.DeleteFile(filePath);
                        }
                    }
                }
                foreach (var formdata in formDataList.SlideKitSelectionData)
                {
                    Row newRow5 = new()
                    {
                        Cells = new List<Cell>()
                    };
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "MIS"), Value = formdata.MISCode });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "Slide Kit Type"), Value = formdata.SlideKitSelection });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "SlideKit Document"), Value = formdata.SlideKitDocument });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "EventId/EventRequestId"), Value = val });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "Fillers Indication"), Value = formdata.IndicatorsForFillers });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "Threads Indication"), Value = formdata.IndicatorsForThreads });

                    IList<Row> row = smartsheet.SheetResources.RowResources.AddRows(sheet5.Id.Value, new Row[] { newRow5 });
                    if (formdata.IsUpload == "Yes")
                    {
                        string[] words = formdata.DocToUpload.Split(':');
                        string r = words[0];
                        string q = words[1];
                        string name = r.Split(".")[0];
                        string filePath = SheetHelper.testingFile(q, val, name);
                        Row addedRow = row[0];
                        Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                               sheet5.Id.Value, addedRow.Id.Value, filePath, "application/msword");



                        if (System.IO.File.Exists(filePath))
                        {
                            SheetHelper.DeleteFile(filePath);
                        }
                    }
                }
                foreach (var formdata in formDataList.ExpenseData)
                {
                    Row newRow6 = new()
                    {
                        Cells = new List<Cell>()
                    };
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Expense"), Value = formdata.ExpenseType });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "EventId/EventRequestID"), Value = val });
                    //newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "AmountExcludingTax?"), Value = formdata.ExpenseAmountExcludingTax });


                    //newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "AmountExcludingTax?"), Value = formdata.AmountExcludingTax });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount Excluding Tax"), Value = formdata.ExpenseAmountExcludingTax });

                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount"), Value = formdata.ExpenseAmountIncludingTax });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formdata.IsBtcorBte });
                    //newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BudgetAmount"), Value = formdata.BudgetAmount });
                    //newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTCAmount"), Value = formdata.BtcAmount });
                    //newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTEAmount"), Value = formdata.BteAmount });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Topic"), Value = formDataList.HandsOnTraining.EventName });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Type"), Value = formDataList.HandsOnTraining.EventType });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Venue name"), Value = formDataList.HandsOnTraining.VenueName });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Date Start"), Value = formDataList.HandsOnTraining.EventDate });
                    newRow6.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event End Date"), Value = formDataList.HandsOnTraining.EventDate });

                    smartsheet.SheetResources.RowResources.AddRows(sheet6.Id.Value, new Row[] { newRow6 });
                }

                try
                {
                    Row newRow8 = new()
                    {
                        Cells = new List<Cell>()
                    };

                    newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "EventId/EventRequestId"), Value = val });
                    newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "EventType"), Value = formDataList.HandsOnTraining.EventType });
                    newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "EventDate"), Value = formDataList.HandsOnTraining.EventDate });
                    newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "VenueName"), Value = formDataList.HandsOnTraining.VenueName });
                    newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "City"), Value = formDataList.HandsOnTraining.City });
                    newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "State"), Value = formDataList.HandsOnTraining.State });
                    newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Other Currency"), Value = formDataList.HandsOnTraining.BenificiaryDetailsData.EnterCurrencyType });
                    newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Beneficiary Name"), Value = formDataList.HandsOnTraining.BenificiaryDetailsData.BenificiaryName });
                    newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Bank Account Number"), Value = formDataList.HandsOnTraining.BenificiaryDetailsData.BankAccountNumber });
                    newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Facility Charges"), Value = formDataList.HandsOnTraining.IsVenueFacilityCharges });
                    newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Anesthetist Required?"), Value = formDataList.HandsOnTraining.IsAnesthetistRequired });
                    newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Currency"), Value = formDataList.HandsOnTraining.BenificiaryDetailsData.Currency });
                    newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Bank Name"), Value = formDataList.HandsOnTraining.BenificiaryDetailsData.BankName });
                    newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "PAN card name"), Value = formDataList.HandsOnTraining.BenificiaryDetailsData.NameasPerPAN });
                    newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Pan Number"), Value = formDataList.HandsOnTraining.BenificiaryDetailsData.PANCardNumber });
                    newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "IFSC Code"), Value = formDataList.HandsOnTraining.BenificiaryDetailsData.IFSCCode });
                    //newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, ""), Value = formDataList.HandsOnTraining.BenificiaryDetailsData.IbnNumber });
                    //newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formDataList.HandsOnTraining.BenificiaryDetailsData.SwiftCode });
                    //newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formDataList.HandsOnTraining.BenificiaryDetailsData.TaxResidenceCertificateDate });
                    newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Email Id"), Value = formDataList.HandsOnTraining.BenificiaryDetailsData.EmailID });

                    smartsheet.SheetResources.RowResources.AddRows(sheet8.Id.Value, new Row[] { newRow8 });
                }
                catch (Exception ex)
                {
                    return BadRequest($"Could not find {ex.Message}");
                }
                foreach (var formdata in formDataList.ProductSelections)
                {
                    Row newRow9 = new()
                    {
                        Cells = new List<Cell>()
                    };
                    newRow9.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet9, "EventId/EventRequestId"), Value = val });
                    newRow9.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet9, "EventType"), Value = formDataList.HandsOnTraining.EventType });
                    newRow9.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet9, "EventDate"), Value = formDataList.HandsOnTraining.EventDate });
                    newRow9.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet9, "Event Topic"), Value = formDataList.HandsOnTraining.EventName });
                    newRow9.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet9, "Product Brand"), Value = formdata.ProductSelectionType });
                    newRow9.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet9, "Product Name"), Value = formdata.ProductName });
                    newRow9.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet9, "No of Samples Required"), Value = formdata.SamplesRequires });

                    var n = smartsheet.SheetResources.RowResources.AddRows(sheet9.Id.Value, new Row[] { newRow9 });
                }



                Row targetRow = addedRows[0];
                long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role");
                Cell cellToUpdateB = new() { ColumnId = honorariumSubmittedColumnId, Value = formDataList.HandsOnTraining.Role };
                Row updateRow = new() { Id = targetRow.Id, Cells = new Cell[] { cellToUpdateB } };
                Cell? cellToUpdate = targetRow.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                if (cellToUpdate != null) { cellToUpdate.Value = formDataList.HandsOnTraining.Role; }

                smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRow });

                return Ok(new
                { Message = " Success!" });
            }
            catch (Exception ex)
            {
                return BadRequest($"Could not find {ex.Message}");
            }
        }


        [HttpPost("AddPanelist"), DisableRequestSizeLimit]
        public IActionResult AddPanelist(EventRequestTrainerData formData)
        {
            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId4);
            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);

            try
            {
                Row newRow1 = new()
                {
                    Cells = new List<Cell>()
                };

                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "MISCode"), Value = formData.MISCode });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "HcpRole"), Value = formData.HCPRole });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "HCPName"), Value = formData.TrainerName });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "TrainerCode"), Value = formData.TrainerCode });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Qualification"), Value = formData.TrainerQualification });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Country"), Value = formData.TrainerCountry });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Speciality"), Value = formData.Speciality });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Tier"), Value = formData.TrainerCategory });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "HCP Type"), Value = formData.TrainerType });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Rationale"), Value = formData.Rationale });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "FCPA Date"), Value = formData.FCPAIssueDate });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "HonorariumRequired"), Value = formData.IsHonorariumApplicable });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "PresentationDuration"), Value = formData.Presentation_Speaking_WorkshopDuration });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "PanelSessionPreparationDuration"), Value = formData.DevelopmentofPresentationPanelSessionPreparation });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "PanelDiscussionDuration"), Value = formData.PaneldiscussionSessionduration });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "QASessionDuration"), Value = formData.QASession });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "BriefingSession"), Value = formData.Speaker_TrainerBriefing });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "TotalSessionHours"), Value = formData.TotalNoOfHours });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "LC BTC/BTE"), Value = formData.IsLCBTC_BTE });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Accomodation BTC/BTE"), Value = formData.IsAccomodationBTC_BTE });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Honorarium Amount Excluding Tax"), Value = formData.HonorariumAmountexcludingTax });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "HonorariumAmount"), Value = formData.HonorariumAmountincludingTax });
                //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(she1t1, "AgreementAmount"), Value = formData.AgreementAmount });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "YTD Spend Including Current Event"), Value = formData.YTDspendIncludingCurrentEvent });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Global FMV"), Value = formData.IsGlobalFMVCheck });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "ExpenseType"), Value = formData.ExpenseType });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Mode of Travel"), Value = formData.TravelSelection });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Travel BTC/BTE"), Value = formData.IsExpenseBTC_BTE });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Travel Excluding Tax"), Value = formData.TravelAmountExcludingTax });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Travel"), Value = formData.TravelAmountIncludingTax });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Accomodation Excluding Tax"), Value = formData.AccomodationAmountExcludingTax });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Accomodation"), Value = formData.AccomodationAmountIncludingTax });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Local Conveyance Excluding Tax"), Value = formData.LocalConveyanceAmountexcludingTax });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "LocalConveyance"), Value = formData.LocalConveyanceAmountincludingTax });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Travel/Accomodation Spend Including Current Event"), Value = formData.TravelandAccomodationspendincludingcurrentevent });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Currency"), Value = formData.BenificiaryDetailsData.Currency });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Other Currency"), Value = formData.BenificiaryDetailsData.EnterCurrencyType });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Beneficiary Name"), Value = formData.BenificiaryDetailsData.BenificiaryName });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Bank Account Number"), Value = formData.BenificiaryDetailsData.BankAccountNumber });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Bank Name"), Value = formData.BenificiaryDetailsData.BankName });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "PAN card name"), Value = formData.BenificiaryDetailsData.NameasPerPAN });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Pan Number"), Value = formData.BenificiaryDetailsData.PANCardNumber });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "IFSC Code"), Value = formData.BenificiaryDetailsData.IFSCCode });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Email Id"), Value = formData.BenificiaryDetailsData.EmailID });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "IBN Number"), Value = formData.BenificiaryDetailsData.IbnNumber });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Swift Code"), Value = formData.BenificiaryDetailsData.SwiftCode });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Tax Residence Certificate"), Value = formData.BenificiaryDetailsData.TaxResidenceCertificateDate });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "AgreementAmount"), Value = formData.AgreementAmount });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formData.EventName });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Type"), Value = formData.EventType });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Venue name"), Value = formData.VenueName });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Date Start"), Value = formData.EventStartDate });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event End Date"), Value = formData.EventEndDate });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "TotalSpend"), Value = formData.FinalAmount });
                newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId"), Value = formData.EventId });

                smartsheet.SheetResources.RowResources.AddRows(sheet1.Id.Value, new Row[] { newRow1 });



                //if (formData.IsDeviationUpload == "Yes")
                //{
                //    List<string> DeviationNames = new();
                //    foreach (string p in formData.DeviationFiles)
                //    {
                //        string[] words = p.Split(':');
                //        string r = words[0];

                //        DeviationNames.Add(r);
                //    }
                //    foreach (string deviationname in DeviationNames)
                //    {
                //        string file = deviationname.Split(".")[0];
                //        string eventId = formData.EventId;
                //        try
                //        {
                //            Row newRow7 = new()
                //            {
                //                Cells = new List<Cell>()
                //            };

                //            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                //            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formData.EventName });
                //            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formData.EventType });
                //            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formData.EventStartDate });
                //            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "StartTime"), Value = formData.EventStartTime });
                //            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EndTime"), Value = formData.EventEndTime });
                //            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "VenueName"), Value = formData.VenueName });
                //            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "City"), Value = formData.City });
                //            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "State"), Value = formData.State });


                //            if (file == "HCPHonorarium6LExceededFile")
                //            {
                //                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Outstanding with initiator for more than 45 days" });
                //                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = "Yes" });//formDataList.HandsOnTraining.EventOpen30days });
                //            }
                //            else if (file == "TrainerHonorarium12LExceededFile")
                //            {
                //                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "5 days from the Event Date" });
                //                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = "Yes" });//formDataList.HandsOnTraining.EventWithin7days });

                //            }
                //            else if (file == "Travel_Accomodation3LExceededFile")
                //            {
                //                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = "Food and Beverages expense exceeds 1500" });
                //                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "PRE-F&B Expense Excluding Tax"), Value = "Yes" });
                //            }
                //            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formData.SalesHeadEmail });
                //            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formData.FinanceEmail });
                //            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formData.InitiatorName });
                //            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formData.InitiatorEmail });

                //            IList<Row> addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });

                //            int j = 1;
                //            foreach (var p in formData.DeviationFiles)
                //            {
                //                string[] words = p.Split(':');
                //                string r = words[0];
                //                string q = words[1];
                //                if (deviationname == r)
                //                {
                //                    string name = r.Split(".")[0];
                //                    string filePath = SheetHelper.testingFile(q, eventId, name);
                //                    Row addedRow = addeddeviationrow[0];
                //                    Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                //                            sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                //                    j++;
                //                    if (System.IO.File.Exists(filePath))
                //                    {
                //                        SheetHelper.DeleteFile(filePath);
                //                    }
                //                }
                //            }
                //        }
                //        catch (Exception ex)
                //        {
                //            return BadRequest(ex.Message);
                //        }
                //    }
                //}

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }


            return Ok(new
            { Message = " Success!" });
        }


    }
}