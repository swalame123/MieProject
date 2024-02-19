using IndiaEventsWebApi.Models.EventTypeSheets;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Globalization;
using System.Text;

namespace IndiaEventsWebApi.Controllers.RequestSheets.Webinar
{
    [Route("api/[controller]")]
    [ApiController]
    public class WebinarTestControllerController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;

        public WebinarTestControllerController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

        }
        [HttpPost("AllObjModelsData")]
        public IActionResult AllObjModelsData(WebinarPayload formDataList)
        {

            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            string sheetId1 = configuration.GetSection("SmartsheetSettings:Class1").Value;
            string sheetId2 = configuration.GetSection("SmartsheetSettings:EventRequestBrandsList").Value;
            string sheetId3 = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
            string sheetId4 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
            string sheetId5 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpSlideKit").Value;
            string sheetId6 = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
            string sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;



            long.TryParse(sheetId1, out long parsedSheetId1);
            long.TryParse(sheetId2, out long parsedSheetId2);
            long.TryParse(sheetId3, out long parsedSheetId3);
            long.TryParse(sheetId4, out long parsedSheetId4);
            long.TryParse(sheetId5, out long parsedSheetId5);
            long.TryParse(sheetId6, out long parsedSheetId6);
            long.TryParse(sheetId7, out long parsedSheetId7);

            Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId1, null, null, null, null, null, null, null);
            Sheet sheet2 = smartsheet.SheetResources.GetSheet(parsedSheetId2, null, null, null, null, null, null, null);
            Sheet sheet3 = smartsheet.SheetResources.GetSheet(parsedSheetId3, null, null, null, null, null, null, null);
            Sheet sheet4 = smartsheet.SheetResources.GetSheet(parsedSheetId4, null, null, null, null, null, null, null);
            Sheet sheet5 = smartsheet.SheetResources.GetSheet(parsedSheetId5, null, null, null, null, null, null, null);
            Sheet sheet6 = smartsheet.SheetResources.GetSheet(parsedSheetId6, null, null, null, null, null, null, null);
            Sheet sheet7 = smartsheet.SheetResources.GetSheet(parsedSheetId7, null, null, null, null, null, null, null);

            StringBuilder addedBrandsData = new StringBuilder();
            StringBuilder addedInviteesData = new StringBuilder();
            StringBuilder addedHcpData = new StringBuilder();
            StringBuilder addedSlideKitData = new StringBuilder();
            StringBuilder addedExpences = new StringBuilder();

            int addedSlideKitDataNo = 1;
            int addedHcpDataNo = 1;
            int addedInviteesDataNo = 1;
            int addedBrandsDataNo = 1;
            int addedExpencesNo = 1;

            var TotalHonorariumAmount = 0;
            var TotalTravelAmount = 0;
            var TotalAccomodateAmount = 0;
            var TotalHCPLcAmount = 0;
            var TotalInviteesLcAmount = 0;
            var TotalExpenseAmount = 0;

            CultureInfo hindi = new CultureInfo("hi-IN");





            foreach (var formdata in formDataList.EventRequestExpenseSheet)
            {
                string rowData = $"{addedExpencesNo}. {formdata.Expense} | AmountExcludingTax: {formdata.AmountExcludingTax}| Amount: {formdata.Amount} | {formdata.BtcorBte}";
                addedExpences.AppendLine(rowData);
                addedExpencesNo++;
                var amount = int.Parse(formdata.Amount);
                TotalExpenseAmount = TotalExpenseAmount + amount;

            }
            string Expense = addedExpences.ToString();

            foreach (var formdata in formDataList.EventRequestHCPSlideKits)
            {

                string rowData = $"{addedSlideKitDataNo}. {formdata.MIS} | {formdata.SlideKitType}";
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




                // string rowData = $"{addedInviteesDataNo}. Name: {formdata.InviteeName} | MIS Code: {formdata.MISCode} | LocalConveyance: {formdata.LocalConveyance} ";
                string rowData = $"{addedInviteesDataNo}. {formdata.InviteeName}";
                addedInviteesData.AppendLine(rowData);
                addedInviteesDataNo++;
                TotalInviteesLcAmount = TotalInviteesLcAmount + int.Parse(formdata.LcAmount);
            }
            string Invitees = addedInviteesData.ToString();


            foreach (var formdata in formDataList.EventRequestHcpRole)
            {

                var HM = int.Parse(formdata.HonarariumAmount);
                var x = string.Format(hindi, "{0:#,#}", HM);
                var t = int.Parse(formdata.Travel) + int.Parse(formdata.Accomdation);
                var y = string.Format(hindi, "{0:#,#}", t);
                //string rowData = $"{addedHcpDataNo}. {formdata.HcpRole} |Name: {formdata.HcpName} | Honr.Amt: {formdata.HonarariumAmount} |Trav.Amt: {formdata.Travel} |Acco.Amt: {formdata.Accomdation} ";
                string rowData = $"{addedHcpDataNo}. {formdata.HcpRole} |{formdata.HcpName} | Honr.Amt: {x} |Trav.&Acc.Amt: {y} ";

                addedHcpData.AppendLine(rowData);
                addedHcpDataNo++;
                TotalHonorariumAmount = TotalHonorariumAmount + int.Parse(formdata.HonarariumAmount);
                TotalTravelAmount = TotalTravelAmount + int.Parse(formdata.Travel);
                TotalAccomodateAmount = TotalAccomodateAmount + int.Parse(formdata.Accomdation);
                TotalHCPLcAmount = TotalHCPLcAmount + int.Parse(formdata.LocalConveyance);
            }
            string HCP = addedHcpData.ToString();



            var FormattedTotalHonorariumAmount = string.Format(hindi, "{0:#,#}", TotalHonorariumAmount);
            var FormattedTotalTravelAmount = string.Format(hindi, "{0:#,#}", TotalTravelAmount);
            var FormattedTotalAccomodateAmount = string.Format(hindi, "{0:#,#}", TotalAccomodateAmount);
            var FormattedTotalHCPLcAmount = string.Format(hindi, "{0:#,#}", TotalHCPLcAmount);
            var FornattedTotalInviteesLcAmount = string.Format(hindi, "{0:#,#}", TotalInviteesLcAmount);
            var FormattedTotalExpenseAmount = string.Format(hindi, "{0:#,#}", TotalExpenseAmount);
            var c = TotalHCPLcAmount + TotalInviteesLcAmount;
            var FormattedTotalLC = string.Format(hindi, "{0:#,#}", c);
            var total = TotalHonorariumAmount + TotalTravelAmount + TotalAccomodateAmount + TotalHCPLcAmount + TotalInviteesLcAmount + TotalExpenseAmount;

            var FormattedTotal = string.Format(hindi, "{0:#,#}", total);
            var s = (TotalTravelAmount + TotalAccomodateAmount);
            var FormattedTotalTAAmount = string.Format(hindi, "{0:#,#}", s);




            try
            {

                var newRow = new Row();
                newRow.Cells = new List<Cell>();
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.Webinar.EventTopic });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "EventType"), Value = formDataList.Webinar.EventType });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "EventDate"), Value = formDataList.Webinar.EventDate });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "StartTime"), Value = formDataList.Webinar.StartTime });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "EndTime"), Value = formDataList.Webinar.EndTime });

                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Meeting Type"), Value = formDataList.Webinar.Meeting_Type });

                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Brands"), Value = brand });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Expenses"), Value = Expense });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Panelists"), Value = HCP });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Invitees"), Value = Invitees });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "SlideKits"), Value = slideKit });
                //newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Role"), Value = formDataList.Webinar.Role });


                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "IsAdvanceRequired"), Value = formDataList.Webinar.IsAdvanceRequired });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "EventOpen30days"), Value = formDataList.Webinar.EventOpen30days });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "EventWithin7days"), Value = formDataList.Webinar.EventWithin7days });
                //newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "RBM/BM"), Value = formDataList.Webinar.RBMorBM });
                //newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Sales Head"), Value = formDataList.Webinar.Sales_Head });
                //newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Marketing Head"), Value = formDataList.Webinar.Marketing_Head });
                //newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Finance Treasury"), Value = formDataList.Webinar.Finance });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "InitiatorName"), Value = formDataList.Webinar.InitiatorName });
                //newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Initiator Email"), Value = formDataList.Webinar.Initiator_Email });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Advance Amount"), Value = formDataList.Webinar.AdvanceAmount });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, " Total Expense BTC"), Value = formDataList.Webinar.TotalExpenseBTC });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Total Expense BTE"), Value = formDataList.Webinar.TotalExpenseBTE });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Total Honorarium Amount"), Value = TotalHonorariumAmount });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Total Travel Amount"), Value = TotalTravelAmount });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Total Travel & Accommodation Amount"), Value = s });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Total Accomodation Amount"), Value = TotalAccomodateAmount });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Total Budget"), Value = total });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Total Local Conveyance"), Value = c });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Total Expense"), Value = TotalExpenseAmount });






                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Initiator Email"), Value = formDataList.Webinar.Initiator_Email });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "RBM/BM"), Value = formDataList.Webinar.RBMorBM });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Sales Head"), Value = formDataList.Webinar.Sales_Head });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Sales Coordinator"), Value = formDataList.Webinar.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Marketing Head"), Value = formDataList.Webinar.Marketing_Head });
                //newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet, "Finance Treasury"), Value = formData.RequestHonorariumList.MarketingHeadEmail });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Compliance"), Value = formDataList.Webinar.ComplianceEmail });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Finance Accounts"), Value = formDataList.Webinar.FinanceAccountsEmail });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Finance Treasury"), Value = formDataList.Webinar.Finance });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Reporting Manager"), Value = formDataList.Webinar.ReportingManagerEmail });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "1 Up Manager"), Value = formDataList.Webinar.FirstLevelEmail });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Medical Affairs Head"), Value = formDataList.Webinar.MedicalAffairsEmail });
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet1, "Role"), Value = formDataList.Webinar.Role });


                var addedRows = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId1, new Row[] { newRow });

                var eventIdColumnId = GetColumnIdByName(sheet1, "EventId/EventRequestId");
                var eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                var val = eventIdCell.DisplayValue;



                var x = 1;
                foreach (var p in formDataList.Webinar.Files)
                {
                    string[] words = p.Split(':');
                    var r = words[0];
                    var q = words[1];


                    byte[] fileBytes = Convert.FromBase64String(q);
                    var folderName = Path.Combine("Resources", "Images");
                    var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                    if (!Directory.Exists(pathToSave))
                    {
                        Directory.CreateDirectory(pathToSave);
                    }

                    string fileType = GetFileType(fileBytes);
                    string fileName = r ;
                    // string fileName = val+x + ": AttachedFile." + fileType;
                    string filePath = Path.Combine(pathToSave, fileName);


                    var addedRow = addedRows[0];

                    System.IO.File.WriteAllBytes(filePath, fileBytes);
                    string type = GetContentType(fileType);
                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                            parsedSheetId1, addedRow.Id.Value, filePath, "application/msword");
                    x++;

                    if (System.IO.File.Exists(filePath))
                    {
                        System.IO.File.Delete(filePath);
                    }
                }




                if (formDataList.Webinar.EventOpen30days == "Yes" || formDataList.Webinar.EventWithin7days == "Yes" || formDataList.Webinar.FB_Expense_Excluding_Tax == "Yes")
                {
                    var eventId = val;
                    try
                    {

                        var newRow7 = new Row();
                        newRow7.Cells = new List<Cell>();

                        newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                        newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.Webinar.EventTopic });

                        newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "EventType"), Value = formDataList.Webinar.EventType });
                        newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "EventDate"), Value = formDataList.Webinar.EventDate });
                        newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "StartTime"), Value = formDataList.Webinar.StartTime });
                        newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "EndTime"), Value = formDataList.Webinar.EndTime });
                        //newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "VenueName"), Value = formDataList.class1.VenueName });
                        //newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "City"), Value = formDataList.class1.City });
                        //newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "State"), Value = formDataList.class1.State });

                        newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "EventOpen30days"), Value = formDataList.Webinar.EventOpen30days });
                        newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "EventWithin7days"), Value = formDataList.Webinar.EventWithin7days });
                        newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "PRE-F&B Expense Excluding Tax"), Value = formDataList.Webinar.FB_Expense_Excluding_Tax });
                        newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.Webinar.Sales_Head });
                        newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "Finance Head"), Value = formDataList.Webinar.Sales_Head });

                        newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "InitiatorName"), Value = formDataList.Webinar.InitiatorName });
                        newRow7.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.Webinar.Initiator_Email });


                        var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId7, new Row[] { newRow7 });




                        var j = 1;
                        foreach (var p in formDataList.Webinar.DeviationFiles)
                        {

                            string[] words = p.Split(':');
                            var r = words[0];
                            var q = words[1];

                            byte[] fileBytes = Convert.FromBase64String(q);
                            var folderName = Path.Combine("Resources", "Images");
                            var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                            if (!Directory.Exists(pathToSave))
                            {
                                Directory.CreateDirectory(pathToSave);
                            }

                            string fileType = GetFileType(fileBytes);
                            string fileName =r ;
                            // string fileName = val+x + ": AttachedFile." + fileType;
                            string filePath = Path.Combine(pathToSave, fileName);


                            var addedRow = addeddeviationrow[0];

                            System.IO.File.WriteAllBytes(filePath, fileBytes);
                            string type = GetContentType(fileType);
                            var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                    parsedSheetId7, addedRow.Id.Value, filePath, "application/msword");
                            j++;
                            if (System.IO.File.Exists(filePath))
                            {
                                System.IO.File.Delete(filePath);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        return BadRequest(ex.Message);
                    }
                }




































                foreach (var formData in formDataList.EventRequestHcpRole)
                {
                    var newRow1 = new Row();
                    newRow1.Cells = new List<Cell>();
                    newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "HcpRole"), Value = formData.HcpRole });

                    newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "MISCode"), Value = formData.MisCode });
                    newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "Travel"), Value = formData.Travel });
                    newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "TotalSpend"), Value = formData.FinalAmount });
                    newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "Accomodation"), Value = formData.Accomdation });
                    newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "LocalConveyance"), Value = formData.LocalConveyance });
                    newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "SpeakerCode"), Value = formData.SpeakerCode });
                    newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "TrainerCode"), Value = formData.TrainerCode });
                    newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "HonorariumRequired"), Value = formData.HonorariumRequired });

                    newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "Speciality"), Value = formData.Speciality });
                    if (formData.HcpRole == "Speaker")
                    {
                        newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "HCPName"), Value = formData.HcpName });
                        newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "Event Topic"), Value = formDataList.Webinar.EventTopic });
                        newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "Event Type"), Value = formDataList.Webinar.EventType });
                        //newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "Venue name"), Value = formDataList.Webinar.VenueName });
                        newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "Event Date Start"), Value = formDataList.Webinar.EventDate });
                        newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "Event End Date"), Value = formDataList.Webinar.EventDate });
                        newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "HonorariumAmount"), Value = formData.HonarariumAmount });
                        newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "PAN card name"), Value = formData.HcpName });
                    }
                    if (formData.HcpRole == "Trainer")
                    {
                        newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "Event Topic"), Value = formDataList.Webinar.EventTopic });
                        newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "Event Type"), Value = formDataList.Webinar.EventType });
                        //newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "Venue name"), Value = formDataList.Webinar.VenueName });
                        newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "Event Date Start"), Value = formDataList.Webinar.EventDate });
                        newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "Event End Date"), Value = formDataList.Webinar.EventDate });
                        newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "HCPName"), Value = formData.HcpName });
                        newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "Event Start Time"), Value = formDataList.Webinar.StartTime });

                    }
                    newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "Tier"), Value = formData.Tier });
                    newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "HCP Type"), Value = formData.GOorNGO });
                    newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "PresentationDuration"), Value = formData.PresentationDuration });
                    newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "PanelSessionPreparationDuration"), Value = formData.PanelSessionPreperationDuration });
                    newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "PanelDiscussionDuration"), Value = formData.PanelDisscussionDuration });
                    newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "QASessionDuration"), Value = formData.QASessionDuration });
                    newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "BriefingSession"), Value = formData.BriefingSession });
                    newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "TotalSessionHours"), Value = formData.TotalSessionHours });
                    newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "Rationale"), Value = formData.Rationale });
                    newRow1.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet4, "EventId/EventRequestId"), Value = val });


                    smartsheet.SheetResources.RowResources.AddRows(parsedSheetId4, new Row[] { newRow1 });







                }

                foreach (var formdata in formDataList.RequestBrandsList)
                {
                    var newRow2 = new Row();
                    newRow2.Cells = new List<Cell>();
                    newRow2.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet2, "% Allocation"), Value = formdata.PercentAllocation });
                    newRow2.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet2, "Brands"), Value = formdata.BrandName });
                    newRow2.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet2, "Project ID"), Value = formdata.ProjectId });
                    newRow2.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet2, "EventId/EventRequestId"), Value = val });


                    smartsheet.SheetResources.RowResources.AddRows(parsedSheetId2, new Row[] { newRow2 });

                }
                foreach (var formdata in formDataList.EventRequestInvitees)
                {
                    var newRow3 = new Row();
                    newRow3.Cells = new List<Cell>();
                    newRow3.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet3, "HCPName"), Value = formdata.InviteeName });

                    newRow3.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet3, "LocalConveyance"), Value = formdata.LocalConveyance });
                    newRow3.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet3, "BTC/BTE"), Value = formdata.BtcorBte });
                    newRow3.Cells.Add(new Cell
                    { ColumnId = GetColumnIdByName(sheet3, "LcAmount"), Value = formdata.LcAmount });
                    newRow3.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet3, "EventId/EventRequestId"), Value = val });
                    newRow3.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet3, "Invitee Source"), Value = formdata.InviteedFrom });
                    if (formdata.InviteedFrom == "Others")
                    {
                        newRow3.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet3, "Speciality"), Value = formdata.Speciality });
                        newRow3.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet3, "HCP Type"), Value = formdata.HCPType });
                    }
                    else
                    {
                        newRow3.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet3, "MISCode"), Value = formdata.MISCode });
                    }


                    smartsheet.SheetResources.RowResources.AddRows(parsedSheetId3, new Row[] { newRow3 });
                }


                foreach (var formdata in formDataList.EventRequestHCPSlideKits)
                {
                    var newRow5 = new Row();
                    newRow5.Cells = new List<Cell>();

                    newRow5.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet5, "MIS"), Value = formdata.MIS });
                    newRow5.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet5, "Slide Kit Type"), Value = formdata.SlideKitType });
                    newRow5.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet5, "SlideKit Document"), Value = formdata.SlideKitDocument });
                    newRow5.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet5, "EventId/EventRequestId"), Value = val });


                    smartsheet.SheetResources.RowResources.AddRows(parsedSheetId5, new Row[] { newRow5 });
                }

                foreach (var formdata in formDataList.EventRequestExpenseSheet)
                {
                    var newRow6 = new Row();
                    newRow6.Cells = new List<Cell>();

                    newRow6.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet6, "Expense"), Value = formdata.Expense });
                    newRow6.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet6, "EventId/EventRequestID"), Value = val });
                    newRow6.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet6, "AmountExcludingTax?"), Value = formdata.AmountExcludingTax });
                    newRow6.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet6, "Amount"), Value = formdata.Amount });
                    newRow6.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet6, "BTC/BTE"), Value = formdata.BtcorBte });
                    newRow6.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet6, "BudgetAmount"), Value = formdata.BudgetAmount });
                    newRow6.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet6, "BTCAmount"), Value = formdata.BtcAmount });
                    newRow6.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet6, "BTEAmount"), Value = formdata.BteAmount });

                    smartsheet.SheetResources.RowResources.AddRows(parsedSheetId6, new Row[] { newRow6 });
                }

                return Ok(new
                { Message = " Success!" });



            }



            catch (Exception ex)
            {
                return BadRequest($"Could not find {ex.Message}");
            }












        }

        private string GetContentType(string fileExtension)
        {
            switch (fileExtension.ToLower())
            {
                case "jpg":
                case "jpeg":
                    return "image/jpeg";
                case "pdf":
                    return "application/pdf";
                case "gif":
                    return "image/gif";
                case "png":
                    return "image/png";
                case "webp":
                    return "image/webp";
                case "doc":
                    return "application/msword";
                case "docx":
                    return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                default:
                    return "application/octet-stream";
            }
        }

        private string GetFileType(byte[] bytes)
        {

            if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xD8)
            {
                return "jpg";
            }
            else if (bytes.Length >= 4 && Encoding.UTF8.GetString(bytes, 0, 4) == "%PDF")
            {
                return "pdf";
            }
            else if (bytes.Length >= 3 && Encoding.UTF8.GetString(bytes, 0, 3) == "GIF")
            {
                return "gif";
            }
            else if (bytes.Length >= 8 && Encoding.UTF8.GetString(bytes, 0, 8) == "PNG\r\n\x1A\n")
            {
                return "png";
            }
            else if (bytes.Length >= 4 && Encoding.UTF8.GetString(bytes, 0, 4) == "RIFF" && Encoding.UTF8.GetString(bytes, 8, 4) == "WEBP")
            {
                return "webp";
            }
            else if (bytes.Length >= 4 && (bytes[0] == 0xD0 && bytes[1] == 0xCF && bytes[2] == 0x11 && bytes[3] == 0xE0))
            {
                return "doc"; // .doc format
            }
            else if (bytes.Length >= 4 && (bytes[0] == 0x50 && bytes[1] == 0x4B && bytes[2] == 0x03 && bytes[3] == 0x04))
            {
                return "docx"; // .docx format
            }
            else
            {
                return "unknown";
            }
        }




        private long GetColumnIdByName(Sheet sheet, string columnname)
        {
            foreach (var column in sheet.Columns)
            {
                if (column.Title == columnname)
                {
                    return column.Id.Value;
                }
            }
            return 0;
        }

    }
}
