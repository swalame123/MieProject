using IndiaEvents.Models.Models.EventTypeSheets;
using IndiaEvents.Models.Models.RequestSheets;
using IndiaEventsWebApi.Helper;
using IndiaEventsWebApi.Models;
using IndiaEventsWebApi.Models.EventTypeSheets;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using MySqlConnector;
using NPOI.Util;
using Org.BouncyCastle.Crypto.Tls;
using Serilog;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Data;
using System.Globalization;
using System.Text;


namespace IndiaEventsWebApi.Controllers.EventsController
{
    [Route("api/[controller]")]
    [ApiController]
    public class AllPreEventsController : ControllerBase
    {

        private readonly string accessToken;
        private readonly IConfiguration configuration;
        private readonly SemaphoreSlim _externalApiSemaphore;
        //private readonly SmartsheetClient smartsheet;
        private readonly string sheetId1;
        private readonly string sheetId2;
        private readonly string sheetId3;
        private readonly string sheetId4;
        private readonly string sheetId5;
        private readonly string sheetId6;
        private readonly string sheetId7;
        private readonly string sheetId8;
        private readonly string sheetId9;
        private readonly string UI_URL;

        public AllPreEventsController(IConfiguration configuration, SemaphoreSlim externalApiSemaphore)
        {

            this.configuration = configuration;
            this._externalApiSemaphore = externalApiSemaphore;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;


            sheetId1 = configuration.GetSection("SmartsheetSettings:Class1").Value;
            sheetId2 = configuration.GetSection("SmartsheetSettings:EventRequestBrandsList").Value;
            sheetId3 = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
            sheetId4 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
            sheetId5 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpSlideKit").Value;
            sheetId6 = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
            sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;
            sheetId8 = configuration.GetSection("SmartsheetSettings:EventRequestBeneficiary").Value;
            sheetId9 = configuration.GetSection("SmartsheetSettings:EventRequestProductBrandsList").Value;
            UI_URL = configuration.GetSection("SmartsheetSettings:UI_URL").Value;
        }
        //private static SemaphoreSlim semaphore;


        [HttpPost("Class1PreEventSmartSheet"), DisableRequestSizeLimit]
        public async Task<IActionResult> Class1PreEventSmartSheet(AllObjModels formDataList)
        {
            try
            {


                SmartsheetClient smartsheet = await Task.Run(() => SmartSheetBuilder.AccessClient(accessToken, _externalApiSemaphore));

                Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
                Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
                Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
                Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
                Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
                Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
                Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);
                Sheet UrlData = SheetHelper.GetSheetById(smartsheet, UI_URL);

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
                //int addedExpencesNoBTE = 1;

                double TotalHonorariumAmount = 0;
                double TotalTravelAmount = 0;
                double TotalAccomodateAmount = 0;
                double TotalHCPLcAmount = 0;
                double TotalInviteesLcAmount = 0;
                double TotalExpenseAmount = 0;

                foreach (var formdata in formDataList.EventRequestExpenseSheet)
                {
                    string rowData = $"{addedExpencesNo}. {formdata.Expense} | AmountExcludingTax: {formdata.AmountExcludingTax}| Amount: {formdata.Amount} | {formdata.BtcorBte}";
                    addedExpences.AppendLine(rowData);
                    addedExpencesNo++;
                    var amount = SheetHelper.NumCheck(formdata.Amount);
                    TotalExpenseAmount = TotalExpenseAmount + amount;
                }
                string Expense = addedExpences.ToString();

                StringBuilder addedExpencesBTE = new();
                int addedExpencesNoBTE = 1;
                foreach (var formdata in formDataList.EventRequestExpenseSheet)
                {
                    if (formdata.BtcorBte.ToLower() == "bte")
                    {
                        string rowData = $"{addedExpencesNoBTE}. {formdata.Expense} | Amount: {formdata.Amount}";
                        addedExpencesBTE.AppendLine(rowData);
                        addedExpencesNoBTE++;
                    }
                }
                string BTEExpense = addedExpencesBTE.ToString();

                foreach (var formdata in formDataList.EventRequestHCPSlideKits)
                {
                    string rowData = $"{addedSlideKitDataNo}. {formdata.HcpName} | {formdata.SlideKitType} | Id :{formdata.SlideKitDocument}";
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
                    double HM = SheetHelper.NumCheck(formdata.HonarariumAmount);
                    double t = SheetHelper.NumCheck(formdata.Travel) + SheetHelper.NumCheck(formdata.Accomdation);
                    double roundedValue = Math.Round(t, 2);
                    string rowData = $"{addedHcpDataNo}. {formdata.HcpRole} |{formdata.HcpName} | Honr.Amt: {HM} |Trav.&Acc.Amt: {roundedValue} |Rationale :{formdata.Rationale}";
                    addedHcpData.AppendLine(rowData);
                    addedHcpDataNo++;
                    TotalTravelAmount += SheetHelper.NumCheck(formdata.Travel);
                    TotalHonorariumAmount += SheetHelper.NumCheck(formdata.HonarariumAmount);
                    TotalAccomodateAmount += SheetHelper.NumCheck(formdata.Accomdation);
                    TotalHCPLcAmount += SheetHelper.NumCheck(formdata.LocalConveyance);
                }
                string HCP = addedHcpData.ToString();
                double cc = TotalHCPLcAmount + TotalInviteesLcAmount;

                double totalAmount = TotalHonorariumAmount + TotalTravelAmount + TotalAccomodateAmount + TotalHCPLcAmount + TotalInviteesLcAmount + TotalExpenseAmount;

                double ss = TotalTravelAmount + TotalAccomodateAmount;

                double c = Math.Round(cc, 2);
                double total = Math.Round(totalAmount, 2);
                double s = Math.Round(ss, 2);
                Dictionary<string, long> Sheet1columns = new();
                foreach (var column in sheet1.Columns)
                {
                    Sheet1columns.Add(column.Title, (long)column.Id);
                }
                try
                {//public string? BTEExpenseDetails { get; set; }
                    Row newRow = new()
                    {
                        Cells = new List<Cell>()
                    };

                    Row? targetRow1 = UrlData.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == "Approver Pre Event URL"));
                    Row? targetRow2 = UrlData.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == "Finance Treasury URL"));
                    Row? targetRow4 = UrlData.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == "Initiator URL"));

                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Approver Pre Event URL"], Value = targetRow1?.Cells[1].Value ?? "no url" });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Finance Treasury URL"], Value = targetRow2?.Cells[1].Value ?? "no url" });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Initiator URL"], Value = targetRow4?.Cells[1].Value ?? "no url" });

                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Event Topic"], Value = formDataList.class1.EventTopic });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Start Time"], Value = formDataList.class1.StartTime });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["End Time"], Value = formDataList.class1.EndTime });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Venue Name"], Value = formDataList.class1.VenueName });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["City"], Value = formDataList.class1.City });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["State"], Value = formDataList.class1.State });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Panelists"], Value = HCP });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Invitees"], Value = Invitees });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["MIPL Invitees"], Value = MenariniInvitees });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Event Type"], Value = formDataList.class1.EventType });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Event Date"], Value = formDataList.class1.EventDate });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Brands"], Value = brand });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Expenses"], Value = Expense });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["BTE Expense Details"], Value = formDataList.class1.BTEExpenseDetails });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["SlideKits"], Value = slideKit });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["IsAdvanceRequired"], Value = formDataList.class1.IsAdvanceRequired });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["EventOpen30days"], Value = formDataList.class1.EventOpen30days });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["EventWithin7days"], Value = formDataList.class1.EventWithin7days });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Initiator Name"], Value = formDataList.class1.InitiatorName });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Advance Amount"], Value = SheetHelper.NumCheck(formDataList.class1.AdvanceAmount) });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns[" Total Expense BTC"], Value = SheetHelper.NumCheck(formDataList.class1.TotalExpenseBTC) });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Total Expense BTE"], Value = SheetHelper.NumCheck(formDataList.class1.TotalExpenseBTE) });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Total Honorarium Amount"], Value = Math.Round(TotalHonorariumAmount, 2) });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Total Travel Amount"], Value = Math.Round(TotalTravelAmount, 2) });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Total Travel & Accommodation Amount"], Value = s });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Total Accommodation Amount"], Value = Math.Round(TotalAccomodateAmount, 2) });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Budget Amount"], Value = total });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Total Local Conveyance"], Value = c });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Total Expense"], Value = Math.Round(TotalExpenseAmount, 2) });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Initiator Email"], Value = formDataList.class1.Initiator_Email });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["RBM/BM"], Value = formDataList.class1.RBMorBM });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Sales Head"], Value = formDataList.class1.Sales_Head });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Sales Coordinator"], Value = formDataList.class1.SalesCoordinatorEmail });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Marketing Coordinator"], Value = formDataList.class1.MarketingCoordinatorEmail });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Marketing Head"], Value = formDataList.class1.Marketing_Head });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Finance Treasury"], Value = formDataList.class1.Finance });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Compliance"], Value = formDataList.class1.ComplianceEmail });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Finance Accounts"], Value = formDataList.class1.FinanceAccountsEmail });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Reporting Manager"], Value = formDataList.class1.ReportingManagerEmail });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["1 Up Manager"], Value = formDataList.class1.FirstLevelEmail });
                    newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Medical Affairs Head"], Value = formDataList.class1.MedicalAffairsEmail });

                    // IList<Row> addedRows = smartsheet.SheetResources.RowResources.AddRows(sheet1.Id.Value, new Row[] { newRow });
                    IList<Row> addedRows = ApiCalls.WebDetails(smartsheet, sheet1, newRow);

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
                        string filePath = SheetHelper.testingFile(q, name);
                        Row addedRow = addedRows[0];
                        // Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                        Attachment attachment = await ApiCalls.AddAttachmentsToSheet(smartsheet, sheet1, addedRow, filePath);
                        x++;
                        if (System.IO.File.Exists(filePath))
                        {
                            SheetHelper.DeleteFile(filePath);
                        }
                    }


                    Dictionary<string, long> Sheet4columns = new();
                    foreach (var column in sheet4.Columns)
                    {
                        Sheet4columns.Add(column.Title, (long)column.Id);
                    }



                    foreach (var formData in formDataList.EventRequestHcpRole)
                    {
                        Row newRow1 = new()
                        {
                            Cells = new List<Cell>()
                        };
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["HcpRole"], Value = formData.HcpRole });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["MISCode"], Value = SheetHelper.MisCodeCheck(formData.MisCode) });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Travel"], Value = SheetHelper.NumCheck(formData.Travel) });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["TotalSpend"], Value = SheetHelper.NumCheck(formData.FinalAmount) });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Accomodation"], Value = SheetHelper.NumCheck(formData.Accomdation) });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["LocalConveyance"], Value = SheetHelper.NumCheck(formData.LocalConveyance) });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["SpeakerCode"], Value = formData.SpeakerCode });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["TrainerCode"], Value = formData.TrainerCode });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["HonorariumRequired"], Value = formData.HonorariumRequired });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["AgreementAmount"], Value = SheetHelper.NumCheck(formData.AgreementAmount) });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["HonorariumAmount"], Value = SheetHelper.NumCheck(formData.HonarariumAmount) });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Speciality"], Value = formData.Speciality });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Event Topic"], Value = formDataList.class1.EventTopic });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Event Type"], Value = formDataList.class1.EventType });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Venue name"], Value = formDataList.class1.VenueName });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Event Date Start"], Value = formDataList.class1.EventDate });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Event End Date"], Value = formDataList.class1.EventEndDate });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["HCPName"], Value = formData.HcpName });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["PAN card name"], Value = formData.PanCardName });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["ExpenseType"], Value = formData.ExpenseType });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Bank Account Number"], Value = formData.BankAccountNumber });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Bank Name"], Value = formData.BankName });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["IFSC Code"], Value = formData.IFSCCode });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["FCPA Date"], Value = formData.Fcpadate });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Currency"], Value = formData.Currency });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Honorarium Amount Excluding Tax"], Value = formData.HonarariumAmountExcludingTax });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Travel Excluding Tax"], Value = formData.TravelExcludingTax });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Accomodation Excluding Tax"], Value = formData.AccomdationExcludingTax });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Local Conveyance Excluding Tax"], Value = formData.LocalConveyanceExcludingTax });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["LC BTC/BTE"], Value = formData.LcBtcorBte });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Travel BTC/BTE"], Value = formData.TravelBtcorBte });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Mode of Travel"], Value = formData.TravelSelection });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Accomodation BTC/BTE"], Value = formData.AccomodationBtcorBte });

                        if (formData.Currency == "Others")
                        {
                            newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Other Currency"], Value = formData.OtherCurrencyType });
                        }
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Beneficiary Name"], Value = formData.BeneficiaryName });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Pan Number"], Value = formData.PanNumber });

                        if (formData.HcpRole == "Others")
                        {

                            newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Other Type"], Value = formData.OthersType });
                        }

                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Tier"], Value = formData.Tier });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["HCP Type"], Value = formData.GOorNGO });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["PresentationDuration"], Value = SheetHelper.NumCheck(formData.PresentationDuration) });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["PanelSessionPreparationDuration"], Value = SheetHelper.NumCheck(formData.PanelSessionPreperationDuration) });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["PanelDiscussionDuration"], Value = SheetHelper.NumCheck(formData.PanelDisscussionDuration) });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["QASessionDuration"], Value = SheetHelper.NumCheck(formData.QASessionDuration) });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["BriefingSession"], Value = SheetHelper.NumCheck(formData.BriefingSession) });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["TotalSessionHours"], Value = SheetHelper.NumCheck(formData.TotalSessionHours) });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Rationale"], Value = formData.Rationale });
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["EventId/EventRequestId"], Value = val });


                        //  IList<Row> row = smartsheet.SheetResources.RowResources.AddRows(sheet4.Id.Value, new Row[] { newRow1 });
                        IList<Row> row = await Task.Run(() => ApiCalls.PanelDetails(smartsheet, sheet4, newRow1));
                        if (formData.IsUpload == "Yes")
                        {
                            foreach (string p in formData.FilesToUpload)
                            {
                                string[] words = p.Split(':');
                                string r = words[0];
                                string q = words[1];
                                string name = r.Split(".")[0];
                                string filePath = SheetHelper.testingFile(q, name);
                                Row addedRow = row[0];
                                //Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                //       sheet4.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                                Attachment attachment = await ApiCalls.AddAttachmentsToSheet(smartsheet, sheet4, addedRow, filePath);

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
                    await Task.Run(() => ApiCalls.BrandsDetails(smartsheet, sheet2, newRows2));

                    //smartsheet.SheetResources.RowResources.AddRows(sheet2.Id.Value, newRows2.ToArray());
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
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "MISCode"), Value = SheetHelper.MisCodeCheck(formdata.MISCode )},
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Topic"), Value = formDataList.class1.EventTopic },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Type"), Value = formDataList.class1.EventType },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Venue name"), Value = formDataList.class1.VenueName },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Date Start"), Value = formDataList.class1.EventDate },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event End Date"), Value = formDataList.class1.EventDate }

                        }
                        };
                        newRows3.Add(newRow3);
                    }

                    await Task.Run(() => ApiCalls.InviteesDetails(smartsheet, sheet3, newRows3));
                    // smartsheet.SheetResources.RowResources.AddRows(sheet3.Id.Value, newRows3.ToArray());
                    foreach (var formdata in formDataList.EventRequestHCPSlideKits)
                    {
                        Row newRow5 = new()
                        {
                            Cells = new List<Cell>()
                        };

                        newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "MIS"), Value = SheetHelper.MisCodeCheck(formdata.MIS) });
                        newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "Slide Kit Type"), Value = formdata.SlideKitType });
                        newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "SlideKit Document"), Value = formdata.SlideKitDocument });
                        newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "EventId/EventRequestId"), Value = val });

                        IList<Row> row = await Task.Run(() => ApiCalls.SlideKitDetails(smartsheet, sheet5, newRow5));

                        //IList<Row> row = smartsheet.SheetResources.RowResources.AddRows(sheet5.Id.Value, new Row[] { newRow5 });
                        if (formdata.IsUpload == "Yes")
                        {
                            foreach (string p in formdata.FilesToUpload)
                            {
                                string[] words = p.Split(':');
                                string r = words[0];
                                string q = words[1];
                                string name = r.Split(".")[0];
                                string filePath = SheetHelper.testingFile(q, name);
                                Row addedRow = row[0];
                                //Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                //       sheet5.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                                Attachment attachment = await ApiCalls.AddAttachmentsToSheet(smartsheet, sheet5, addedRow, filePath);


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
                    //smartsheet.SheetResources.RowResources.AddRows(sheet6.Id.Value, newRows6.ToArray());
                    await Task.Run(() => ApiCalls.ExpenseDetails(smartsheet, sheet6, newRows6));

                    if (formDataList.class1.EventOpen30days == "Yes" || formDataList.class1.EventWithin7days == "Yes" || formDataList.class1.FB_Expense_Excluding_Tax == "Yes" || formDataList.class1.IsDeviationUpload == "Yes")
                    {
                        List<string> DeviationNames = new List<string>();
                        foreach (var p in formDataList.class1.DeviationDetails)
                        {

                            string[] words = p.DeviationFile.Split(':')[0].Split("*");
                            string r = words[1];
                            DeviationNames.Add(r);
                        }
                        foreach (var pp in formDataList.class1.DeviationDetails)
                        {
                            foreach (var deviationname in DeviationNames)
                            {
                                string file = deviationname.Split(".")[0];
                                string eventId = val;
                                if (pp.DeviationFile.Split(':')[0].Split("*")[1] == deviationname)
                                {
                                    try
                                    {
                                        Row newRow7 = new()
                                        {
                                            Cells = new List<Cell>()
                                        };
                                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.class1.EventTopic });
                                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Type"), Value = formDataList.class1.EventType });
                                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Date"), Value = formDataList.class1.EventDate });
                                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Start Time"), Value = formDataList.class1.StartTime });
                                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "End Time"), Value = formDataList.class1.EndTime });
                                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Venue Name"), Value = formDataList.class1.VenueName });
                                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "City"), Value = formDataList.class1.City });
                                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "State"), Value = formDataList.class1.State });

                                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "MIS Code"), Value = SheetHelper.MisCodeCheck(pp.MisCode) });
                                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HCP Name"), Value = pp.HcpName });
                                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Honorarium Amount"), Value = pp.HonorariumAmountExcludingTax });
                                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Travel & Accommodation Amount"), Value = pp.TravelorAccomodationAmountExcludingTax });
                                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Other Expenses"), Value = pp.OtherExpenseAmountExcludingTax });

                                        if (file == "30DaysDeviationFile")
                                        {
                                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:30DaysDeviationFile").Value });
                                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = formDataList.class1.EventOpen30days });
                                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = SheetHelper.NumCheck(formDataList.class1.EventOpen30dayscount) });
                                        }
                                        else if (file == "7DaysDeviationFile")
                                        {
                                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:5DaysDeviationFile").Value });
                                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = formDataList.class1.EventWithin7days });

                                        }
                                        else if (file == "ExpenseExcludingTax")
                                        {
                                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:ExpenseExcludingTax").Value });
                                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "PRE-F&B Expense Excluding Tax"), Value = formDataList.class1.FB_Expense_Excluding_Tax });
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

                                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.class1.Sales_Head });
                                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formDataList.class1.FinanceHead });
                                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Name"), Value = formDataList.class1.InitiatorName });
                                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.class1.Initiator_Email });
                                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Coordinator"), Value = formDataList.class1.SalesCoordinatorEmail });

                                        //IList<Row> addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });
                                        IList<Row> addeddeviationrow = ApiCalls.DeviationData(smartsheet, sheet7, newRow7);

                                        int j = 1;
                                        foreach (var p in formDataList.class1.DeviationDetails)
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
                                                //Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                                                //Attachment attachmentinmain = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet1.Id.Value, addedRows[0].Id.Value, filePath, "application/msword");

                                                Attachment attachment = await ApiCalls.AddAttachmentsToSheet(smartsheet, sheet7, addedRow, filePath);
                                                Attachment attachmentinmain = await ApiCalls.AddAttachmentsToSheet(smartsheet, sheet1, addedRows[0], filePath);




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


                    Row targetRow = addedRows[0];
                    long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role");
                    Cell cellToUpdateB = new() { ColumnId = honorariumSubmittedColumnId, Value = formDataList.class1.Role };
                    Row updateRow = new() { Id = targetRow.Id, Cells = new Cell[] { cellToUpdateB } };
                    Cell? cellToUpdate = targetRow.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                    if (cellToUpdate != null) { cellToUpdate.Value = formDataList.class1.Role; }

                    // smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRow });

                    await Task.Run(() => ApiCalls.UpdateRole(smartsheet, sheet1, updateRow));


                    //return Ok(new
                    //{ Message = " Success!" });
                    DateTime currentDate = DateTime.Now;
                    return Ok(new
                    {
                        Message = $"Thank you. Your event creation request has been received. " +
                    "You should receive a confirmation email with the details of your event after a few minutes."
                    });

                }
                catch (Exception ex)
                {
                    Log.Error($"Error occured on AllPreEventsController Attachementfile method {ex.Message} at {DateTime.Now}");
                    Log.Error(ex.StackTrace);
                    return BadRequest(ex.Message);
                }
            }
            catch (Exception ex)
            {

                return BadRequest(new
                {
                    Message = ex.Message + "------" + ex.StackTrace
                });
            }

        }

        [HttpPost("Class1PreEvent"), DisableRequestSizeLimit]
        public async Task<IActionResult> Class1PreEvent(AllObjModels formDataList)
        {

            string strMessage = string.Empty;

            //int timeInterval = 4000;
            //await Task.Delay(timeInterval);
            try
            {
                //semaphore = new SemaphoreSlim(0, 1);
                Log.Information("starting of api " + DateTime.Now);
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                Sheet UrlData = SheetHelper.GetSheetById(smartsheet, UI_URL);
                /////////////////////

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
                double TotalHonorariumAmount = 0;
                double TotalTravelAmount = 0;
                double TotalAccomodateAmount = 0;
                double TotalHCPLcAmount = 0;
                double TotalInviteesLcAmount = 0;
                double TotalExpenseAmount = 0;
                foreach (var formdata in formDataList.EventRequestExpenseSheet)
                {
                    string rowData = $"{addedExpencesNo}. {formdata.Expense} | AmountExcludingTax: {formdata.AmountExcludingTax}| Amount: {formdata.Amount} | {formdata.BtcorBte}";
                    addedExpences.AppendLine(rowData);
                    addedExpencesNo++;
                    var amount = SheetHelper.NumCheck(formdata.Amount);
                    TotalExpenseAmount = TotalExpenseAmount + amount;
                }
                string Expense = addedExpences.ToString();

                StringBuilder addedExpencesBTE = new();
                int addedExpencesNoBTE = 1;
                foreach (var formdata in formDataList.EventRequestExpenseSheet)
                {
                    if (formdata.BtcorBte.ToLower() == "bte")
                    {
                        string rowData = $"{addedExpencesNoBTE}. {formdata.Expense} | Amount: {formdata.Amount}";
                        addedExpencesBTE.AppendLine(rowData);
                        addedExpencesNoBTE++;
                    }
                }
                string BTEExpense = addedExpencesBTE.ToString();

                foreach (var formdata in formDataList.EventRequestHCPSlideKits)
                {
                    string rowData = $"{addedSlideKitDataNo}. {formdata.HcpName} | {formdata.SlideKitType} | Id :{formdata.SlideKitDocument}";
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
                    double HM = SheetHelper.NumCheck(formdata.HonarariumAmount);
                    double t = SheetHelper.NumCheck(formdata.Travel) + SheetHelper.NumCheck(formdata.Accomdation);

                    double roundedValue = Math.Round(t, 2);

                    string rowData = $"{addedHcpDataNo}. {formdata.HcpRole} |{formdata.HcpName} | Honr.Amt: {HM} |Trav.&Acc.Amt: {roundedValue} |Rationale : {formdata.Rationale}";
                    addedHcpData.AppendLine(rowData);
                    addedHcpDataNo++;
                    TotalHonorariumAmount = TotalHonorariumAmount + SheetHelper.NumCheck(formdata.HonarariumAmount);
                    TotalTravelAmount = TotalTravelAmount + SheetHelper.NumCheck(formdata.Travel);
                    TotalAccomodateAmount = TotalAccomodateAmount + SheetHelper.NumCheck(formdata.Accomdation);
                    TotalHCPLcAmount = TotalHCPLcAmount + SheetHelper.NumCheck(formdata.LocalConveyance);
                }
                string HCP = addedHcpData.ToString();


                double cc = TotalHCPLcAmount + TotalInviteesLcAmount;

                double totalAmount = TotalHonorariumAmount + TotalTravelAmount + TotalAccomodateAmount + TotalHCPLcAmount + TotalInviteesLcAmount + TotalExpenseAmount;

                double ss = TotalTravelAmount + TotalAccomodateAmount;

                double c = Math.Round(cc, 2);
                double total = Math.Round(totalAmount, 2);
                double s = Math.Round(ss, 2);

                Row? targetRow1 = UrlData.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == "Approver Pre Event URL"));
                Row? targetRow2 = UrlData.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == "Finance Treasury URL"));
                Row? targetRow4 = UrlData.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == "Initiator URL"));
                String Attachmentpaths = "";
                foreach (var p in formDataList.class1.Files)
                {
                    string[] words = p.Split(':');
                    string r = words[0];
                    string q = words[1];
                    string name = r.Split(".")[0];
                    string filePath = SheetHelper.SQlFileinsertion(q, name);
                    Attachmentpaths = Attachmentpaths + "," + filePath;
                }

                string MyConnection = configuration.GetSection("ConnectionStrings:mysql").Value;
                MySqlConnection MyConn = new MySqlConnection(MyConnection);
                MySqlCommand com = new MySqlCommand("Class1Preevent", MyConn);
                com.CommandType = CommandType.StoredProcedure;
                com.Parameters.AddWithValue("@ApproverPreEventURL", targetRow1?.Cells[1].Value ?? "no url");
                com.Parameters.AddWithValue("@FinanceTreasuryURL", targetRow2?.Cells[1].Value ?? "no url");
                com.Parameters.AddWithValue("@InitiatorURL", targetRow4?.Cells[1].Value ?? "no url");
                com.Parameters.AddWithValue("@EventTopic", formDataList.class1.EventTopic);
                com.Parameters.AddWithValue("@EventType", formDataList.class1.EventType);
                com.Parameters.AddWithValue("@EventDate", formDataList.class1.EventDate);
                com.Parameters.AddWithValue("@StartTime", formDataList.class1.StartTime);
                com.Parameters.AddWithValue("@EndTime", formDataList.class1.EndTime);
                com.Parameters.AddWithValue("@VenueName", formDataList.class1.VenueName);
                com.Parameters.AddWithValue("@City", formDataList.class1.City);
                com.Parameters.AddWithValue("@State", formDataList.class1.State);
                com.Parameters.AddWithValue("@MeetingType", " ");
                com.Parameters.AddWithValue("@Brands", brand);
                com.Parameters.AddWithValue("@Expenses", Expense);
                com.Parameters.AddWithValue("@Panelists", HCP);
                com.Parameters.AddWithValue("@Invitees", Invitees);
                com.Parameters.AddWithValue("@MIPLInvitees", MenariniInvitees);
                com.Parameters.AddWithValue("@SlideKits", slideKit);
                com.Parameters.AddWithValue("@IsAdvanceRequired", formDataList.class1.IsAdvanceRequired);
                com.Parameters.AddWithValue("@EventOpen30days", formDataList.class1.EventOpen30days);
                com.Parameters.AddWithValue("@EventWithin7days", formDataList.class1.EventWithin7days);
                com.Parameters.AddWithValue("@InitiatorName", formDataList.class1.InitiatorName);
                com.Parameters.AddWithValue("@AdvanceAmount", SheetHelper.NumCheck(formDataList.class1.AdvanceAmount));
                com.Parameters.AddWithValue("@TotalExpenseBTC", SheetHelper.NumCheck(formDataList.class1.TotalExpenseBTC));
                com.Parameters.AddWithValue("@TotalExpenseBTE", SheetHelper.NumCheck(formDataList.class1.TotalExpenseBTE));
                com.Parameters.AddWithValue("@TotalHonorariumAmount", Math.Round(TotalHonorariumAmount, 2));
                com.Parameters.AddWithValue("@TotalTravelAmount", Math.Round(TotalTravelAmount, 2));
                com.Parameters.AddWithValue("@TotalTravelAccommodationAmount", s);
                com.Parameters.AddWithValue("@TotalAccommodationAmount", Math.Round(TotalAccomodateAmount, 2));
                com.Parameters.AddWithValue("@BudgetAmount", total);
                com.Parameters.AddWithValue("@TotalLocalConveyance", c);
                com.Parameters.AddWithValue("@TotalExpense", Math.Round(TotalExpenseAmount, 2));
                com.Parameters.AddWithValue("@InitiatorEmail", formDataList.class1.Initiator_Email);
                com.Parameters.AddWithValue("@RBMBM", formDataList.class1.RBMorBM);
                com.Parameters.AddWithValue("@SalesHead", formDataList.class1.Sales_Head);
                com.Parameters.AddWithValue("@SalesCoordinator", formDataList.class1.SalesCoordinatorEmail);
                com.Parameters.AddWithValue("@MarketingCoordinator", formDataList.class1.MarketingCoordinatorEmail);
                com.Parameters.AddWithValue("@MarketingHead", formDataList.class1.Marketing_Head);
                com.Parameters.AddWithValue("@Compliance", formDataList.class1.ComplianceEmail);
                com.Parameters.AddWithValue("@FinanceAccounts", formDataList.class1.FinanceAccountsEmail);
                com.Parameters.AddWithValue("@FinanceTreasury", formDataList.class1.Finance);
                com.Parameters.AddWithValue("@ReportingManager", formDataList.class1.ReportingManagerEmail);
                com.Parameters.AddWithValue("@1UpManager", formDataList.class1.FirstLevelEmail);
                com.Parameters.AddWithValue("@MedicalAffairsHead", formDataList.class1.MedicalAffairsEmail);
                com.Parameters.AddWithValue("@BTEExpenseDetails", formDataList.class1.BTEExpenseDetails);
                com.Parameters.AddWithValue("@AttachmentPaths", Attachmentpaths);
                com.Parameters.AddWithValue("@webinarRole", formDataList.class1.Role);
                await MyConn.OpenAsync();
                //com.ExecuteNonQuery();
                MySqlDataReader reader = com.ExecuteReader();
                String RefID = "";
                while (reader.Read())
                {
                    RefID = reader["ID"].ToString();
                }
                await MyConn.CloseAsync();


                await MyConn.OpenAsync();
                com = new MySqlCommand("SPEventRequestPanelDetails", MyConn);
                com.CommandType = CommandType.StoredProcedure;
                foreach (var formData in formDataList.EventRequestHcpRole)
                {
                    String PanelAttachmentpaths = "";
                    if (formData.IsUpload == "Yes")
                    {
                        int j = 1;
                        foreach (string p in formData.FilesToUpload)
                        {
                            string[] words = p.Split(':');
                            string r = words[0];
                            string q = words[1];
                            string name = r.Split(".")[0];
                            string filePath = SheetHelper.SQlFileinsertion(q, name);
                            PanelAttachmentpaths = PanelAttachmentpaths + "," + filePath;
                        }

                    }
                    com.Parameters.AddWithValue("@HcpRole", formData.HcpRole);
                    com.Parameters.AddWithValue("@MISCode", SheetHelper.MisCodeCheck(formData.MisCode));
                    com.Parameters.AddWithValue("@Travel", SheetHelper.NumCheck(formData.Travel));
                    com.Parameters.AddWithValue("@TotalSpend", SheetHelper.NumCheck(formData.FinalAmount));
                    com.Parameters.AddWithValue("@Accomodation", SheetHelper.NumCheck(formData.Accomdation));
                    com.Parameters.AddWithValue("@LocalConveyance", SheetHelper.NumCheck(formData.LocalConveyance));
                    com.Parameters.AddWithValue("@SpeakerCode", formData.SpeakerCode);
                    com.Parameters.AddWithValue("@TrainerCode", formData.TrainerCode);
                    com.Parameters.AddWithValue("@HonorariumRequired", formData.HonorariumRequired);
                    com.Parameters.AddWithValue("@AgreementAmount", SheetHelper.NumCheck(formData.AgreementAmount));
                    com.Parameters.AddWithValue("@HonorariumAmount", SheetHelper.NumCheck(formData.HonarariumAmount));
                    com.Parameters.AddWithValue("@Speciality", formData.Speciality);
                    com.Parameters.AddWithValue("@EventTopic", formDataList.class1.EventTopic);
                    com.Parameters.AddWithValue("@EventType", formDataList.class1.EventType);
                    com.Parameters.AddWithValue("@EventDateStart", formDataList.class1.EventDate);
                    com.Parameters.AddWithValue("@EventEndDate", formDataList.class1.EventEndDate);
                    com.Parameters.AddWithValue("@HCPName", formData.HcpName);
                    com.Parameters.AddWithValue("@PANcardname", formData.PanCardName);
                    com.Parameters.AddWithValue("@ExpenseType", formData.ExpenseType);
                    com.Parameters.AddWithValue("@BankAccountNumber", formData.BankAccountNumber);
                    com.Parameters.AddWithValue("@BankName", formData.BankName);
                    com.Parameters.AddWithValue("@IFSCCode", formData.IFSCCode);
                    com.Parameters.AddWithValue("@FCPADate", formData.Fcpadate);
                    com.Parameters.AddWithValue("@Currency", formData.Currency);
                    com.Parameters.AddWithValue("@HonorariumAmountExcludingTax", formData.HonarariumAmountExcludingTax);
                    com.Parameters.AddWithValue("@TravelExcludingTax", formData.TravelExcludingTax);
                    com.Parameters.AddWithValue("@AccomodationExcludingTax", formData.AccomdationExcludingTax);
                    com.Parameters.AddWithValue("@LocalConveyanceExcludingTax", formData.LocalConveyanceExcludingTax);
                    com.Parameters.AddWithValue("@LCBTCBTE", formData.LcBtcorBte);
                    com.Parameters.AddWithValue("@TravelBTCBTE", formData.TravelBtcorBte);
                    com.Parameters.AddWithValue("@AccomodationBTCBTE", formData.AccomodationBtcorBte);
                    com.Parameters.AddWithValue("@ModeofTravel", formData.TravelSelection);

                    if (formData.Currency == "Others")
                    {
                        com.Parameters.AddWithValue("@OtherCurrency", formData.OtherCurrencyType);
                    }
                    else
                    {
                        com.Parameters.AddWithValue("@OtherCurrency", "");
                    }
                    com.Parameters.AddWithValue("@BeneficiaryName", formData.BeneficiaryName);
                    com.Parameters.AddWithValue("@PanNumber", formData.PanNumber);

                    if (formData.HcpRole == "Others")
                    {
                        com.Parameters.AddWithValue("@OtherType", formData.OthersType);
                    }
                    else
                    {
                        com.Parameters.AddWithValue("@OtherType", "");
                    }

                    com.Parameters.AddWithValue("@Tier", formData.Tier);
                    com.Parameters.AddWithValue("@HCPType", formData.GOorNGO);
                    com.Parameters.AddWithValue("@PresentationDuration", SheetHelper.NumCheck(formData.PresentationDuration));
                    com.Parameters.AddWithValue("@PanelSessionPreparationDuration", SheetHelper.NumCheck(formData.PanelSessionPreperationDuration));
                    com.Parameters.AddWithValue("@PanelDiscussionDuration", SheetHelper.NumCheck(formData.PanelDisscussionDuration));
                    com.Parameters.AddWithValue("@QASessionDuration", SheetHelper.NumCheck(formData.QASessionDuration));
                    com.Parameters.AddWithValue("@BriefingSession", SheetHelper.NumCheck(formData.BriefingSession));
                    com.Parameters.AddWithValue("@TotalSessionHours", SheetHelper.NumCheck(formData.TotalSessionHours));
                    com.Parameters.AddWithValue("@Rationale", formData.Rationale);
                    com.Parameters.AddWithValue("@EventIdEventRequestId", RefID);
                    com.Parameters.AddWithValue("@AttachmentPaths", PanelAttachmentpaths);
                    com.ExecuteNonQuery();
                    com.Parameters.Clear();
                }


                await MyConn.CloseAsync();


                await MyConn.OpenAsync();
                com = new MySqlCommand("SPEventRequestsBrandsList", MyConn);
                com.CommandType = CommandType.StoredProcedure;
                foreach (var formdata in formDataList.RequestBrandsList)
                {
                    com.Parameters.AddWithValue("@Allocation", formdata.PercentAllocation);
                    com.Parameters.AddWithValue("@Brands", formdata.BrandName);
                    com.Parameters.AddWithValue("@ProjectID", formdata.ProjectId);
                    com.Parameters.AddWithValue("@EventIdEventRequestId", RefID);
                    com.ExecuteNonQuery();
                    com.Parameters.Clear();
                }
                await MyConn.CloseAsync();


                await MyConn.OpenAsync();
                com = new MySqlCommand("SPEventRequestInvitees", MyConn);
                com.CommandType = CommandType.StoredProcedure;
                foreach (var formdata in formDataList.EventRequestInvitees)
                {
                    com.Parameters.AddWithValue("@HCPName", formdata.InviteeName);
                    com.Parameters.AddWithValue("@Designation", formdata.Designation);
                    com.Parameters.AddWithValue("@EmployeeCode", formdata.EmployeeCode);
                    com.Parameters.AddWithValue("@LocalConveyance", formdata.LocalConveyance);
                    com.Parameters.AddWithValue("@BTCBTE", formdata.BtcorBte);
                    com.Parameters.AddWithValue("@LcAmount", SheetHelper.NumCheck(formdata.LcAmount));
                    com.Parameters.AddWithValue("@LcAmountExcludingTax", formdata.LcAmountExcludingTax);
                    com.Parameters.AddWithValue("@EventIdEventRequestId", RefID);
                    com.Parameters.AddWithValue("@InviteeSource", formdata.InviteedFrom);
                    com.Parameters.AddWithValue("@HCPType", formdata.HCPType);
                    com.Parameters.AddWithValue("@Speciality", formdata.Speciality);
                    com.Parameters.AddWithValue("@MISCode", SheetHelper.MisCodeCheck(formdata.MISCode));
                    com.Parameters.AddWithValue("@EventTopic", formDataList.class1.EventTopic);
                    com.Parameters.AddWithValue("@EventType", formDataList.class1.EventType);
                    com.Parameters.AddWithValue("@EventDateStart", formDataList.class1.EventDate);
                    com.Parameters.AddWithValue("@EventEndDate", formDataList.class1.EventDate);
                    com.ExecuteNonQuery();
                    com.Parameters.Clear();
                }
                await MyConn.CloseAsync();


                await MyConn.OpenAsync();
                com = new MySqlCommand("SPEventRequestHCPSlideKitDetails", MyConn);
                com.CommandType = CommandType.StoredProcedure;
                String SlidekitsAttachent = "";
                foreach (var formdata in formDataList.EventRequestHCPSlideKits)
                {
                    if (formdata.IsUpload == "Yes")
                    {
                        int j = 1;
                        foreach (string p in formdata.FilesToUpload)
                        {
                            string[] words = p.Split(':');
                            string r = words[0];
                            string q = words[1];
                            string name = r.Split(".")[0];
                            string filePath = SheetHelper.SQlFileinsertion(q, name);
                            SlidekitsAttachent = SlidekitsAttachent + "," + filePath;
                        }
                    }
                    com.Parameters.AddWithValue("@MIS", SheetHelper.MisCodeCheck(formdata.MIS));
                    com.Parameters.AddWithValue("@SlideKitType", formdata.SlideKitType);
                    com.Parameters.AddWithValue("@SlideKitDocument", formdata.SlideKitDocument);
                    com.Parameters.AddWithValue("@EventIdEventRequestId", RefID);
                    com.Parameters.AddWithValue("@AttachmentPaths", SlidekitsAttachent);
                    com.ExecuteNonQuery();
                    com.Parameters.Clear();
                }

                MyConn.CloseAsync();


                MyConn.Open();
                com = new MySqlCommand("SPEventRequestExpensesSheet", MyConn);
                com.CommandType = CommandType.StoredProcedure;

                foreach (var formdata in formDataList.EventRequestExpenseSheet)
                {
                    com.Parameters.AddWithValue("@Expense", formdata.Expense);
                    com.Parameters.AddWithValue("@EventIdEventRequestID", RefID);
                    com.Parameters.AddWithValue("@AmountExcludingTaxq", formdata.AmountExcludingTax);
                    com.Parameters.AddWithValue("@AmountExcludingTax", formdata.ExcludingTaxAmount);
                    com.Parameters.AddWithValue("@Amount", SheetHelper.NumCheck(formdata.Amount));
                    com.Parameters.AddWithValue("@BTCBTE", formdata.BtcorBte);
                    com.Parameters.AddWithValue("@BudgetAmount", SheetHelper.NumCheck(formdata.BudgetAmount));
                    com.Parameters.AddWithValue("@BTCAmount", SheetHelper.NumCheck(formdata.BtcAmount));
                    com.Parameters.AddWithValue("@BTEAmount", SheetHelper.NumCheck(formdata.BteAmount));
                    com.Parameters.AddWithValue("@EventTopic", formDataList.class1.EventTopic);
                    com.Parameters.AddWithValue("@EventType", formDataList.class1.EventType);
                    com.Parameters.AddWithValue("@EventDateStart", formDataList.class1.EventDate);
                    com.Parameters.AddWithValue("@EventEndDate", formDataList.class1.EventDate);
                    com.ExecuteNonQuery();
                    com.Parameters.Clear();
                }
                MyConn.CloseAsync();

                if (formDataList.class1.EventOpen30days == "Yes" || formDataList.class1.EventWithin7days == "Yes" || formDataList.class1.FB_Expense_Excluding_Tax == "Yes" || formDataList.class1.IsDeviationUpload == "Yes")
                {
                    List<string> DeviationNames = new List<string>();
                    foreach (var p in formDataList.class1.DeviationDetails)
                    {
                        string[] words = p.DeviationFile.Split(':')[0].Split("*");
                        string r = words[1];
                        DeviationNames.Add(r);
                    }
                    await MyConn.OpenAsync();
                    com = new MySqlCommand("SPDeviation_Process", MyConn);
                    com.CommandType = CommandType.StoredProcedure;

                    foreach (var pp in formDataList.class1.DeviationDetails)
                    {
                        foreach (var deviationname in DeviationNames)
                        {
                            string file = deviationname.Split(".")[0];
                            string DeviationAttachmentpath = "";
                            if (pp.DeviationFile.Split(':')[0].Split("*")[1] == deviationname)
                            {
                                try
                                {
                                    foreach (var p in formDataList.class1.DeviationDetails)
                                    {
                                        string[] nameSplit = p.DeviationFile.Split("*");
                                        string[] words = nameSplit[1].Split(':');
                                        string r = words[0];
                                        string q = words[1];
                                        if (deviationname == r)
                                        {
                                            string name = nameSplit[0];
                                            string filePath = SheetHelper.SQlFileinsertion(q, name);
                                            DeviationAttachmentpath = DeviationAttachmentpath + "," + filePath;
                                            //Attachment attachmentinmain = await ApiCalls.AddAttachmentsToSheet(smartsheet, sheet1, addedRows[0], filePath);
                                        }
                                    }

                                    com.Parameters.AddWithValue("@EventIdEventRequestId", RefID);
                                    com.Parameters.AddWithValue("@EventTopic", formDataList.class1.EventTopic);
                                    com.Parameters.AddWithValue("@EventType", formDataList.class1.EventType);
                                    com.Parameters.AddWithValue("@EventDate", formDataList.class1.EventDate);
                                    com.Parameters.AddWithValue("@StartTime", formDataList.class1.StartTime);
                                    com.Parameters.AddWithValue("@EndTime", formDataList.class1.EndTime);
                                    com.Parameters.AddWithValue("@MISCode", SheetHelper.MisCodeCheck(pp.MisCode));
                                    com.Parameters.AddWithValue("@HCPName", pp.HcpName);
                                    com.Parameters.AddWithValue("@HonorariumAmount", pp.HonorariumAmountExcludingTax);
                                    com.Parameters.AddWithValue("@TravelAccommodationAmount", pp.TravelorAccomodationAmountExcludingTax);
                                    com.Parameters.AddWithValue("@OtherExpenses", pp.OtherExpenseAmountExcludingTax);


                                    if (file == "30DaysDeviationFile")
                                    {
                                        com.Parameters.AddWithValue("@DeviationType", configuration.GetSection("DeviationNamesInPreEvent:30DaysDeviationFile").Value);
                                        com.Parameters.AddWithValue("@EventOpen45days", formDataList.class1.EventOpen30days);
                                        com.Parameters.AddWithValue("@OutstandingEvents", SheetHelper.NumCheck(formDataList.class1.EventOpen30dayscount));
                                    }
                                    else
                                    {
                                        com.Parameters.AddWithValue("@EventOpen45days", "");
                                        com.Parameters.AddWithValue("@OutstandingEvents", "");
                                    }
                                    if (file == "7DaysDeviationFile")
                                    {
                                        com.Parameters.AddWithValue("@DeviationType", configuration.GetSection("DeviationNamesInPreEvent:5DaysDeviationFile").Value);
                                        com.Parameters.AddWithValue("@EventWithin5days", formDataList.class1.EventWithin7days);
                                    }
                                    else
                                    {
                                        com.Parameters.AddWithValue("@EventWithin5days", "");
                                    }
                                    if (file == "ExpenseExcludingTax")
                                    {
                                        com.Parameters.AddWithValue("@DeviationType", configuration.GetSection("DeviationNamesInPreEvent:ExpenseExcludingTax").Value);
                                        com.Parameters.AddWithValue("@PREExpenseExcludingTax", formDataList.class1.FB_Expense_Excluding_Tax);
                                    }
                                    else
                                    {
                                        com.Parameters.AddWithValue("@PREExpenseExcludingTax", "");
                                    }
                                    if (file.Contains("Travel_Accomodation3LExceededFile"))
                                    {
                                        com.Parameters.AddWithValue("@DeviationType", configuration.GetSection("DeviationNamesInPreEvent:Travel_Accomodation3LExceededFile").Value);
                                        com.Parameters.AddWithValue("@TravelAccomodationExceededTrigger", "Yes");
                                    }
                                    else
                                    {
                                        com.Parameters.AddWithValue("@TravelAccomodationExceededTrigger", "");
                                    }
                                    if (file.Contains("TrainerHonorarium12LExceededFile"))
                                    {
                                        com.Parameters.AddWithValue("@DeviationType", configuration.GetSection("DeviationNamesInPreEvent:TrainerHonorarium12LExceededFile").Value);
                                        com.Parameters.AddWithValue("@TrainerHonorariumExceededTrigger", "Yes");
                                    }
                                    else
                                    {
                                        com.Parameters.AddWithValue("@TrainerHonorariumExceededTrigger", "");
                                    }
                                    if (file.Contains("HCPHonorarium6LExceededFile"))
                                    {
                                        com.Parameters.AddWithValue("@DeviationType", configuration.GetSection("DeviationNamesInPreEvent:HCPHonorarium6LExceededFile").Value);
                                        com.Parameters.AddWithValue("@HCPHonorariumExceededTrigger", "Yes");
                                    }
                                    else
                                    {
                                        com.Parameters.AddWithValue("@HCPHonorariumExceededTrigger", "");
                                    }
                                    com.Parameters.AddWithValue("@SalesHead", formDataList.class1.Sales_Head);
                                    com.Parameters.AddWithValue("@FinanceHead", formDataList.class1.FinanceHead);
                                    com.Parameters.AddWithValue("@InitiatorName", formDataList.class1.InitiatorName);
                                    com.Parameters.AddWithValue("@InitiatorEmail", formDataList.class1.Initiator_Email);
                                    com.Parameters.AddWithValue("@SalesCoordinator", formDataList.class1.SalesCoordinatorEmail);
                                    com.Parameters.AddWithValue("@AttachmentPaths", DeviationAttachmentpath);
                                    com.Parameters.AddWithValue("@EndDate", "");
                                    com.ExecuteNonQuery();
                                    com.Parameters.Clear();
                                }
                                catch (Exception ex)
                                {
                                    return BadRequest(new
                                    {
                                        Message = ex.Message + "------" + ex.StackTrace
                                    });

                                }

                            }

                        }

                    }
                }

                //Row addedrow = addedRows[0];
                //long ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role");
                //Cell UpdateB = new Cell { ColumnId = ColumnId, Value = formDataList.Webinar.Role };
                //Row updateRows = new Row { Id = addedrow.Id, Cells = new Cell[] { UpdateB } };
                //Cell? cellsToUpdate = addedrow.Cells.FirstOrDefault(c => c.ColumnId == ColumnId);
                //if (cellsToUpdate != null) { cellsToUpdate.Value = formDataList.Webinar.Role; }

                //strMessage += "==Before adding Role to WebSheet " + "==" + DateTime.Now.ToString() + "==";
                //await Task.Run(() => ApiCalls.UpdateRole(smartsheet, sheet1, updateRows)); //smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRows }));
                //strMessage += "==Before adding Role to WebSheet " + "==" + DateTime.Now.ToString() + "==";
                //Log.Information("End of api " + DateTime.Now);
                MyConn.CloseAsync();
                return Ok(new
                {
                    Message = $"Thank you. Your event creation request has been received." +
                "You should receive a confirmation email with the details of your event after a few minutes."
                });

            }
            catch (Exception ex)
            {
                Log.Error($"Error occured on webinar method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(new
                { Message = ex.Message + "------" + ex.StackTrace });
            }
        }

        [HttpPost("ClassIIPreEvent"), DisableRequestSizeLimit]
        public IActionResult ClassIIPreEvent(Class2 formDataList)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
                Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
                Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
                Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
                Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
                Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
                Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);

                StringBuilder addedBrandsData = new();
                StringBuilder addedInviteesData = new();
                StringBuilder addedHcpData = new();
                StringBuilder addedSlideKitData = new();
                StringBuilder addedExpences = new();
                StringBuilder SelectedProduct = new();
                StringBuilder addedMEnariniInviteesData = new();
                int addedSlideKitDataNo = 1;
                int addedHcpDataNo = 1;
                int addedInviteesDataNo = 1;
                int addedBrandsDataNo = 1;
                int addedExpencesNo = 1;
                int SelectedProductNo = 1;
                int addedInviteesDataNoforMenarini = 1;

                foreach (var formdata in formDataList.InviteesData)
                {
                    if (formdata.InviteedFrom == "Menarini Employees")
                    {
                        string row = $"{addedInviteesDataNoforMenarini}. {formdata.Invitee_Employee_HcpName}";
                        addedMEnariniInviteesData.AppendLine(row);
                        addedInviteesDataNoforMenarini++;
                    }
                    else
                    {
                        string rowData = $"{addedInviteesDataNo}. {formdata.Invitee_Employee_HcpName}";
                        addedInviteesData.AppendLine(rowData);
                        addedInviteesDataNo++;
                    }
                }
                string Invitees = addedInviteesData.ToString();
                string MenariniInvitees = addedMEnariniInviteesData.ToString();
                foreach (var formdata in formDataList.ExpenseSheetData)
                {
                    string rowData = $"{addedExpencesNo}. {formdata.ExpenseType} | AmountExcludingTax: {formdata.ExpenseAmountExcludingTax}| Amount: {formdata.ExpenseAmountIncludingTax} | {formdata.IsBtcorBte}";
                    addedExpences.AppendLine(rowData);
                    addedExpencesNo++;
                }
                string Expense = addedExpences.ToString();

                foreach (var formdata in formDataList.SlideKitSelectionData)
                {
                    string rowData = $"{addedSlideKitDataNo}. {formdata.MisCode} | {formdata.SlideKitSelectionType} | Id :{formdata.SlideKitDocument}";
                    addedSlideKitData.AppendLine(rowData);
                    addedSlideKitDataNo++;
                }
                string slideKit = addedSlideKitData.ToString();
                foreach (var formdata in formDataList.BrandsListData)
                {
                    string rowData = $"{addedBrandsDataNo}. {formdata.BrandName} | {formdata.ProjectId} | {formdata.PercentAllocation}";
                    addedBrandsData.AppendLine(rowData);
                    addedBrandsDataNo++;
                }
                string brand = addedBrandsData.ToString();
                foreach (var formdata in formDataList.PanelistData)
                {

                    string rowData = $"{addedHcpDataNo}. {formdata.HcpType} |{formdata.HcpName} | Honr.Amt: {formdata.HonorariumAmountincludingTax} |Trav.&Acc.Amt: {formdata.TravelAmountIncludingTax + formdata.AccomodationAmountIncludingTax} |Rationale :{formdata.Rationale}";
                    addedHcpData.AppendLine(rowData);
                    addedHcpDataNo++;

                }
                string HCP = addedHcpData.ToString();
                StringBuilder addedExpencesBTE = new();
                int addedExpencesNoBTE = 1;
                foreach (var formdata in formDataList.ExpenseSheetData)
                {
                    if (formdata.IsBtcorBte.ToLower() == "bte")
                    {
                        string rowData = $"{addedExpencesNoBTE}. {formdata.ExpenseType} | Amount: {formdata.ExpenseAmountIncludingTax}";
                        addedExpencesBTE.AppendLine(rowData);
                        addedExpencesNoBTE++;
                    }
                }
                string BTEExpense = addedExpencesBTE.ToString();





                Row newRow = new Row
                {
                    Cells = new List<Cell>()
                };

                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Type"), Value = formDataList.ClassII.EventType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Date"), Value = formDataList.ClassII.EventDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.ClassII.EventName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Start Time"), Value = formDataList.ClassII.StartTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "End Time"), Value = formDataList.ClassII.EndTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Venue Name"), Value = formDataList.ClassII.VenueName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Objective"), Value = formDataList.ClassII.Objective });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "City"), Value = formDataList.ClassII.City });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "State"), Value = formDataList.ClassII.State });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Mode of Training"), Value = formDataList.ClassII.ModeOfTraining });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "HOT Webinar Vendor Name"), Value = formDataList.ClassII.VendorName });

                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Panelists"), Value = HCP });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Invitees"), Value = Invitees });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "MIPL Invitees"), Value = MenariniInvitees });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Brands"), Value = brand });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Expenses"), Value = Expense });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "SlideKits"), Value = slideKit });

                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Name"), Value = formDataList.ClassII.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Advance Amount"), Value = formDataList.ClassII.AdvanceAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, " Total Expense BTC"), Value = formDataList.ClassII.TotalExpenseBTC });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense BTE"), Value = formDataList.ClassII.TotalExpenseBTE });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Honorarium Amount"), Value = formDataList.ClassII.TotalHonorariumAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel Amount"), Value = formDataList.ClassII.TotalTravelAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel & Accommodation Amount"), Value = formDataList.ClassII.TotalTravelAccommodationAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Accommodation Amount"), Value = formDataList.ClassII.TotalAccomodationAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Budget Amount"), Value = formDataList.ClassII.TotalBudget });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Local Conveyance"), Value = formDataList.ClassII.TotalLocalConveyance });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense"), Value = formDataList.ClassII.TotalExpense });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Email"), Value = formDataList.ClassII.InitiatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "RBM/BM"), Value = formDataList.ClassII.RBMorBMEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Head"), Value = formDataList.ClassII.SalesHeadEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Coordinator"), Value = formDataList.ClassII.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Marketing Coordinator"), Value = formDataList.ClassII.MarketingCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Marketing Head"), Value = formDataList.ClassII.MarketingHeadEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Treasury"), Value = formDataList.ClassII.FinanceEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Compliance"), Value = formDataList.ClassII.ComplianceEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Accounts"), Value = formDataList.ClassII.FinanceAccountsEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Reporting Manager"), Value = formDataList.ClassII.ReportingManagerEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "1 Up Manager"), Value = formDataList.ClassII.FirstLevelEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Affairs Head"), Value = formDataList.ClassII.MedicalAffairsEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, ""), Value = formDataList.ClassII.AllIndiaDoctorsInvitedForTheEvent });
                //newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role"), Value = formDataList.class1.Role });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "BTE Expense Details"), Value = BTEExpense });


                // IList<Row> addedRows = smartsheet.SheetResources.RowResources.AddRows(sheet1.Id.Value, new Row[] { newRow });

                IList<Row> addedRows = smartsheet.SheetResources.RowResources.AddRows(sheet1.Id.Value, new Row[] { newRow });
                long eventIdColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                Cell? eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                string val = eventIdCell.DisplayValue;
                int x = 1;
                foreach (var p in formDataList.ClassII.FilesUpload)
                {
                    string[] words = p.Split(':');
                    string r = words[0];
                    string q = words[1];
                    string name = r.Split(".")[0];
                    string filePath = SheetHelper.testingFile(q, name);
                    Row addedRow = addedRows[0];
                    Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                    x++;
                    if (System.IO.File.Exists(filePath))
                    {
                        SheetHelper.DeleteFile(filePath);
                    }
                }
                if (formDataList.ClassII.IsDeviationUpload == "Yes")
                {
                    List<string> DeviationNames = new List<string>();
                    foreach (var p in formDataList.ClassII.DeviationFiles)
                    {

                        string[] words = p.Split(':')[0].Split("*");
                        string r = words[1];
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
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.ClassII.EventName });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Type"), Value = formDataList.ClassII.EventType });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Date"), Value = formDataList.ClassII.EventDate });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Start Time"), Value = formDataList.ClassII.StartTime });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "End Time"), Value = formDataList.ClassII.EndTime });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Venue Name"), Value = formDataList.ClassII.VenueName });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "City"), Value = formDataList.ClassII.City });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "State"), Value = formDataList.ClassII.State });

                            if (file == "30DaysDeviationFile")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:30DaysDeviationFile").Value });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = "Yes" /*formDataList.ClassII.EventOpen30days*/ });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = formDataList.ClassII.EventOpen45dayscount });
                            }
                            else if (file == "7DaysDeviationFile")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:5DaysDeviationFile").Value });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = formDataList.ClassII.IsDeviationUpload });

                            }
                            else if (file == "ExpenseExcludingTax")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:ExpenseExcludingTax").Value });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "PRE-F&B Expense Excluding Tax"), Value = formDataList.ClassII.IsDeviationUpload });
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

                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.ClassII.SalesHeadEmail });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formDataList.ClassII.FinanceHeadEmail });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Name"), Value = formDataList.ClassII.InitiatorName });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.ClassII.InitiatorEmail });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Coordinator"), Value = formDataList.ClassII.SalesCoordinatorEmail });

                            IList<Row> addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });

                            int j = 1;
                            foreach (var p in formDataList.ClassII.DeviationFiles)
                            {
                                string[] nameSplit = p.Split("*");
                                string[] words = nameSplit[1].Split(':');
                                string r = words[0];
                                string q = words[1];
                                if (deviationname == r)
                                {
                                    string name = nameSplit[0];
                                    string filePath = SheetHelper.testingFile(q, name);
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

                List<Row> newRows2 = new();
                foreach (var formdata in formDataList.BrandsListData)
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
                foreach (var formdata in formDataList.InviteesData)
                {
                    Row newRow3 = new()
                    {
                        Cells = new List<Cell>()
                        {
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "HCPName"), Value = formdata.Invitee_Employee_HcpName },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Employee Code"), Value = formdata.EmployeeCode },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "LocalConveyance"), Value = formdata.IsLocalConveyance },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "BTC/BTE"), Value = formdata.BtcorBte },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "LcAmount"), Value =formdata.LcAmountIncludingTax },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Lc Amount Excluding Tax"), Value = formdata.LcAmountExcludingTax },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "EventId/EventRequestId"), Value = val },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Invitee Source"), Value = formdata.InviteedFrom },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "HCP Type"), Value = formdata.HCPType },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Speciality"), Value = formdata.Speciality },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "MISCode"), Value = SheetHelper.MisCodeCheck(formdata.MISCode )},
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Topic"), Value = formDataList.ClassII.EventName },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Type"), Value = formDataList.ClassII.EventType },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Venue name"), Value = formDataList.ClassII.VenueName },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Date Start"), Value = formDataList.ClassII.EventDate },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event End Date"), Value = formDataList.ClassII.EventDate }

                        }
                    };
                    newRows3.Add(newRow3);
                }
                smartsheet.SheetResources.RowResources.AddRows(sheet3.Id.Value, newRows3.ToArray());

                foreach (var formData in formDataList.PanelistData)
                {
                    Row newRow1 = new()
                    {
                        Cells = new List<Cell>()
                    };
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HcpRole"), Value = formData.HcpRole });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "MISCode"), Value = SheetHelper.MisCodeCheck(formData.MISCode) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCPName"), Value = formData.HcpName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "SpeakerCode"), Value = formData.SpeakerCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TrainerCode"), Value = formData.TrainerCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Speciality"), Value = formData.Speciality });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Tier"), Value = formData.HcpCategory });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCP Type"), Value = formData.HcpType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Rationale"), Value = formData.Rationale });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "FCPA Date"), Value = formData.FCPAIssueDate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HonorariumRequired"), Value = formData.IsHonorariumApplicable });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PresentationDuration"), Value = formData.Presentation_Speaking_WorkshopDuration });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PanelSessionPreparationDuration"), Value = formData.DevelopmentofPresentationPanelSessionPreparation });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PanelDiscussionDuration"), Value = formData.PaneldiscussionSessionduration });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "QASessionDuration"), Value = formData.QASession });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "BriefingSession"), Value = formData.Speaker_TrainerBriefing });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TotalSessionHours"), Value = formData.TotalNoOfHours });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "AgreementAmount"), Value = formData.AgreementAmount });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "AgreementAmount"), Value = formData.HonorariumAmountincludingTax });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Honorarium Amount Excluding Tax"), Value = formData.HonorariumAmountexcludingTax });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "ExpenseType"), Value = formData.ExpenseType });

                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "LC BTC/BTE"), Value = formData.IsLCBTC_BTE });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Travel BTC/BTE"), Value = formData.IsTravelBTC_BTE });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Accomodation BTC/BTE"), Value = formData.IsAccomodationBTC_BTE });

                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Travel"), Value = formData.TravelAmountIncludingTax });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Accomodation"), Value = formData.AccomodationAmountIncludingTax });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "LocalConveyance"), Value = formData.LocalConveyanceAmountincludingTax });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Travel Excluding Tax"), Value = formData.TravelAmountExcludingTax });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Accomodation Excluding Tax"), Value = formData.AccomodationAmountExcludingTax });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Local Conveyance Excluding Tax"), Value = formData.LocalConveyanceAmountexcludingTax });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TotalSpend"), Value = formData.FinalAmount });

                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "LocalConveyance"), Value = formData.HcpQualification });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "LocalConveyance"), Value = formData.HcpCountry });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "LocalConveyance"), Value = formData.YTDspendIncludingCurrentEvent });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "LocalConveyance"), Value = formData.TravelandAccomodationspendincludingcurrentevent });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "LocalConveyance"), Value = formData.TravelSelection });



                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Topic"), Value = formDataList.ClassII.EventName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Type"), Value = formDataList.ClassII.EventType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Venue name"), Value = formDataList.ClassII.VenueName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Date Start"), Value = formDataList.ClassII.EventDate });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event End Date"), Value = formDataList.ClassII.EventEndDate });

                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "PAN card name"), Value = formData.BenificiaryDetailsData.NameasPerPAN });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Bank Account Number"), Value = formData.BenificiaryDetailsData.BankAccountNumber });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Bank Name"), Value = formData.BenificiaryDetailsData.BankName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "IFSC Code"), Value = formData.BenificiaryDetailsData.IFSCCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Currency"), Value = formData.BenificiaryDetailsData.Currency });

                    if (formData.BenificiaryDetailsData.Currency == "Others")
                    {
                        newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Other Currency"), Value = formData.BenificiaryDetailsData.EnterCurrencyType });
                    }
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Beneficiary Name"), Value = formData.BenificiaryDetailsData.BenificiaryName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Pan Number"), Value = formData.BenificiaryDetailsData.PANCardNumber });

                    if (formData.HcpRole == "Others")
                    {

                        newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Other Type"), Value = formData.HCPRoleName });
                    }

                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "EventId/EventRequestId"), Value = val });


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

                foreach (var formdata in formDataList.SlideKitSelectionData)
                {
                    Row newRow5 = new()
                    {
                        Cells = new List<Cell>()
                    };

                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "MIS"), Value = SheetHelper.MisCodeCheck(formdata.MisCode) });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "Slide Kit Type"), Value = formdata.SlideKitSelectionType });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "SlideKit Document"), Value = formdata.SlideKitDocument });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "EventId/EventRequestId"), Value = val });
                    //newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "EventId/EventRequestId"), Value = HcpName });
                    //newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "EventId/EventRequestId"), Value = HcpType });


                    IList<Row> row = smartsheet.SheetResources.RowResources.AddRows(sheet5.Id.Value, new Row[] { newRow5 });
                    if (formdata.IsUpload == "Yes")
                    {
                        foreach (string p in formdata.FilesToUpload)
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


                List<Row> newRows6 = new();
                foreach (var formdata in formDataList.ExpenseSheetData)
                {
                    Row newRow6 = new()
                    {
                        Cells = new List<Cell>()
                        {
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Expense"), Value = formdata.ExpenseType },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "EventId/EventRequestID"), Value = val },
                            //new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "AmountExcludingTax?"), Value = formdata.AmountExcludingTax },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount Excluding Tax"), Value = formdata.ExpenseAmountExcludingTax },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Amount"), Value = formdata.ExpenseAmountIncludingTax },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formdata.IsBtcorBte },
                            //new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BudgetAmount"), Value = SheetHelper.NumCheck(formdata.BudgetAmount) },
                            //new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTCAmount"), Value = SheetHelper.NumCheck(formdata.BtcAmount) },
                            //new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTEAmount"), Value = SheetHelper.NumCheck(formdata.BteAmount) },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Topic"), Value = formDataList.ClassII.EventName },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Type"), Value = formDataList.ClassII.EventType },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Venue name"), Value = formDataList.ClassII.VenueName },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event Date Start"), Value = formDataList.ClassII.EventDate },
                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "Event End Date"), Value = formDataList.ClassII.EventDate }
                        }
                    };
                    newRows6.Add(newRow6);
                }
                smartsheet.SheetResources.RowResources.AddRows(sheet6.Id.Value, newRows6.ToArray());


            }
            catch (Exception ex)
            {
                //Log.Error($"Error occured on AllPreEventsController Attachementfile method {ex.Message} at {DateTime.Now}");
                //Log.Error(ex.StackTrace);
                //return BadRequest(ex.Message);

                return BadRequest(new
                {
                    Message = ex.Message + "------" + ex.StackTrace
                });

            }


            return Ok(formDataList);
        }

        [HttpPost("WebinarPreEventSmartsheet"), DisableRequestSizeLimit]
        public async Task<IActionResult> WebinarPreEventSmartsheet(WebinarPayload formDataList)
        {




            string strMessage = string.Empty;

            //int timeInterval = 4000;
            //await Task.Delay(timeInterval);
            try
            {
                //semaphore = new SemaphoreSlim(0, 1);
                Log.Information("starting of api " + DateTime.Now);
                //SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                strMessage += "==Before get token==" + DateTime.Now.ToString() + "==";
                SmartsheetClient smartsheet = await Task.Run(() => SmartSheetBuilder.AccessClient(accessToken, _externalApiSemaphore));
                strMessage += "==After get token==" + DateTime.Now.ToString() + "==";
                Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
                strMessage += "==Websheet Get Completed==" + DateTime.Now.ToString() + "==";
                Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
                strMessage += "==BrandsList Get Completed==" + DateTime.Now.ToString() + "==";
                Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
                strMessage += "==Invitees Get Completed==" + DateTime.Now.ToString() + "==";
                Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
                strMessage += "==Panel Get Completed==" + DateTime.Now.ToString() + "==";
                Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
                strMessage += "==HcpSlideKit Get Completed==" + DateTime.Now.ToString() + "==";
                Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
                strMessage += "==ExpensesSheet Get Completed==" + DateTime.Now.ToString() + "==";
                Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);
                strMessage += "==Deviation Get Completed==" + DateTime.Now.ToString() + "==";
                Sheet UrlData = SheetHelper.GetSheetById(smartsheet, UI_URL);
                /////////////////////

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
                double TotalHonorariumAmount = 0;
                double TotalTravelAmount = 0;
                double TotalAccomodateAmount = 0;
                double TotalHCPLcAmount = 0;
                double TotalInviteesLcAmount = 0;
                double TotalExpenseAmount = 0;
                foreach (var formdata in formDataList.EventRequestExpenseSheet)
                {
                    string rowData = $"{addedExpencesNo}. {formdata.Expense} | AmountExcludingTax: {formdata.AmountExcludingTax}| Amount: {formdata.Amount} | {formdata.BtcorBte}";
                    addedExpences.AppendLine(rowData);
                    addedExpencesNo++;
                    var amount = SheetHelper.NumCheck(formdata.Amount);
                    TotalExpenseAmount = TotalExpenseAmount + amount;
                }
                string Expense = addedExpences.ToString();

                StringBuilder addedExpencesBTE = new();
                int addedExpencesNoBTE = 1;
                foreach (var formdata in formDataList.EventRequestExpenseSheet)
                {
                    if (formdata.BtcorBte.ToLower() == "bte")
                    {
                        string rowData = $"{addedExpencesNoBTE}. {formdata.Expense} | Amount: {formdata.Amount}";
                        addedExpencesBTE.AppendLine(rowData);
                        addedExpencesNoBTE++;
                    }
                }
                string BTEExpense = addedExpencesBTE.ToString();

                foreach (var formdata in formDataList.EventRequestHCPSlideKits)
                {
                    string rowData = $"{addedSlideKitDataNo}. {formdata.HcpName} | {formdata.SlideKitType} | Id :{formdata.SlideKitDocument}";
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
                    double HM = SheetHelper.NumCheck(formdata.HonarariumAmount);
                    double t = SheetHelper.NumCheck(formdata.Travel) + SheetHelper.NumCheck(formdata.Accomdation);

                    double roundedValue = Math.Round(t, 2);

                    string rowData = $"{addedHcpDataNo}. {formdata.HcpRole} |{formdata.HcpName} | Honr.Amt: {HM} |Trav.&Acc.Amt: {roundedValue} |Rationale : {formdata.Rationale}";
                    addedHcpData.AppendLine(rowData);
                    addedHcpDataNo++;
                    TotalHonorariumAmount = TotalHonorariumAmount + SheetHelper.NumCheck(formdata.HonarariumAmount);
                    TotalTravelAmount = TotalTravelAmount + SheetHelper.NumCheck(formdata.Travel);
                    TotalAccomodateAmount = TotalAccomodateAmount + SheetHelper.NumCheck(formdata.Accomdation);
                    TotalHCPLcAmount = TotalHCPLcAmount + SheetHelper.NumCheck(formdata.LocalConveyance);
                }
                string HCP = addedHcpData.ToString();


                double cc = TotalHCPLcAmount + TotalInviteesLcAmount;

                double totalAmount = TotalHonorariumAmount + TotalTravelAmount + TotalAccomodateAmount + TotalHCPLcAmount + TotalInviteesLcAmount + TotalExpenseAmount;

                double ss = TotalTravelAmount + TotalAccomodateAmount;

                double c = Math.Round(cc, 2);
                double total = Math.Round(totalAmount, 2);
                double s = Math.Round(ss, 2);

                Dictionary<string, long> Sheet1columns = new();
                foreach (var column in sheet1.Columns)
                {
                    Sheet1columns.Add(column.Title, (long)column.Id);
                }
                Row newRow = new()
                {
                    Cells = new List<Cell>()
                };

                Row? targetRow1 = UrlData.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == "Approver Pre Event URL"));
                Row? targetRow2 = UrlData.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == "Finance Treasury URL"));
                Row? targetRow4 = UrlData.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == "Initiator URL"));

                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Approver Pre Event URL"], Value = targetRow1?.Cells[1].Value ?? "no url" });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Finance Treasury URL"], Value = targetRow2?.Cells[1].Value ?? "no url" });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Initiator URL"], Value = targetRow4?.Cells[1].Value ?? "no url" });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Event Topic"], Value = formDataList.Webinar.EventTopic });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Event Type"], Value = formDataList.Webinar.EventType });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Event Date"], Value = formDataList.Webinar.EventDate });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Start Time"], Value = formDataList.Webinar.StartTime });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["End Time"], Value = formDataList.Webinar.EndTime });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Meeting Type"], Value = formDataList.Webinar.Meeting_Type });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Brands"], Value = brand });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Expenses"], Value = Expense });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Panelists"], Value = HCP });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Invitees"], Value = Invitees });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["MIPL Invitees"], Value = MenariniInvitees });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["SlideKits"], Value = slideKit });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["IsAdvanceRequired"], Value = formDataList.Webinar.IsAdvanceRequired });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["EventOpen30days"], Value = formDataList.Webinar.EventOpen30days });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["EventWithin7days"], Value = formDataList.Webinar.EventWithin7days });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Initiator Name"], Value = formDataList.Webinar.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Advance Amount"], Value = SheetHelper.NumCheck(formDataList.Webinar.AdvanceAmount) });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns[" Total Expense BTC"], Value = SheetHelper.NumCheck(formDataList.Webinar.TotalExpenseBTC) });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Total Expense BTE"], Value = SheetHelper.NumCheck(formDataList.Webinar.TotalExpenseBTE) });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Total Honorarium Amount"], Value = Math.Round(TotalHonorariumAmount, 2) });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Total Travel Amount"], Value = Math.Round(TotalTravelAmount, 2) });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Total Travel & Accommodation Amount"], Value = s });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Total Accommodation Amount"], Value = Math.Round(TotalAccomodateAmount, 2) });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Budget Amount"], Value = total });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Total Local Conveyance"], Value = c });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Total Expense"], Value = Math.Round(TotalExpenseAmount, 2) });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Initiator Email"], Value = formDataList.Webinar.Initiator_Email });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["RBM/BM"], Value = formDataList.Webinar.RBMorBM });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Sales Head"], Value = formDataList.Webinar.Sales_Head });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Sales Coordinator"], Value = formDataList.Webinar.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Marketing Coordinator"], Value = formDataList.Webinar.MarketingCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Marketing Head"], Value = formDataList.Webinar.Marketing_Head });
                //newRow.Cells.Add(new Cell { ColumnId Sheet1columns[, "Finance Treasury"), Value = formData.RequestHonorariumList.MarketingHeadEmail });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Compliance"], Value = formDataList.Webinar.ComplianceEmail });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Finance Accounts"], Value = formDataList.Webinar.FinanceAccountsEmail });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Finance Treasury"], Value = formDataList.Webinar.Finance });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Reporting Manager"], Value = formDataList.Webinar.ReportingManagerEmail });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["1 Up Manager"], Value = formDataList.Webinar.FirstLevelEmail });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Medical Affairs Head"], Value = formDataList.Webinar.MedicalAffairsEmail });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["BTE Expense Details"], Value = formDataList.Webinar.BTEExpenseDetails });

                strMessage += "==Before adding to websheet==" + DateTime.Now.ToString() + "==";
                IList<Row> addedRows = ApiCalls.WebDetails(smartsheet, sheet1, newRow);
                strMessage += "==after adding to websheet==" + DateTime.Now.ToString() + "==";
                //IList<Row> addedRows = await Task.Run(() => smartsheet.SheetResources.RowResources.AddRows(sheet1.Id.Value, new Row[] { newRow }));
                ///////
                //IList<Row> addedRows = (IList<Row>)ApiCalls.AddWebinarData(smartsheet, sheet1, formDataList);

                long eventIdColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                Cell? eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                string val = eventIdCell.DisplayValue;

                int i = 1;

                foreach (var p in formDataList.Webinar.Files)
                {
                    string[] words = p.Split(':');
                    string r = words[0];
                    string q = words[1];
                    string name = r.Split(".")[0];
                    string filePath = SheetHelper.testingFile(q, name);
                    Row addedRow = addedRows[0];
                    strMessage += "==Before adding to websheet attachment" + i.ToString() + "==" + DateTime.Now.ToString() + "==";
                    Attachment attachment = await ApiCalls.AddAttachmentsToSheet(smartsheet, sheet1, addedRow, filePath);
                    strMessage += "==after adding to websheet attachment" + i.ToString() + "==" + DateTime.Now.ToString() + "==";
                    i++;
                    //Attachment attachment = await Task.Run(() => smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                    //        sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword"));

                    if (System.IO.File.Exists(filePath))
                    {
                        SheetHelper.DeleteFile(filePath);
                    }
                }


                i = 1;
                Dictionary<string, long> Sheet4columns = new();
                foreach (var column in sheet4.Columns)
                {
                    Sheet4columns.Add(column.Title, (long)column.Id);
                }
                foreach (var formData in formDataList.EventRequestHcpRole)
                {
                    Row newRow1 = new()
                    {
                        Cells = new List<Cell>()
                    };
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["HcpRole"], Value = formData.HcpRole });
                    //newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HcpRole"), Value = formData.HcpRole });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["MISCode"], Value = SheetHelper.MisCodeCheck(formData.MisCode) });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Travel"], Value = SheetHelper.NumCheck(formData.Travel) });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["TotalSpend"], Value = SheetHelper.NumCheck(formData.FinalAmount) });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Accomodation"], Value = SheetHelper.NumCheck(formData.Accomdation) });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["LocalConveyance"], Value = SheetHelper.NumCheck(formData.LocalConveyance) });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["SpeakerCode"], Value = formData.SpeakerCode });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["TrainerCode"], Value = formData.TrainerCode });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["HonorariumRequired"], Value = formData.HonorariumRequired });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["AgreementAmount"], Value = SheetHelper.NumCheck(formData.AgreementAmount) });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["HonorariumAmount"], Value = SheetHelper.NumCheck(formData.HonarariumAmount) });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Speciality"], Value = formData.Speciality });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Event Topic"], Value = formDataList.Webinar.EventTopic });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Event Type"], Value = formDataList.Webinar.EventType });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Event Date Start"], Value = formDataList.Webinar.EventDate });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Event End Date"], Value = formDataList.Webinar.EventEndDate });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["HCPName"], Value = formData.HcpName });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["PAN card name"], Value = formData.PanCardName });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["ExpenseType"], Value = formData.ExpenseType });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Bank Account Number"], Value = formData.BankAccountNumber });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Bank Name"], Value = formData.BankName });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["IFSC Code"], Value = formData.IFSCCode });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["FCPA Date"], Value = formData.Fcpadate });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Currency"], Value = formData.Currency });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Honorarium Amount Excluding Tax"], Value = formData.HonarariumAmountExcludingTax });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Travel Excluding Tax"], Value = formData.TravelExcludingTax });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Accomodation Excluding Tax"], Value = formData.AccomdationExcludingTax });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Local Conveyance Excluding Tax"], Value = formData.LocalConveyanceExcludingTax });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["LC BTC/BTE"], Value = formData.LcBtcorBte });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Travel BTC/BTE"], Value = formData.TravelBtcorBte });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Accomodation BTC/BTE"], Value = formData.AccomodationBtcorBte });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Mode of Travel"], Value = formData.TravelSelection });
                    if (formData.Currency == "Others")
                    {
                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Other Currency"], Value = formData.OtherCurrencyType });
                    }
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Beneficiary Name"], Value = formData.BeneficiaryName });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Pan Number"], Value = formData.PanNumber });

                    if (formData.HcpRole == "Others")
                    {

                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Other Type"], Value = formData.OthersType });
                    }

                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Tier"], Value = formData.Tier });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["HCP Type"], Value = formData.GOorNGO });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["PresentationDuration"], Value = SheetHelper.NumCheck(formData.PresentationDuration) });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["PanelSessionPreparationDuration"], Value = SheetHelper.NumCheck(formData.PanelSessionPreperationDuration) });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["PanelDiscussionDuration"], Value = SheetHelper.NumCheck(formData.PanelDisscussionDuration) });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["QASessionDuration"], Value = SheetHelper.NumCheck(formData.QASessionDuration) });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["BriefingSession"], Value = SheetHelper.NumCheck(formData.BriefingSession) });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["TotalSessionHours"], Value = SheetHelper.NumCheck(formData.TotalSessionHours) });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Rationale"], Value = formData.Rationale });
                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["EventId/EventRequestId"], Value = val });

                    strMessage += "==Before adding to Panel " + i.ToString() + "==" + DateTime.Now.ToString() + "==";
                    IList<Row> row = await Task.Run(() => ApiCalls.PanelDetails(smartsheet, sheet4, newRow1));
                    strMessage += "==After adding to Panel " + i.ToString() + "==" + DateTime.Now.ToString() + "==";
                    i++;
                    //IList<Row> row = await Task.Run(() => smartsheet.SheetResources.
                    //RowResources.AddRows(sheet4.Id.Value, new Row[] { newRow1 }));
                    if (formData.IsUpload == "Yes")
                    {
                        int j = 1;
                        foreach (string p in formData.FilesToUpload)
                        {
                            string[] words = p.Split(':');
                            string r = words[0];
                            string q = words[1];
                            string name = r.Split(".")[0];
                            string filePath = SheetHelper.testingFile(q, name);
                            Row addedRow = row[0];
                            //Attachment attachment = await Task.Run(() => smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                            //       sheet4.Id.Value, addedRow.Id.Value, filePath, "application/msword"));

                            strMessage += "==Before adding to panel attachment" + j.ToString() + "==" + DateTime.Now.ToString() + "==";
                            Attachment attachment = await ApiCalls.AddAttachmentsToSheet(smartsheet, sheet4, addedRow, filePath);
                            strMessage += "==after adding to panel attachment" + j.ToString() + "==" + DateTime.Now.ToString() + "==";
                            j++;
                            if (System.IO.File.Exists(filePath))
                            {
                                SheetHelper.DeleteFile(filePath);
                            }
                        }

                    }


                }
                Dictionary<string, long> Sheet2columns = new();
                foreach (var column in sheet2.Columns)
                {
                    Sheet2columns.Add(column.Title, (long)column.Id);
                }
                List<Row> newRows2 = new();
                foreach (var formdata in formDataList.RequestBrandsList)
                {
                    Row newRow2 = new()
                    {
                        Cells = new List<Cell>()
                        {
                            new(){ ColumnId = Sheet2columns[ "% Allocation"], Value = formdata.PercentAllocation },
                            new(){ ColumnId = Sheet2columns[ "Brands"], Value = formdata.BrandName },
                            new(){ ColumnId = Sheet2columns[ "Project ID"], Value = formdata.ProjectId },
                            new(){ ColumnId = Sheet2columns[ "EventId/EventRequestId"], Value = val }
                        }
                    };

                    newRows2.Add(newRow2);
                }
                strMessage += "==Before adding to Brands Bulk" + "==" + DateTime.Now.ToString() + "==";
                await Task.Run(() => /*smartsheet.SheetResources.RowResources.AddRows(sheet2.Id.Value, newRows2.ToArray())*/
                ApiCalls.BrandsDetails(smartsheet, sheet2, newRows2));
                strMessage += "==after adding to Brands Bulk" + "==" + DateTime.Now.ToString() + "==";

                Dictionary<string, long> Sheet3columns = new();
                foreach (var column in sheet3.Columns)
                {
                    Sheet3columns.Add(column.Title, (long)column.Id);
                }
                List<Row> newRows3 = new();
                foreach (var formdata in formDataList.EventRequestInvitees)
                {
                    Row newRow3 = new()
                    {
                        Cells = new List<Cell>()
                        {
                            new () { ColumnId = Sheet3columns[ "HCPName"], Value = formdata.InviteeName },
                            new () { ColumnId = Sheet3columns[ "Designation"], Value = formdata.Designation },
                            new () { ColumnId = Sheet3columns[ "Employee Code"], Value = formdata.EmployeeCode },
                            new () { ColumnId = Sheet3columns[ "LocalConveyance"], Value = formdata.LocalConveyance },
                            new () { ColumnId = Sheet3columns[ "BTC/BTE"], Value = formdata.BtcorBte },
                            new () { ColumnId = Sheet3columns[ "LcAmount"], Value = SheetHelper.NumCheck(formdata.LcAmount) },
                            new () { ColumnId = Sheet3columns[ "Lc Amount Excluding Tax"], Value = formdata.LcAmountExcludingTax },
                            new () { ColumnId = Sheet3columns[ "EventId/EventRequestId"], Value = val },
                            new () { ColumnId = Sheet3columns[ "Invitee Source"], Value = formdata.InviteedFrom },
                            new () { ColumnId = Sheet3columns[ "HCP Type"], Value = formdata.HCPType },
                            new () { ColumnId = Sheet3columns[ "Speciality"], Value = formdata.Speciality },
                            new () { ColumnId = Sheet3columns[ "MISCode"], Value = SheetHelper.MisCodeCheck(formdata.MISCode) },
                            new () { ColumnId = Sheet3columns[ "Event Topic"], Value = formDataList.Webinar.EventTopic },
                            new () { ColumnId = Sheet3columns[ "Event Type"], Value = formDataList.Webinar.EventType },
                            new () { ColumnId = Sheet3columns[ "Event Date Start"], Value = formDataList.Webinar.EventDate },
                            new () { ColumnId = Sheet3columns[ "Event End Date"], Value = formDataList.Webinar.EventDate }

                        }
                    };
                    newRows3.Add(newRow3);
                }
                strMessage += "==Before adding to Invitees Bulk" + "==" + DateTime.Now.ToString() + "==";
                await Task.Run(() => /*smartsheet.SheetResources.RowResources.AddRows(sheet3.Id.Value, newRows3.ToArray()));*/
                ApiCalls.InviteesDetails(smartsheet, sheet3, newRows3));
                strMessage += "==After adding to Invitees Bulk" + "==" + DateTime.Now.ToString() + "==";
                i = 1;

                Dictionary<string, long> Sheet5columns = new();
                foreach (var column in sheet5.Columns)
                {
                    Sheet5columns.Add(column.Title, (long)column.Id);
                }
                foreach (var formdata in formDataList.EventRequestHCPSlideKits)
                {
                    Row newRow5 = new()
                    {
                        Cells = new List<Cell>()
                    };

                    newRow5.Cells.Add(new Cell { ColumnId = Sheet5columns["MIS"], Value = SheetHelper.MisCodeCheck(formdata.MIS) });
                    newRow5.Cells.Add(new Cell { ColumnId = Sheet5columns["Slide Kit Type"], Value = formdata.SlideKitType });
                    newRow5.Cells.Add(new Cell { ColumnId = Sheet5columns["SlideKit Document"], Value = formdata.SlideKitDocument });
                    newRow5.Cells.Add(new Cell { ColumnId = Sheet5columns["EventId/EventRequestId"], Value = val });

                    strMessage += "==Before adding to Slidekit Data" + i.ToString() + "==" + DateTime.Now.ToString() + "==";
                    IList<Row> row = await Task.Run(() => ApiCalls.SlideKitDetails(smartsheet, sheet5, newRow5));/*smartsheet.SheetResources.RowResources.AddRows(sheet5.Id.Value, new Row[] { newRow5 }));*/
                    strMessage += "==After adding to SlideKit Data" + i.ToString() + "==" + DateTime.Now.ToString() + "==";
                    i++;
                    if (formdata.IsUpload == "Yes")
                    {
                        int j = 1;
                        foreach (string p in formdata.FilesToUpload)
                        {
                            string[] words = p.Split(':');
                            string r = words[0];
                            string q = words[1];
                            string name = r.Split(".")[0];
                            string filePath = SheetHelper.testingFile(q, name);
                            Row addedRow = row[0];
                            //Attachment attachment = await Task.Run(() => smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                            //       sheet5.Id.Value, addedRow.Id.Value, filePath, "application/msword"));
                            strMessage += "==Before adding to SlideKit Attachment" + j.ToString() + "==" + DateTime.Now.ToString() + "==";
                            Attachment attachment = await ApiCalls.AddAttachmentsToSheet(smartsheet, sheet5, addedRow, filePath);
                            strMessage += "==After adding to  SlideKit Attachment" + j.ToString() + "==" + DateTime.Now.ToString() + "==";
                            j++;
                            if (System.IO.File.Exists(filePath))
                            {
                                SheetHelper.DeleteFile(filePath);
                            }
                        }
                    }
                }
                Dictionary<string, long> Sheet6columns = new();
                foreach (var column in sheet6.Columns)
                {
                    Sheet6columns.Add(column.Title, (long)column.Id);
                }
                List<Row> newRows6 = new();
                foreach (var formdata in formDataList.EventRequestExpenseSheet)
                {
                    Row newRow6 = new()
                    {
                        Cells = new List<Cell>()
                        {
                            new () { ColumnId = Sheet6columns[ "Expense"], Value = formdata.Expense },
                            new () { ColumnId = Sheet6columns[ "EventId/EventRequestID"], Value = val },
                            new () { ColumnId = Sheet6columns[ "AmountExcludingTax?"], Value = formdata.AmountExcludingTax },
                            new () { ColumnId = Sheet6columns[ "Amount Excluding Tax"], Value = formdata.ExcludingTaxAmount },
                            new () { ColumnId = Sheet6columns[ "Amount"], Value = SheetHelper.NumCheck(formdata.Amount) },
                            new () { ColumnId = Sheet6columns[ "BTC/BTE"], Value = formdata.BtcorBte },
                            new () { ColumnId = Sheet6columns[ "BudgetAmount"], Value = SheetHelper.NumCheck(formdata.BudgetAmount) },
                            new () { ColumnId = Sheet6columns[ "BTCAmount"], Value = SheetHelper.NumCheck(formdata.BtcAmount) },
                            new () { ColumnId = Sheet6columns[ "BTEAmount"], Value = SheetHelper.NumCheck(formdata.BteAmount) },
                            new () { ColumnId = Sheet6columns[ "Event Topic"], Value = formDataList.Webinar.EventTopic },
                            new () { ColumnId = Sheet6columns[ "Event Type"], Value = formDataList.Webinar.EventType },
                            new () { ColumnId = Sheet6columns[ "Event Date Start"], Value = formDataList.Webinar.EventDate },
                            new () { ColumnId = Sheet6columns[ "Event End Date"], Value = formDataList.Webinar.EventDate }
                        }
                    };
                    newRows6.Add(newRow6);
                }
                strMessage += "==Before adding to Expense Bulk" + "==" + DateTime.Now.ToString() + "==";
                await Task.Run(() => ApiCalls.ExpenseDetails(smartsheet, sheet6, newRows6));/* smartsheet.SheetResources.RowResources.AddRows(sheet6.Id.Value, newRows6.ToArray()));*/
                strMessage += "==After adding to Expense Bulk" + "==" + DateTime.Now.ToString() + "==";
                if (formDataList.Webinar.EventOpen30days == "Yes" || formDataList.Webinar.EventWithin7days == "Yes" || formDataList.Webinar.FB_Expense_Excluding_Tax == "Yes" || formDataList.Webinar.IsDeviationUpload == "Yes")
                {
                    i = 1;
                    List<string> DeviationNames = new List<string>();
                    foreach (var p in formDataList.Webinar.DeviationDetails)
                    {

                        string[] words = p.DeviationFile.Split(':')[0].Split("*");
                        string r = words[1];
                        DeviationNames.Add(r);
                    }
                    Dictionary<string, long> Sheet7columns = new();
                    foreach (var column in sheet7.Columns)
                    {
                        Sheet7columns.Add(column.Title, (long)column.Id);
                    }
                    foreach (var pp in formDataList.Webinar.DeviationDetails)
                    {
                        foreach (var deviationname in DeviationNames)
                        {
                            string file = deviationname.Split(".")[0];
                            string eventId = val;
                            if (pp.DeviationFile.Split(':')[0].Split("*")[1] == deviationname)
                            {
                                try
                                {
                                    Row newRow7 = new()
                                    {
                                        Cells = new List<Cell>()
                                    };
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["EventId/EventRequestId"], Value = eventId });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Event Topic"], Value = formDataList.Webinar.EventTopic });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Event Type"], Value = formDataList.Webinar.EventType });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Event Date"], Value = formDataList.Webinar.EventDate });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Start Time"], Value = formDataList.Webinar.StartTime });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["End Time"], Value = formDataList.Webinar.EndTime });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["MIS Code"], Value = SheetHelper.MisCodeCheck(pp.MisCode) });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["HCP Name"], Value = pp.HcpName });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Honorarium Amount"], Value = pp.HonorariumAmountExcludingTax });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Travel & Accommodation Amount"], Value = pp.TravelorAccomodationAmountExcludingTax });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Other Expenses"], Value = pp.OtherExpenseAmountExcludingTax });

                                    if (file == "30DaysDeviationFile")
                                    {
                                        newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Deviation Type"], Value = configuration.GetSection("DeviationNamesInPreEvent:30DaysDeviationFile").Value });
                                        newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["EventOpen45days"], Value = formDataList.Webinar.EventOpen30days });
                                        newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Outstanding Events"], Value = SheetHelper.NumCheck(formDataList.Webinar.EventOpen30dayscount) });
                                    }
                                    else if (file == "7DaysDeviationFile")
                                    {
                                        newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Deviation Type"], Value = configuration.GetSection("DeviationNamesInPreEvent:5DaysDeviationFile").Value });
                                        newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["EventWithin5days"], Value = formDataList.Webinar.EventWithin7days });

                                    }
                                    else if (file == "ExpenseExcludingTax")
                                    {
                                        newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Deviation Type"], Value = configuration.GetSection("DeviationNamesInPreEvent:ExpenseExcludingTax").Value });
                                        newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["PRE-F&B Expense Excluding Tax"], Value = formDataList.Webinar.FB_Expense_Excluding_Tax });
                                    }
                                    else if (file.Contains("Travel_Accomodation3LExceededFile"))
                                    {
                                        newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Deviation Type"], Value = configuration.GetSection("DeviationNamesInPreEvent:Travel_Accomodation3LExceededFile").Value });
                                        newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Travel/Accomodation 3,00,000 Exceeded Trigger"], Value = "Yes" });//formDataList.class1.FB_Expense_Excluding_Tax });
                                    }
                                    else if (file.Contains("TrainerHonorarium12LExceededFile"))
                                    {
                                        newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Deviation Type"], Value = configuration.GetSection("DeviationNamesInPreEvent:TrainerHonorarium12LExceededFile").Value/*"Honorarium Aggregate Limit of 12,00,000 is Exceeded"*/ });
                                        newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Trainer Honorarium 12,00,000 Exceeded Trigger"], Value = "Yes" }); //formDataList.class1.FB_Expense_Excluding_Tax });
                                    }
                                    else if (file.Contains("HCPHonorarium6LExceededFile"))
                                    {
                                        newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Deviation Type"], Value = configuration.GetSection("DeviationNamesInPreEvent:HCPHonorarium6LExceededFile").Value/*"Honorarium Aggregate Limit of 6,00,000 is Exceeded"*/ });
                                        newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["HCP Honorarium 6,00,000 Exceeded Trigger"], Value = "Yes" }); // formDataList.class1.FB_Expense_Excluding_Tax });
                                    }

                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Sales Head"], Value = formDataList.Webinar.Sales_Head });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Finance Head"], Value = formDataList.Webinar.FinanceHead });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Initiator Name"], Value = formDataList.Webinar.InitiatorName });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Initiator Email"], Value = formDataList.Webinar.Initiator_Email });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Sales Coordinator"], Value = formDataList.Webinar.SalesCoordinatorEmail });
                                    strMessage += "==Before adding to Deviation Data" + i.ToString() + "==" + DateTime.Now.ToString() + "==";
                                    IList<Row> addeddeviationrow = ApiCalls.DeviationData(smartsheet, sheet7, newRow7);//await Task.Run(() => smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 }));
                                    strMessage += "==Before adding to Deviation Data" + i.ToString() + "==" + DateTime.Now.ToString() + "==";
                                    int j = 1;
                                    foreach (var p in formDataList.Webinar.DeviationDetails)
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
                                            strMessage += "==Before adding to Deviation Attachment" + j.ToString() + "==" + DateTime.Now.ToString() + "==";
                                            Attachment attachment = await ApiCalls.AddAttachmentsToSheet(smartsheet, sheet7, addedRow, filePath);
                                            strMessage += "==After adding to Deviation Attachment" + j.ToString() + "==" + DateTime.Now.ToString() + "==";
                                            strMessage += "==Before adding to WebSheet Attachment" + j.ToString() + "==" + DateTime.Now.ToString() + "==";
                                            Attachment attachmentinmain = await ApiCalls.AddAttachmentsToSheet(smartsheet, sheet1, addedRows[0], filePath);
                                            strMessage += "==After adding to WebSheet Attachment" + j.ToString() + "==" + DateTime.Now.ToString() + "==";

                                            //Attachment attachment = await Task.Run(() => smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword"));
                                            //Attachment attachmentinmain = await Task.Run(() => smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet1.Id.Value, addedRows[0].Id.Value, filePath, "application/msword"));
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



                                    return BadRequest(new
                                    {
                                        Message = ex.Message + "------" + ex.StackTrace
                                    });

                                }

                            }

                        }

                    }
                }

                Row addedrow = addedRows[0];
                long ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role");
                Cell UpdateB = new Cell { ColumnId = ColumnId, Value = formDataList.Webinar.Role };
                Row updateRows = new Row { Id = addedrow.Id, Cells = new Cell[] { UpdateB } };
                Cell? cellsToUpdate = addedrow.Cells.FirstOrDefault(c => c.ColumnId == ColumnId);
                if (cellsToUpdate != null) { cellsToUpdate.Value = formDataList.Webinar.Role; }

                strMessage += "==Before adding Role to WebSheet " + "==" + DateTime.Now.ToString() + "==";
                await Task.Run(() => ApiCalls.UpdateRole(smartsheet, sheet1, updateRows)); //smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRows }));
                strMessage += "==Before adding Role to WebSheet " + "==" + DateTime.Now.ToString() + "==";
                Log.Information("End of api " + DateTime.Now);
                return Ok(new
                { Message = " Success!" });/* { Message = " Success!" });*/
            }
            catch (Exception ex)
            {
                //return BadRequest($"Could not find {ex.Message}");
                Log.Error($"Error occured on webinar method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                //return BadRequest(ex.Message);
                return BadRequest(new
                { Message = ex.Message + "------" + ex.StackTrace });
            }
            //finally
            //{
            //    semaphore.Release();
            //}
        }

        [HttpPost("WebinarPreEvent"), DisableRequestSizeLimit]
        public async Task<IActionResult> WebinarPreEvent(WebinarPayload formDataList)
        {

            string strMessage = string.Empty;

            //int timeInterval = 4000;
            //await Task.Delay(timeInterval);
            try
            {
                //semaphore = new SemaphoreSlim(0, 1);
                Log.Information("starting of api " + DateTime.Now);
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                Sheet UrlData = SheetHelper.GetSheetById(smartsheet, UI_URL);
                /////////////////////

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
                double TotalHonorariumAmount = 0;
                double TotalTravelAmount = 0;
                double TotalAccomodateAmount = 0;
                double TotalHCPLcAmount = 0;
                double TotalInviteesLcAmount = 0;
                double TotalExpenseAmount = 0;
                foreach (var formdata in formDataList.EventRequestExpenseSheet)
                {
                    string rowData = $"{addedExpencesNo}. {formdata.Expense} | AmountExcludingTax: {formdata.AmountExcludingTax}| Amount: {formdata.Amount} | {formdata.BtcorBte}";
                    addedExpences.AppendLine(rowData);
                    addedExpencesNo++;
                    var amount = SheetHelper.NumCheck(formdata.Amount);
                    TotalExpenseAmount = TotalExpenseAmount + amount;
                }
                string Expense = addedExpences.ToString();

                StringBuilder addedExpencesBTE = new();
                int addedExpencesNoBTE = 1;
                foreach (var formdata in formDataList.EventRequestExpenseSheet)
                {
                    if (formdata.BtcorBte.ToLower() == "bte")
                    {
                        string rowData = $"{addedExpencesNoBTE}. {formdata.Expense} | Amount: {formdata.Amount}";
                        addedExpencesBTE.AppendLine(rowData);
                        addedExpencesNoBTE++;
                    }
                }
                string BTEExpense = addedExpencesBTE.ToString();

                foreach (var formdata in formDataList.EventRequestHCPSlideKits)
                {
                    string rowData = $"{addedSlideKitDataNo}. {formdata.HcpName} | {formdata.SlideKitType} | Id :{formdata.SlideKitDocument}";
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
                    double HM = SheetHelper.NumCheck(formdata.HonarariumAmount);
                    double t = SheetHelper.NumCheck(formdata.Travel) + SheetHelper.NumCheck(formdata.Accomdation);

                    double roundedValue = Math.Round(t, 2);

                    string rowData = $"{addedHcpDataNo}. {formdata.HcpRole} |{formdata.HcpName} | Honr.Amt: {HM} |Trav.&Acc.Amt: {roundedValue} |Rationale : {formdata.Rationale}";
                    addedHcpData.AppendLine(rowData);
                    addedHcpDataNo++;
                    TotalHonorariumAmount = TotalHonorariumAmount + SheetHelper.NumCheck(formdata.HonarariumAmount);
                    TotalTravelAmount = TotalTravelAmount + SheetHelper.NumCheck(formdata.Travel);
                    TotalAccomodateAmount = TotalAccomodateAmount + SheetHelper.NumCheck(formdata.Accomdation);
                    TotalHCPLcAmount = TotalHCPLcAmount + SheetHelper.NumCheck(formdata.LocalConveyance);
                }
                string HCP = addedHcpData.ToString();


                double cc = TotalHCPLcAmount + TotalInviteesLcAmount;

                double totalAmount = TotalHonorariumAmount + TotalTravelAmount + TotalAccomodateAmount + TotalHCPLcAmount + TotalInviteesLcAmount + TotalExpenseAmount;

                double ss = TotalTravelAmount + TotalAccomodateAmount;

                double c = Math.Round(cc, 2);
                double total = Math.Round(totalAmount, 2);
                double s = Math.Round(ss, 2);

                Row? targetRow1 = UrlData.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == "Approver Pre Event URL"));
                Row? targetRow2 = UrlData.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == "Finance Treasury URL"));
                Row? targetRow4 = UrlData.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == "Initiator URL"));
                String Attachmentpaths = "";
                foreach (var p in formDataList.Webinar.Files)
                {
                    string[] words = p.Split(':');
                    string r = words[0];
                    string q = words[1];
                    string name = r.Split(".")[0];
                    string filePath = SheetHelper.SQlFileinsertion(q, name);
                    Attachmentpaths = Attachmentpaths + "," + filePath;
                }

                string MyConnection = configuration.GetSection("ConnectionStrings:mysql").Value;
                MySqlConnection MyConn = new MySqlConnection(MyConnection);
                MySqlCommand com = new MySqlCommand("WebinarPreevent", MyConn);
                com.CommandType = CommandType.StoredProcedure;
                com.Parameters.AddWithValue("@ApproverPreEventURL", targetRow1?.Cells[1].Value ?? "no url");
                com.Parameters.AddWithValue("@FinanceTreasuryURL", targetRow2?.Cells[1].Value ?? "no url");
                com.Parameters.AddWithValue("@InitiatorURL", targetRow4?.Cells[1].Value ?? "no url");
                com.Parameters.AddWithValue("@EventTopic", formDataList.Webinar.EventTopic);
                com.Parameters.AddWithValue("@EventType", formDataList.Webinar.EventType);
                com.Parameters.AddWithValue("@EventDate", formDataList.Webinar.EventDate);
                com.Parameters.AddWithValue("@StartTime", formDataList.Webinar.StartTime);
                com.Parameters.AddWithValue("@EndTime", formDataList.Webinar.EndTime);
                com.Parameters.AddWithValue("@MeetingType", formDataList.Webinar.Meeting_Type);
                com.Parameters.AddWithValue("@Brands", brand);
                com.Parameters.AddWithValue("@Expenses", Expense);
                com.Parameters.AddWithValue("@Panelists", HCP);
                com.Parameters.AddWithValue("@Invitees", Invitees);
                com.Parameters.AddWithValue("@MIPLInvitees", MenariniInvitees);
                com.Parameters.AddWithValue("@SlideKits", slideKit);
                com.Parameters.AddWithValue("@IsAdvanceRequired", formDataList.Webinar.IsAdvanceRequired);
                com.Parameters.AddWithValue("@EventOpen30days", formDataList.Webinar.EventOpen30days);
                com.Parameters.AddWithValue("@EventWithin7days", formDataList.Webinar.EventWithin7days);
                com.Parameters.AddWithValue("@InitiatorName", formDataList.Webinar.InitiatorName);
                com.Parameters.AddWithValue("@AdvanceAmount", SheetHelper.NumCheck(formDataList.Webinar.AdvanceAmount));
                com.Parameters.AddWithValue("@TotalExpenseBTC", SheetHelper.NumCheck(formDataList.Webinar.TotalExpenseBTC));
                com.Parameters.AddWithValue("@TotalExpenseBTE", SheetHelper.NumCheck(formDataList.Webinar.TotalExpenseBTE));
                com.Parameters.AddWithValue("@TotalHonorariumAmount", Math.Round(TotalHonorariumAmount, 2));
                com.Parameters.AddWithValue("@TotalTravelAmount", Math.Round(TotalTravelAmount, 2));
                com.Parameters.AddWithValue("@TotalTravelAccommodationAmount", s);
                com.Parameters.AddWithValue("@TotalAccommodationAmount", Math.Round(TotalAccomodateAmount, 2));
                com.Parameters.AddWithValue("@BudgetAmount", total);
                com.Parameters.AddWithValue("@TotalLocalConveyance", c);
                com.Parameters.AddWithValue("@TotalExpense", Math.Round(TotalExpenseAmount, 2));
                com.Parameters.AddWithValue("@InitiatorEmail", formDataList.Webinar.Initiator_Email);
                com.Parameters.AddWithValue("@RBMBM", formDataList.Webinar.RBMorBM);
                com.Parameters.AddWithValue("@SalesHead", formDataList.Webinar.Sales_Head);
                com.Parameters.AddWithValue("@SalesCoordinator", formDataList.Webinar.SalesCoordinatorEmail);
                com.Parameters.AddWithValue("@MarketingCoordinator", formDataList.Webinar.MarketingCoordinatorEmail);
                com.Parameters.AddWithValue("@MarketingHead", formDataList.Webinar.Marketing_Head);
                com.Parameters.AddWithValue("@Compliance", formDataList.Webinar.ComplianceEmail);
                com.Parameters.AddWithValue("@FinanceAccounts", formDataList.Webinar.FinanceAccountsEmail);
                com.Parameters.AddWithValue("@FinanceTreasury", formDataList.Webinar.Finance);
                com.Parameters.AddWithValue("@ReportingManager", formDataList.Webinar.ReportingManagerEmail);
                com.Parameters.AddWithValue("@1UpManager", formDataList.Webinar.FirstLevelEmail);
                com.Parameters.AddWithValue("@MedicalAffairsHead", formDataList.Webinar.MedicalAffairsEmail);
                com.Parameters.AddWithValue("@BTEExpenseDetails", formDataList.Webinar.BTEExpenseDetails);
                com.Parameters.AddWithValue("@AttachmentPaths", Attachmentpaths);
                com.Parameters.AddWithValue("@webinarRole", formDataList.Webinar.Role);
                await MyConn.OpenAsync();
                //com.ExecuteNonQuery();
                MySqlDataReader reader = com.ExecuteReader();
                String RefID = "";
                while (reader.Read())
                {
                    RefID = reader["ID"].ToString();
                }
                await MyConn.CloseAsync();


                await MyConn.OpenAsync();
                com = new MySqlCommand("SPEventRequestPanelDetails", MyConn);
                com.CommandType = CommandType.StoredProcedure;
                foreach (var formData in formDataList.EventRequestHcpRole)
                {
                    String PanelAttachmentpaths = "";
                    if (formData.IsUpload == "Yes")
                    {
                        int j = 1;
                        foreach (string p in formData.FilesToUpload)
                        {
                            string[] words = p.Split(':');
                            string r = words[0];
                            string q = words[1];
                            string name = r.Split(".")[0];
                            string filePath = SheetHelper.SQlFileinsertion(q, name);
                            PanelAttachmentpaths = PanelAttachmentpaths + "," + filePath;
                        }

                    }
                    com.Parameters.AddWithValue("@HcpRole", formData.HcpRole);
                    com.Parameters.AddWithValue("@MISCode", SheetHelper.MisCodeCheck(formData.MisCode));
                    com.Parameters.AddWithValue("@Travel", SheetHelper.NumCheck(formData.Travel));
                    com.Parameters.AddWithValue("@TotalSpend", SheetHelper.NumCheck(formData.FinalAmount));
                    com.Parameters.AddWithValue("@Accomodation", SheetHelper.NumCheck(formData.Accomdation));
                    com.Parameters.AddWithValue("@LocalConveyance", SheetHelper.NumCheck(formData.LocalConveyance));
                    com.Parameters.AddWithValue("@SpeakerCode", formData.SpeakerCode);
                    com.Parameters.AddWithValue("@TrainerCode", formData.TrainerCode);
                    com.Parameters.AddWithValue("@HonorariumRequired", formData.HonorariumRequired);
                    com.Parameters.AddWithValue("@AgreementAmount", SheetHelper.NumCheck(formData.AgreementAmount));
                    com.Parameters.AddWithValue("@HonorariumAmount", SheetHelper.NumCheck(formData.HonarariumAmount));
                    com.Parameters.AddWithValue("@Speciality", formData.Speciality);
                    com.Parameters.AddWithValue("@EventTopic", formDataList.Webinar.EventTopic);
                    com.Parameters.AddWithValue("@EventType", formDataList.Webinar.EventType);
                    com.Parameters.AddWithValue("@EventDateStart", formDataList.Webinar.EventDate);
                    com.Parameters.AddWithValue("@EventEndDate", formDataList.Webinar.EventEndDate);
                    com.Parameters.AddWithValue("@HCPName", formData.HcpName);
                    com.Parameters.AddWithValue("@PANcardname", formData.PanCardName);
                    com.Parameters.AddWithValue("@ExpenseType", formData.ExpenseType);
                    com.Parameters.AddWithValue("@BankAccountNumber", formData.BankAccountNumber);
                    com.Parameters.AddWithValue("@BankName", formData.BankName);
                    com.Parameters.AddWithValue("@IFSCCode", formData.IFSCCode);
                    com.Parameters.AddWithValue("@FCPADate", formData.Fcpadate);
                    com.Parameters.AddWithValue("@Currency", formData.Currency);
                    com.Parameters.AddWithValue("@HonorariumAmountExcludingTax", formData.HonarariumAmountExcludingTax);
                    com.Parameters.AddWithValue("@TravelExcludingTax", formData.TravelExcludingTax);
                    com.Parameters.AddWithValue("@AccomodationExcludingTax", formData.AccomdationExcludingTax);
                    com.Parameters.AddWithValue("@LocalConveyanceExcludingTax", formData.LocalConveyanceExcludingTax);
                    com.Parameters.AddWithValue("@LCBTCBTE", formData.LcBtcorBte);
                    com.Parameters.AddWithValue("@TravelBTCBTE", formData.TravelBtcorBte);
                    com.Parameters.AddWithValue("@AccomodationBTCBTE", formData.AccomodationBtcorBte);
                    com.Parameters.AddWithValue("@ModeofTravel", formData.TravelSelection);

                    if (formData.Currency == "Others")
                    {
                        com.Parameters.AddWithValue("@OtherCurrency", formData.OtherCurrencyType);
                    }
                    else
                    {
                        com.Parameters.AddWithValue("@OtherCurrency", "");
                    }
                    com.Parameters.AddWithValue("@BeneficiaryName", formData.BeneficiaryName);
                    com.Parameters.AddWithValue("@PanNumber", formData.PanNumber);

                    if (formData.HcpRole == "Others")
                    {
                        com.Parameters.AddWithValue("@OtherType", formData.OthersType);
                    }
                    else
                    {
                        com.Parameters.AddWithValue("@OtherType", "");
                    }

                    com.Parameters.AddWithValue("@Tier", formData.Tier);
                    com.Parameters.AddWithValue("@HCPType", formData.GOorNGO);
                    com.Parameters.AddWithValue("@PresentationDuration", SheetHelper.NumCheck(formData.PresentationDuration));
                    com.Parameters.AddWithValue("@PanelSessionPreparationDuration", SheetHelper.NumCheck(formData.PanelSessionPreperationDuration));
                    com.Parameters.AddWithValue("@PanelDiscussionDuration", SheetHelper.NumCheck(formData.PanelDisscussionDuration));
                    com.Parameters.AddWithValue("@QASessionDuration", SheetHelper.NumCheck(formData.QASessionDuration));
                    com.Parameters.AddWithValue("@BriefingSession", SheetHelper.NumCheck(formData.BriefingSession));
                    com.Parameters.AddWithValue("@TotalSessionHours", SheetHelper.NumCheck(formData.TotalSessionHours));
                    com.Parameters.AddWithValue("@Rationale", formData.Rationale);
                    com.Parameters.AddWithValue("@EventIdEventRequestId", RefID);
                    com.Parameters.AddWithValue("@AttachmentPaths", PanelAttachmentpaths);
                    com.ExecuteNonQuery();
                    com.Parameters.Clear();
                }


                await MyConn.CloseAsync();


                await MyConn.OpenAsync();
                com = new MySqlCommand("SPEventRequestsBrandsList", MyConn);
                com.CommandType = CommandType.StoredProcedure;
                foreach (var formdata in formDataList.RequestBrandsList)
                {
                    com.Parameters.AddWithValue("@Allocation", formdata.PercentAllocation);
                    com.Parameters.AddWithValue("@Brands", formdata.BrandName);
                    com.Parameters.AddWithValue("@ProjectID", formdata.ProjectId);
                    com.Parameters.AddWithValue("@EventIdEventRequestId", RefID);
                    com.ExecuteNonQuery();
                    com.Parameters.Clear();
                }
                await MyConn.CloseAsync();


                await MyConn.OpenAsync();
                com = new MySqlCommand("SPEventRequestInvitees", MyConn);
                com.CommandType = CommandType.StoredProcedure;
                foreach (var formdata in formDataList.EventRequestInvitees)
                {
                    com.Parameters.AddWithValue("@HCPName", formdata.InviteeName);
                    com.Parameters.AddWithValue("@Designation", formdata.Designation);
                    com.Parameters.AddWithValue("@EmployeeCode", formdata.EmployeeCode);
                    com.Parameters.AddWithValue("@LocalConveyance", formdata.LocalConveyance);
                    com.Parameters.AddWithValue("@BTCBTE", formdata.BtcorBte);
                    com.Parameters.AddWithValue("@LcAmount", SheetHelper.NumCheck(formdata.LcAmount));
                    com.Parameters.AddWithValue("@LcAmountExcludingTax", formdata.LcAmountExcludingTax);
                    com.Parameters.AddWithValue("@EventIdEventRequestId", RefID);
                    com.Parameters.AddWithValue("@InviteeSource", formdata.InviteedFrom);
                    com.Parameters.AddWithValue("@HCPType", formdata.HCPType);
                    com.Parameters.AddWithValue("@Speciality", formdata.Speciality);
                    com.Parameters.AddWithValue("@MISCode", SheetHelper.MisCodeCheck(formdata.MISCode));
                    com.Parameters.AddWithValue("@EventTopic", formDataList.Webinar.EventTopic);
                    com.Parameters.AddWithValue("@EventType", formDataList.Webinar.EventType);
                    com.Parameters.AddWithValue("@EventDateStart", formDataList.Webinar.EventDate);
                    com.Parameters.AddWithValue("@EventEndDate", formDataList.Webinar.EventDate);
                    com.ExecuteNonQuery();
                    com.Parameters.Clear();
                }
                await MyConn.CloseAsync();


                await MyConn.OpenAsync();
                com = new MySqlCommand("SPEventRequestHCPSlideKitDetails", MyConn);
                com.CommandType = CommandType.StoredProcedure;
                String SlidekitsAttachent = "";
                foreach (var formdata in formDataList.EventRequestHCPSlideKits)
                {
                    if (formdata.IsUpload == "Yes")
                    {
                        int j = 1;
                        foreach (string p in formdata.FilesToUpload)
                        {
                            string[] words = p.Split(':');
                            string r = words[0];
                            string q = words[1];
                            string name = r.Split(".")[0];
                            string filePath = SheetHelper.SQlFileinsertion(q, name);
                            SlidekitsAttachent = SlidekitsAttachent + "," + filePath;
                        }
                    }
                    com.Parameters.AddWithValue("@MIS", SheetHelper.MisCodeCheck(formdata.MIS));
                    com.Parameters.AddWithValue("@SlideKitType", formdata.SlideKitType);
                    com.Parameters.AddWithValue("@SlideKitDocument", formdata.SlideKitDocument);
                    com.Parameters.AddWithValue("@EventIdEventRequestId", RefID);
                    com.Parameters.AddWithValue("@AttachmentPaths", SlidekitsAttachent);
                    com.ExecuteNonQuery();
                    com.Parameters.Clear();
                }

                MyConn.CloseAsync();


                MyConn.Open();
                com = new MySqlCommand("SPEventRequestExpensesSheet", MyConn);
                com.CommandType = CommandType.StoredProcedure;

                foreach (var formdata in formDataList.EventRequestExpenseSheet)
                {
                    com.Parameters.AddWithValue("@Expense", formdata.Expense);
                    com.Parameters.AddWithValue("@EventIdEventRequestID", RefID);
                    com.Parameters.AddWithValue("@AmountExcludingTaxq", formdata.AmountExcludingTax);
                    com.Parameters.AddWithValue("@AmountExcludingTax", formdata.ExcludingTaxAmount);
                    com.Parameters.AddWithValue("@Amount", SheetHelper.NumCheck(formdata.Amount));
                    com.Parameters.AddWithValue("@BTCBTE", formdata.BtcorBte);
                    com.Parameters.AddWithValue("@BudgetAmount", SheetHelper.NumCheck(formdata.BudgetAmount));
                    com.Parameters.AddWithValue("@BTCAmount", SheetHelper.NumCheck(formdata.BtcAmount));
                    com.Parameters.AddWithValue("@BTEAmount", SheetHelper.NumCheck(formdata.BteAmount));
                    com.Parameters.AddWithValue("@EventTopic", formDataList.Webinar.EventTopic);
                    com.Parameters.AddWithValue("@EventType", formDataList.Webinar.EventType);
                    com.Parameters.AddWithValue("@EventDateStart", formDataList.Webinar.EventDate);
                    com.Parameters.AddWithValue("@EventEndDate", formDataList.Webinar.EventDate);
                    com.ExecuteNonQuery();
                    com.Parameters.Clear();
                }
                MyConn.CloseAsync();

                if (formDataList.Webinar.EventOpen30days == "Yes" || formDataList.Webinar.EventWithin7days == "Yes" || formDataList.Webinar.FB_Expense_Excluding_Tax == "Yes" || formDataList.Webinar.IsDeviationUpload == "Yes")
                {
                    List<string> DeviationNames = new List<string>();
                    foreach (var p in formDataList.Webinar.DeviationDetails)
                    {
                        string[] words = p.DeviationFile.Split(':')[0].Split("*");
                        string r = words[1];
                        DeviationNames.Add(r);
                    }
                    await MyConn.OpenAsync();
                    com = new MySqlCommand("SPDeviation_Process", MyConn);
                    com.CommandType = CommandType.StoredProcedure;

                    foreach (var pp in formDataList.Webinar.DeviationDetails)
                    {
                        foreach (var deviationname in DeviationNames)
                        {
                            string file = deviationname.Split(".")[0];
                            string DeviationAttachmentpath = "";
                            if (pp.DeviationFile.Split(':')[0].Split("*")[1] == deviationname)
                            {
                                try
                                {
                                    foreach (var p in formDataList.Webinar.DeviationDetails)
                                    {
                                        string[] nameSplit = p.DeviationFile.Split("*");
                                        string[] words = nameSplit[1].Split(':');
                                        string r = words[0];
                                        string q = words[1];
                                        if (deviationname == r)
                                        {
                                            string name = nameSplit[0];
                                            string filePath = SheetHelper.SQlFileinsertion(q, name);
                                            DeviationAttachmentpath = DeviationAttachmentpath + "," + filePath;
                                            //Attachment attachmentinmain = await ApiCalls.AddAttachmentsToSheet(smartsheet, sheet1, addedRows[0], filePath);
                                        }
                                    }

                                    com.Parameters.AddWithValue("@EventIdEventRequestId", RefID);
                                    com.Parameters.AddWithValue("@EventTopic", formDataList.Webinar.EventTopic);
                                    com.Parameters.AddWithValue("@EventType", formDataList.Webinar.EventType);
                                    com.Parameters.AddWithValue("@EventDate", formDataList.Webinar.EventDate);
                                    com.Parameters.AddWithValue("@StartTime", formDataList.Webinar.StartTime);
                                    com.Parameters.AddWithValue("@EndTime", formDataList.Webinar.EndTime);
                                    com.Parameters.AddWithValue("@MISCode", SheetHelper.MisCodeCheck(pp.MisCode));
                                    com.Parameters.AddWithValue("@HCPName", pp.HcpName);
                                    com.Parameters.AddWithValue("@HonorariumAmount", pp.HonorariumAmountExcludingTax);
                                    com.Parameters.AddWithValue("@TravelAccommodationAmount", pp.TravelorAccomodationAmountExcludingTax);
                                    com.Parameters.AddWithValue("@OtherExpenses", pp.OtherExpenseAmountExcludingTax);


                                    if (file == "30DaysDeviationFile")
                                    {
                                        com.Parameters.AddWithValue("@DeviationType", configuration.GetSection("DeviationNamesInPreEvent:30DaysDeviationFile").Value);
                                        com.Parameters.AddWithValue("@EventOpen45days", formDataList.Webinar.EventOpen30days);
                                        com.Parameters.AddWithValue("@OutstandingEvents", SheetHelper.NumCheck(formDataList.Webinar.EventOpen30dayscount));
                                    }
                                    else
                                    {
                                        com.Parameters.AddWithValue("@EventOpen45days", "");
                                        com.Parameters.AddWithValue("@OutstandingEvents", "");
                                    }
                                    if (file == "7DaysDeviationFile")
                                    {
                                        com.Parameters.AddWithValue("@DeviationType", configuration.GetSection("DeviationNamesInPreEvent:5DaysDeviationFile").Value);
                                        com.Parameters.AddWithValue("@EventWithin5days", formDataList.Webinar.EventWithin7days);
                                    }
                                    else
                                    {
                                        com.Parameters.AddWithValue("@EventWithin5days", "");
                                    }
                                    if (file == "ExpenseExcludingTax")
                                    {
                                        com.Parameters.AddWithValue("@DeviationType", configuration.GetSection("DeviationNamesInPreEvent:ExpenseExcludingTax").Value);
                                        com.Parameters.AddWithValue("@PREExpenseExcludingTax", formDataList.Webinar.FB_Expense_Excluding_Tax);
                                    }
                                    else
                                    {
                                        com.Parameters.AddWithValue("@PREExpenseExcludingTax", "");
                                    }
                                    if (file.Contains("Travel_Accomodation3LExceededFile"))
                                    {
                                        com.Parameters.AddWithValue("@DeviationType", configuration.GetSection("DeviationNamesInPreEvent:Travel_Accomodation3LExceededFile").Value);
                                        com.Parameters.AddWithValue("@TravelAccomodationExceededTrigger", "Yes");
                                    }
                                    else
                                    {
                                        com.Parameters.AddWithValue("@TravelAccomodationExceededTrigger", "");
                                    }
                                    if (file.Contains("TrainerHonorarium12LExceededFile"))
                                    {
                                        com.Parameters.AddWithValue("@DeviationType", configuration.GetSection("DeviationNamesInPreEvent:TrainerHonorarium12LExceededFile").Value);
                                        com.Parameters.AddWithValue("@TrainerHonorariumExceededTrigger", "Yes");
                                    }
                                    else
                                    {
                                        com.Parameters.AddWithValue("@TrainerHonorariumExceededTrigger", "");
                                    }
                                    if (file.Contains("HCPHonorarium6LExceededFile"))
                                    {
                                        com.Parameters.AddWithValue("@DeviationType", configuration.GetSection("DeviationNamesInPreEvent:HCPHonorarium6LExceededFile").Value);
                                        com.Parameters.AddWithValue("@HCPHonorariumExceededTrigger", "Yes");
                                    }
                                    else
                                    {
                                        com.Parameters.AddWithValue("@HCPHonorariumExceededTrigger", "");
                                    }
                                    com.Parameters.AddWithValue("@SalesHead", formDataList.Webinar.Sales_Head);
                                    com.Parameters.AddWithValue("@FinanceHead", formDataList.Webinar.FinanceHead);
                                    com.Parameters.AddWithValue("@InitiatorName", formDataList.Webinar.InitiatorName);
                                    com.Parameters.AddWithValue("@InitiatorEmail", formDataList.Webinar.Initiator_Email);
                                    com.Parameters.AddWithValue("@SalesCoordinator", formDataList.Webinar.SalesCoordinatorEmail);
                                    com.Parameters.AddWithValue("@AttachmentPaths", DeviationAttachmentpath);
                                    com.Parameters.AddWithValue("@EndDate", "");
                                    com.ExecuteNonQuery();
                                    com.Parameters.Clear();
                                }
                                catch (Exception ex)
                                {
                                    return BadRequest(new
                                    {
                                        Message = ex.Message + "------" + ex.StackTrace
                                    });

                                }

                            }

                        }

                    }
                }

                //Row addedrow = addedRows[0];
                //long ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role");
                //Cell UpdateB = new Cell { ColumnId = ColumnId, Value = formDataList.Webinar.Role };
                //Row updateRows = new Row { Id = addedrow.Id, Cells = new Cell[] { UpdateB } };
                //Cell? cellsToUpdate = addedrow.Cells.FirstOrDefault(c => c.ColumnId == ColumnId);
                //if (cellsToUpdate != null) { cellsToUpdate.Value = formDataList.Webinar.Role; }

                //strMessage += "==Before adding Role to WebSheet " + "==" + DateTime.Now.ToString() + "==";
                //await Task.Run(() => ApiCalls.UpdateRole(smartsheet, sheet1, updateRows)); //smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRows }));
                //strMessage += "==Before adding Role to WebSheet " + "==" + DateTime.Now.ToString() + "==";
                //Log.Information("End of api " + DateTime.Now);
                MyConn.CloseAsync();
                //return Ok(new
                //{ Message = " Success!" });/* { Message = " Success!" });*/
                DateTime currentDate = DateTime.Now;
                return Ok(new
                {
                    Message = $"Thank you. Your event creation request has been received. " +
                "You should receive a confirmation email with the details of your event after a few minutes."
                });

            }
            catch (Exception ex)
            {
                //return BadRequest($"Could not find {ex.Message}");
                Log.Error($"Error occured on webinar method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                //return BadRequest(ex.Message);
                return BadRequest(new
                { Message = ex.Message + "------" + ex.StackTrace });
            }
            //finally
            //{
            //    semaphore.Release();
            //}
        }

        [HttpPost("HCPConsultantPreEvent"), DisableRequestSizeLimit]
        public IActionResult HCPConsultantPreEvent(HCPConsultantPayload formDataList)
        {
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
            Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
            Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
            Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);

            StringBuilder addedBrandsData = new();
            StringBuilder addedHcpData = new();
            StringBuilder addedExpences = new();

            int addedHcpDataNo = 1;
            int addedBrandsDataNo = 1;
            int addedExpencesNo = 1;

            double TotalHonorariumAmount = 0;
            double TotalTravelAmount = 0;
            double TotalAccomodateAmount = 0;
            double TotalHCPLcAmount = 0;
            double TotalInviteesLcAmount = 0;
            double TotalExpenseAmount = 0;
            double TotalRegstAmount = 0;



            string EventOpen30Days = !string.IsNullOrEmpty(formDataList.HcpConsultant.EventOpen30days) ? "Yes" : "No";
            string EventWithin7Days = !string.IsNullOrEmpty(formDataList.HcpConsultant.EventWithin7days) ? "Yes" : "No";
            string BrouchereUpload = !string.IsNullOrEmpty(formDataList.HcpConsultant.BrochureFile) ? "Yes" : "No";
            string FCPA = !string.IsNullOrEmpty(formDataList.HcpConsultant.FcpaFile) ? "Yes" : "No";


            foreach (var formdata in formDataList.ExpenseSheet)
            {
                string rowData = $"{addedExpencesNo}. {formdata.Expense} | RegstAmount: {formdata.RegstAmount}| {formdata.BTC_BTE}";
                addedExpences.AppendLine(rowData);
                addedExpencesNo++;

                double amount = SheetHelper.NumCheck(formdata.ExpenseAmount);
                double regst = SheetHelper.NumCheck(formdata.RegstAmount);
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
                double HM = SheetHelper.NumCheck(formdata.RegistrationAmount);
                double t = SheetHelper.NumCheck(formdata.TravelAmount) + SheetHelper.NumCheck(formdata.AccomAmount);
                string rowData = $"{addedHcpDataNo}. {formdata.MisCode} |{formdata.HcpName} | Regst.Amt: {HM} |Trav.&Acc.Amt: {t} |Rationale :{formdata.Rationale}";

                addedHcpData.AppendLine(rowData);
                addedHcpDataNo++;
                TotalHonorariumAmount = TotalHonorariumAmount + SheetHelper.NumCheck(formdata.RegistrationAmount);
                TotalTravelAmount = TotalTravelAmount + SheetHelper.NumCheck(formdata.TravelAmount);
                TotalAccomodateAmount = TotalAccomodateAmount + SheetHelper.NumCheck(formdata.AccomAmount);
                TotalHCPLcAmount = TotalHCPLcAmount + SheetHelper.NumCheck(formdata.LcAmount);
            }
            string HCP = addedHcpData.ToString();
            double c = TotalHCPLcAmount + TotalInviteesLcAmount;
            double total = TotalHonorariumAmount + TotalTravelAmount + TotalAccomodateAmount + TotalHCPLcAmount + TotalInviteesLcAmount;

            double s = TotalTravelAmount + TotalAccomodateAmount;

            StringBuilder addedExpencesBTE = new();
            int addedExpencesNoBTE = 1;
            foreach (var formdata in formDataList.ExpenseSheet)
            {
                if (formdata.BTC_BTE.ToLower() == "bte")
                {
                    string rowData = $"{addedExpencesNoBTE}. {formdata.Expense} | Amount: {formdata.ExpenseAmount}";
                    addedExpencesBTE.AppendLine(rowData);
                    addedExpencesNoBTE++;
                }
            }
            string BTEExpense = addedExpencesBTE.ToString();
            try
            {
                var newRow = new Row();
                newRow.Cells = new List<Cell>();
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.HcpConsultant.EventTopic });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Type"), Value = formDataList.HcpConsultant.EventType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Date"), Value = formDataList.HcpConsultant.EventDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "End Date"), Value = formDataList.HcpConsultant.EventEndDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Venue Name"), Value = formDataList.HcpConsultant.VenueName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Start Time"), Value = formDataList.HcpConsultant.StartTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "End Time"), Value = formDataList.HcpConsultant.EndTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sponsorship Society Name"), Value = formDataList.HcpConsultant.SponsorshipSocietyName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Venue Country"), Value = formDataList.HcpConsultant.Country });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "IsAdvanceRequired"), Value = formDataList.HcpConsultant.IsAdvanceRequired });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Advance Amount"), Value = SheetHelper.NumCheck(formDataList.HcpConsultant.AdvanceAmount) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Brands"), Value = brand });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Expenses"), Value = Expense });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Panelists"), Value = HCP });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Name"), Value = formDataList.HcpConsultant.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total HCP Registration Amount"), Value = TotalHonorariumAmount });
                //newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total HCP Registration Amount"), Value = TotalExpenseAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel Amount"), Value = TotalTravelAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel & Accommodation Amount"), Value = s });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Accommodation Amount"), Value = TotalAccomodateAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Local Conveyance"), Value = c });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense"), Value = TotalExpenseAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Budget Amount"), Value = total });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Email"), Value = formDataList.HcpConsultant.Initiator_Email });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "RBM/BM"), Value = formDataList.HcpConsultant.RBMorBM });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Head"), Value = formDataList.HcpConsultant.Sales_Head });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Coordinator"), Value = formDataList.HcpConsultant.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Marketing Coordinator"), Value = formDataList.HcpConsultant.MarketingCoordinatorEmail });
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
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "BTE Expense Details"), Value = BTEExpense });

                IList<Row> addedRows = smartsheet.SheetResources.RowResources.AddRows(sheet1.Id.Value, new Row[] { newRow });

                long eventIdColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                Cell? eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                string val = eventIdCell.DisplayValue;

                if (BrouchereUpload == "Yes")
                {


                    string filename = formDataList.HcpConsultant.BrochureFile.Split(":")[0].Split(".")[0];
                    string filePath = SheetHelper.testingFile(formDataList.HcpConsultant.BrochureFile.Split(":")[1], filename);


                    Row addedRow = addedRows[0];
                    Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                            sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");


                    if (System.IO.File.Exists(filePath))
                    {
                        System.IO.File.Delete(filePath);
                    }
                }



                if (EventOpen30Days == "Yes" || EventWithin7Days == "Yes" || formDataList.HcpConsultant.AggregateDeviation == "Yes")
                {
                    List<string> DeviationNames = new List<string>
                    {
                        $"30DaysDeviationFile:{formDataList.HcpConsultant.EventOpen30days}",
                        $"7DaysDeviationFile:{formDataList.HcpConsultant.EventWithin7days}",
                        $"AgregateSpendDeviationFile:{formDataList.HcpConsultant.AggregateDeviationFiles}"
                    };
                    string eventId = val;
                    foreach (var name in DeviationNames)
                    {
                        //string[] nameSplit = name.Split(":");
                        string[] y = name.Split(':');
                        //string[] y = name.Split(':');
                        string fn = y[0];
                        string bs = y[1];

                        if (bs != "")
                        {
                            try
                            {

                                Row newRow7 = new()
                                {
                                    Cells = new List<Cell>()
                                };

                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.HcpConsultant.EventTopic });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Type"), Value = formDataList.HcpConsultant.EventType });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Date"), Value = formDataList.HcpConsultant.EventDate });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Start Time"), Value = formDataList.HcpConsultant.StartTime });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "End Time"), Value = formDataList.HcpConsultant.EndTime });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Venue Name"), Value = formDataList.HcpConsultant.VenueName });
                                if (fn == "30DaysDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = EventOpen30Days });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:30DaysDeviationFile").Value });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = SheetHelper.NumCheck(formDataList.HcpConsultant.EventOpen30dayscount) });

                                }
                                else if (fn == "7DaysDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = EventWithin7Days });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:5DaysDeviationFile").Value });
                                }
                                else if (fn == "AgregateSpendDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HCP exceeds 5,00,000 Trigger"), Value = formDataList.HcpConsultant.AggregateDeviation });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:AgregateSpendDeviationFile5L").Value });
                                }
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.HcpConsultant.Sales_Head });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formDataList.HcpConsultant.FinanceHead });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Name"), Value = formDataList.HcpConsultant.InitiatorName });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.HcpConsultant.Initiator_Email });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Coordinator"), Value = formDataList.HcpConsultant.SalesCoordinatorEmail });

                                var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });

                                if (fn == "AgregateSpendDeviationFile")
                                {
                                    foreach (var file in formDataList.HcpConsultant.AggregateDeviationFiles)
                                    {
                                        //string filename = nameSplit[1].Split(".")[0];
                                        string bFile = file.Split("*")[1];
                                        string bName = file.Split("*")[0].Split(".")[0];
                                        string filePath = SheetHelper.testingFile(bFile, bName);
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
                                    //string filename = nameSplit[1].Split(".")[0];
                                    string bFile = bs.Split("*")[1];
                                    string bName = bs.Split("*")[0].Split(".")[0];
                                    string filePath = SheetHelper.testingFile(bFile, bName);
                                    //string filePath = SheetHelper.testingFile(bs, filename);

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
                                Log.Error($"Error occured on AllPreEventsController Attachementfile method {ex.Message} at {DateTime.Now}");
                                Log.Error(ex.StackTrace);
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
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "MISCode"), Value = SheetHelper.MisCodeCheck(formData.MisCode) });
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
                            string filePath = SheetHelper.testingFile(q, name);
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
                //return BadRequest($"Could not find {ex.Message}");
                Log.Error($"Error occured on AllPreEventsController Attachementfile method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(ex.Message);
            }
        }

        [HttpPost("StallFabricationPreEventSmartSheet"), DisableRequestSizeLimit]
        public IActionResult StallFabricationPreEventSmartSheet(AllStallFabrication formDataList)
        {
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
            Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
            Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);

            Sheet UrlData = SheetHelper.GetSheetById(smartsheet, UI_URL);

            StringBuilder addedBrandsData = new StringBuilder();
            StringBuilder addedExpences = new StringBuilder();
            int addedBrandsDataNo = 1;
            int addedExpencesNo = 1;
            double TotalExpenseAmount = 0;
            string uploadDeviationForTableContainsData = !string.IsNullOrEmpty(formDataList.StallFabrication.TableContainsDataUpload) ? "Yes" : "No";
            string EventWithin7Days = !string.IsNullOrEmpty(formDataList.StallFabrication.EventWithin7daysUpload) ? "Yes" : "No";
            string BrouchereUpload = !string.IsNullOrEmpty(formDataList.StallFabrication.EventBrouchereUpload) ? "Yes" : "No";
            string InvoiceUpload = !string.IsNullOrEmpty(formDataList.StallFabrication.Invoice_QuotationUpload) ? "Yes" : "No";

            foreach (var formdata in formDataList.ExpenseSheets)
            {
                string rowData = $"{addedExpencesNo}. {formdata.Expense} | AmountExcludingTax: {formdata.AmountExcludingTax}| Amount: {formdata.Amount} | {formdata.BtcorBte}";
                addedExpences.AppendLine(rowData);
                addedExpencesNo++;
                var amount = SheetHelper.NumCheck(formdata.Amount);
                TotalExpenseAmount = TotalExpenseAmount + amount;
            }
            string Expense = addedExpences.ToString();

            StringBuilder addedExpencesBTE = new();
            int addedExpencesNoBTE = 1;
            foreach (var formdata in formDataList.ExpenseSheets)
            {
                if (formdata.BtcorBte.ToLower() == "bte")
                {
                    string rowData = $"{addedExpencesNoBTE}. {formdata.Expense} | Amount: {formdata.Amount}";
                    addedExpencesBTE.AppendLine(rowData);
                    addedExpencesNoBTE++;
                }
            }
            string BTEExpense = addedExpencesBTE.ToString();


            foreach (var formdata in formDataList.EventBrands)
            {
                string rowData = $"{addedBrandsDataNo}. {formdata.BrandName} | {formdata.ProjectId} | {formdata.PercentAllocation}";
                addedBrandsData.AppendLine(rowData);
                addedBrandsDataNo++;
            }
            string brand = addedBrandsData.ToString();



            double total = TotalExpenseAmount;
            Dictionary<string, long> Sheet1columns = new();
            foreach (var column in sheet1.Columns)
            {
                Sheet1columns.Add(column.Title, (long)column.Id);
            }
            try
            {
                Row newRow = new()
                {
                    Cells = new List<Cell>()
                };
                Row? targetRow1 = UrlData.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == "Approver Pre Event URL"));
                Row? targetRow2 = UrlData.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == "Finance Treasury URL"));
                Row? targetRow4 = UrlData.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == "Initiator URL"));

                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Approver Pre Event URL"], Value = targetRow1?.Cells[1].Value ?? "no url" });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Finance Treasury URL"], Value = targetRow2?.Cells[1].Value ?? "no url" });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Initiator URL"], Value = targetRow4?.Cells[1].Value ?? "no url" });

                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Event Topic"], Value = formDataList.StallFabrication.EventName });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Event Type"], Value = formDataList.StallFabrication.EventType });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Event Date"], Value = formDataList.StallFabrication.StartDate });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["End Date"], Value = formDataList.StallFabrication.EndDate });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Class III Event Code"], Value = formDataList.StallFabrication.Class_III_EventCode });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Brands"], Value = brand });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Expenses"], Value = Expense });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Initiator Name"], Value = formDataList.StallFabrication.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Total Expense"], Value = TotalExpenseAmount });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Budget Amount"], Value = total });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["IsAdvanceRequired"], Value = formDataList.StallFabrication.IsAdvanceRequired });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Advance Amount"], Value = SheetHelper.NumCheck(formDataList.StallFabrication.AdvanceAmount) });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Initiator Email"], Value = formDataList.StallFabrication.Initiator_Email });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["RBM/BM"], Value = formDataList.StallFabrication.RBMorBM });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Sales Head"], Value = formDataList.StallFabrication.Sales_Head });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Sales Coordinator"], Value = formDataList.StallFabrication.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Marketing Coordinator"], Value = formDataList.StallFabrication.MarketingCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Marketing Head"], Value = formDataList.StallFabrication.Marketing_Head });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Compliance"], Value = formDataList.StallFabrication.ComplianceEmail });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Finance Accounts"], Value = formDataList.StallFabrication.FinanceAccountsEmail });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Finance Treasury"], Value = formDataList.StallFabrication.Finance });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Reporting Manager"], Value = formDataList.StallFabrication.ReportingManagerEmail });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["1 Up Manager"], Value = formDataList.StallFabrication.FirstLevelEmail });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Medical Affairs Head"], Value = formDataList.StallFabrication.MedicalAffairsEmail });
                //newRow.Cells.Add(new Cell { ColumnId Sheet1columns[1, "Role"), Value = formDataList.StallFabrication.Role });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns[" Total Expense BTC"], Value = SheetHelper.NumCheck(formDataList.StallFabrication.TotalExpenseBTC) });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["Total Expense BTE"], Value = SheetHelper.NumCheck(formDataList.StallFabrication.TotalExpenseBTE) });
                newRow.Cells.Add(new Cell { ColumnId = Sheet1columns["BTE Expense Details"], Value = BTEExpense });

                IList<Row> addedRows = smartsheet.SheetResources.RowResources.AddRows(sheet1.Id.Value, new Row[] { newRow });
                long eventIdColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                Cell? eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                string val = eventIdCell.DisplayValue;



                if (BrouchereUpload == "Yes")
                {
                    string name = formDataList.StallFabrication.EventBrouchereUpload.Split(":")[0].Split(".")[0];
                    string filePath = SheetHelper.testingFile(formDataList.StallFabrication.EventBrouchereUpload.Split(":")[1], name);

                    Row addedRow = addedRows[0];

                    Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                            sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");


                    if (System.IO.File.Exists(filePath))
                    {
                        SheetHelper.DeleteFile(filePath);
                    }
                }
                if (InvoiceUpload == "Yes")
                {
                    string name = formDataList.StallFabrication.Invoice_QuotationUpload.Split(":")[0].Split(".")[0];

                    string filePath = SheetHelper.testingFile(formDataList.StallFabrication.Invoice_QuotationUpload.Split(":")[1], name);

                    Row addedRow = addedRows[0];

                    Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                            sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");


                    if (System.IO.File.Exists(filePath))
                    {
                        SheetHelper.DeleteFile(filePath);
                    }
                }


                if (formDataList.StallFabrication.IsDeviationUpload == "Yes")
                {
                    string eventId = val;
                    List<string> list = new()
                    {
                        "30days","7days"
                    };
                    foreach (var item in list)
                    {
                        if (uploadDeviationForTableContainsData == "Yes" && item == "30days")
                        {
                            try
                            {
                                Row newRow7 = new()
                                {
                                    Cells = new List<Cell>()
                                };
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.StallFabrication.EventName });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Type"), Value = formDataList.StallFabrication.EventType });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Date"), Value = formDataList.StallFabrication.StartDate });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "End Date"), Value = formDataList.StallFabrication.EndDate });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = uploadDeviationForTableContainsData });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.StallFabrication.Sales_Head });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Name"), Value = formDataList.StallFabrication.InitiatorName });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.StallFabrication.Initiator_Email });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:30DaysDeviationFile").Value });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = SheetHelper.NumCheck(formDataList.StallFabrication.EventOpen30dayscount) });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Coordinator"), Value = formDataList.StallFabrication.SalesCoordinatorEmail });

                                IList<Row> addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });

                                string name = formDataList.StallFabrication.TableContainsDataUpload.Split(":")[0].Split(".")[0];
                                string filePath = SheetHelper.testingFile(formDataList.StallFabrication.TableContainsDataUpload.Split(":")[1], name);

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
                            catch (Exception ex)
                            {
                                Log.Error($"Error occured on AllPreEventsController Attachementfile method {ex.Message} at {DateTime.Now}");
                                Log.Error(ex.StackTrace);
                                return BadRequest(ex.Message);
                                //return BadRequest(ex.Message);
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
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Type"), Value = formDataList.StallFabrication.EventType });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Date"), Value = formDataList.StallFabrication.StartDate });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "End Date"), Value = formDataList.StallFabrication.EndDate });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = EventWithin7Days });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.StallFabrication.Sales_Head });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Name"), Value = formDataList.StallFabrication.InitiatorName });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.StallFabrication.Initiator_Email });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:5DaysDeviationFile").Value });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Coordinator"), Value = formDataList.StallFabrication.SalesCoordinatorEmail });


                                IList<Row> addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });

                                string name = formDataList.StallFabrication.EventWithin7daysUpload.Split(":")[0].Split(".")[0];

                                string filePath = SheetHelper.testingFile(formDataList.StallFabrication.EventWithin7daysUpload.Split(":")[1], name);



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
                            catch (Exception ex)
                            {
                                //return BadRequest(ex.Message);
                                Log.Error($"Error occured on AllPreEventsController Attachementfile method {ex.Message} at {DateTime.Now}");
                                Log.Error(ex.StackTrace);
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

                DateTime currentDate = DateTime.Now;
                return Ok(new
                {
                    Message = $"Thank you. Your event creation request has been received. " +
                "You should receive a confirmation email with the details of your event after a few minutes."
                });




            }



            catch (Exception ex)
            {
                //return BadRequest($"Could not find {ex.Message}");
                Log.Error($"Error occured on AllPreEventsController Attachementfile method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(ex.Message);
            }


        }

        [HttpPost("StallFabricationPreEvent"), DisableRequestSizeLimit]
        public async Task<IActionResult> StallFabricationPreEvent(AllStallFabrication formDataList)
        {

            string strMessage = string.Empty;


            try
            {
                Log.Information("starting of api " + DateTime.Now);
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                Sheet UrlData = SheetHelper.GetSheetById(smartsheet, UI_URL);
                StringBuilder addedBrandsData = new StringBuilder();
                StringBuilder addedExpences = new StringBuilder();
                int addedBrandsDataNo = 1;
                int addedExpencesNo = 1;
                double TotalExpenseAmount = 0;

                foreach (var formdata in formDataList.ExpenseSheets)
                {
                    string rowData = $"{addedExpencesNo}. {formdata.Expense} | AmountExcludingTax: {formdata.AmountExcludingTax}| Amount: {formdata.Amount} | {formdata.BtcorBte}";
                    addedExpences.AppendLine(rowData);
                    addedExpencesNo++;
                    var amount = SheetHelper.NumCheck(formdata.Amount);
                    TotalExpenseAmount = TotalExpenseAmount + amount;
                }
                string Expense = addedExpences.ToString();

                StringBuilder addedExpencesBTE = new();
                int addedExpencesNoBTE = 1;
                foreach (var formdata in formDataList.ExpenseSheets)
                {
                    if (formdata.BtcorBte.ToLower() == "bte")
                    {
                        string rowData = $"{addedExpencesNoBTE}. {formdata.Expense} | Amount: {formdata.Amount}";
                        addedExpencesBTE.AppendLine(rowData);
                        addedExpencesNoBTE++;
                    }
                }
                string BTEExpense = addedExpencesBTE.ToString();


                foreach (var formdata in formDataList.EventBrands)
                {
                    string rowData = $"{addedBrandsDataNo}. {formdata.BrandName} | {formdata.ProjectId} | {formdata.PercentAllocation}";
                    addedBrandsData.AppendLine(rowData);
                    addedBrandsDataNo++;
                }
                string brand = addedBrandsData.ToString();


                string uploadDeviationForTableContainsData = !string.IsNullOrEmpty(formDataList.StallFabrication.TableContainsDataUpload) ? "Yes" : "No";
                string EventWithin7Days = !string.IsNullOrEmpty(formDataList.StallFabrication.EventWithin7daysUpload) ? "Yes" : "No";
                string BrouchereUpload = !string.IsNullOrEmpty(formDataList.StallFabrication.EventBrouchereUpload) ? "Yes" : "No";
                string InvoiceUpload = !string.IsNullOrEmpty(formDataList.StallFabrication.Invoice_QuotationUpload) ? "Yes" : "No";

                double total = TotalExpenseAmount;


                Row? targetRow1 = UrlData.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == "Approver Pre Event URL"));
                Row? targetRow2 = UrlData.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == "Finance Treasury URL"));
                Row? targetRow4 = UrlData.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == "Initiator URL"));

                String Attachmentpaths = "";

                List<string> attachmentsBase64 = new List<string>();

                //if (uploadDeviationForTableContainsData == "Yes") attachmentsBase64.Add(formDataList.StallFabrication.TableContainsDataUpload);
                //if (EventWithin7Days == "yes") attachmentsBase64.Add(formDataList.StallFabrication.EventWithin7daysUpload);
                if (BrouchereUpload == "Yes") attachmentsBase64.Add(formDataList.StallFabrication.EventBrouchereUpload);
                if (InvoiceUpload == "Yes") attachmentsBase64.Add(formDataList.StallFabrication.Invoice_QuotationUpload);


                foreach (var p in attachmentsBase64)
                {
                    string[] words = p.Split(':');
                    string r = words[0];
                    string q = words[1];
                    string name = r.Split(".")[0];
                    string filePath = SheetHelper.SQlFileinsertion(q, name);
                    Attachmentpaths = Attachmentpaths + "," + filePath;
                }

                string MyConnection = configuration.GetSection("ConnectionStrings:mysql").Value;
                MySqlConnection MyConn = new MySqlConnection(MyConnection);
                MySqlCommand com = new MySqlCommand("StallFabricationPreevent", MyConn);

                com.CommandType = CommandType.StoredProcedure;
                com.Parameters.AddWithValue("@ApproverPreEventURL", targetRow1?.Cells[1].Value ?? "no url");
                com.Parameters.AddWithValue("@FinanceTreasuryURL", targetRow2?.Cells[1].Value ?? "no url");
                com.Parameters.AddWithValue("@InitiatorURL", targetRow4?.Cells[1].Value ?? "no url");
                com.Parameters.AddWithValue("@EventTopic", formDataList.StallFabrication.EventName);
                com.Parameters.AddWithValue("@EventType", formDataList.StallFabrication.EventType);
                com.Parameters.AddWithValue("@EventDate", formDataList.StallFabrication.StartDate);

                com.Parameters.AddWithValue("@EndDate", formDataList.StallFabrication.EndDate);

                com.Parameters.AddWithValue("@Brands", brand);
                com.Parameters.AddWithValue("@Expenses", Expense);

                com.Parameters.AddWithValue("@IsAdvanceRequired", formDataList.StallFabrication.IsAdvanceRequired);

                com.Parameters.AddWithValue("@InitiatorName", formDataList.StallFabrication.InitiatorName);
                com.Parameters.AddWithValue("@AdvanceAmount", SheetHelper.NumCheck(formDataList.StallFabrication.AdvanceAmount));
                com.Parameters.AddWithValue("@TotalExpenseBTC", SheetHelper.NumCheck(formDataList.StallFabrication.TotalExpenseBTC));
                com.Parameters.AddWithValue("@TotalExpenseBTE", SheetHelper.NumCheck(formDataList.StallFabrication.TotalExpenseBTE));

                com.Parameters.AddWithValue("@BudgetAmount", total);

                com.Parameters.AddWithValue("@TotalExpense", Math.Round(TotalExpenseAmount, 2));
                com.Parameters.AddWithValue("@InitiatorEmail", formDataList.StallFabrication.Initiator_Email);
                com.Parameters.AddWithValue("@RBMBM", formDataList.StallFabrication.RBMorBM);
                com.Parameters.AddWithValue("@SalesHead", formDataList.StallFabrication.Sales_Head);
                com.Parameters.AddWithValue("@SalesCoordinator", formDataList.StallFabrication.SalesCoordinatorEmail);
                com.Parameters.AddWithValue("@MarketingCoordinator", formDataList.StallFabrication.MarketingCoordinatorEmail);
                com.Parameters.AddWithValue("@MarketingHead", formDataList.StallFabrication.Marketing_Head);
                com.Parameters.AddWithValue("@Compliance", formDataList.StallFabrication.ComplianceEmail);
                com.Parameters.AddWithValue("@FinanceAccounts", formDataList.StallFabrication.FinanceAccountsEmail);
                com.Parameters.AddWithValue("@FinanceTreasury", formDataList.StallFabrication.Finance);
                com.Parameters.AddWithValue("@ReportingManager", formDataList.StallFabrication.ReportingManagerEmail);
                com.Parameters.AddWithValue("@1UpManager", formDataList.StallFabrication.FirstLevelEmail);
                com.Parameters.AddWithValue("@MedicalAffairsHead", formDataList.StallFabrication.MedicalAffairsEmail);
                com.Parameters.AddWithValue("@BTEExpenseDetails", addedExpencesBTE);
                com.Parameters.AddWithValue("@AttachmentPaths", Attachmentpaths);


                com.Parameters.AddWithValue("@Class_III_EventCode", formDataList.StallFabrication.Class_III_EventCode);
                com.Parameters.AddWithValue("@webinarRole", formDataList.StallFabrication.Role);



                MyConn.Open();
                //com.ExecuteNonQuery();
                MySqlDataReader reader = com.ExecuteReader();
                String RefID = "";
                while (reader.Read())
                {
                    RefID = reader["ID"].ToString();
                }
                MyConn.CloseAsync();

                MyConn.Open();
                com = new MySqlCommand("SPEventRequestsBrandsList", MyConn);
                com.CommandType = CommandType.StoredProcedure;
                foreach (var formdata in formDataList.EventBrands)
                {
                    com.Parameters.AddWithValue("@Allocation", formdata.PercentAllocation);
                    com.Parameters.AddWithValue("@Brands", formdata.BrandName);
                    com.Parameters.AddWithValue("@ProjectID", formdata.ProjectId);
                    com.Parameters.AddWithValue("@EventIdEventRequestId", RefID);
                    com.ExecuteNonQuery();
                    com.Parameters.Clear();
                }
                MyConn.CloseAsync();






                MyConn.Open();
                com = new MySqlCommand("SPEventRequestExpensesSheet", MyConn);
                com.CommandType = CommandType.StoredProcedure;

                foreach (var formdata in formDataList.ExpenseSheets)
                {
                    com.Parameters.AddWithValue("@Expense", formdata.Expense);
                    com.Parameters.AddWithValue("@EventIdEventRequestID", RefID);
                    com.Parameters.AddWithValue("@AmountExcludingTaxq", formdata.AmountExcludingTax);
                    com.Parameters.AddWithValue("@AmountExcludingTax", formdata.ExcludingTaxAmount);
                    com.Parameters.AddWithValue("@Amount", SheetHelper.NumCheck(formdata.Amount));
                    com.Parameters.AddWithValue("@BTCBTE", formdata.BtcorBte);
                    com.Parameters.AddWithValue("@BudgetAmount", SheetHelper.NumCheck(formdata.BudgetAmount));
                    com.Parameters.AddWithValue("@BTCAmount", SheetHelper.NumCheck(formdata.BtcAmount));
                    com.Parameters.AddWithValue("@BTEAmount", SheetHelper.NumCheck(formdata.BteAmount));
                    com.Parameters.AddWithValue("@EventTopic", formDataList.StallFabrication.EventName);
                    com.Parameters.AddWithValue("@EventType", formDataList.StallFabrication.EventType);
                    com.Parameters.AddWithValue("@EventDateStart", formDataList.StallFabrication.StartDate);
                    com.Parameters.AddWithValue("@EventEndDate", formDataList.StallFabrication.EndDate);
                    com.ExecuteNonQuery();
                    com.Parameters.Clear();
                }
                MyConn.CloseAsync();






                #region

                if (formDataList.StallFabrication.IsDeviationUpload == "Yes")
                {
                    List<string> DeviationFiles = new List<string>();

                    if (uploadDeviationForTableContainsData == "Yes")
                    {
                        string[] deviationFile = formDataList.StallFabrication.TableContainsDataUpload.Split(":");
                        string name = deviationFile[0].Split(".")[0];
                        string ext = deviationFile[0].Split(".")[1];
                        string base64 = deviationFile[1];
                        string concatFile = name + "*" + "30DaysDeviationFile" + "." + ext + ":" + base64;
                        DeviationFiles.Add(concatFile);
                    }
                    if (EventWithin7Days == "Yes")
                    {
                        string[] deviationFile = formDataList.StallFabrication.EventWithin7daysUpload.Split(":");
                        string name = deviationFile[0].Split(".")[0];
                        string ext = deviationFile[0].Split(".")[1];
                        string base64 = deviationFile[1];
                        string concatFile = name + "*" + "7DaysDeviationFile" + "." + ext + ":" + base64;
                        DeviationFiles.Add(concatFile);
                    }



                    List<string> DeviationNames = new List<string>();
                    foreach (var p in DeviationFiles)
                    {
                        string[] words = p.Split(':')[0].Split("*");
                        // string[] words = p.DeviationFile.Split(':')[0].Split("*");
                        string r = words[1];
                        DeviationNames.Add(r);
                    }
                    MyConn.Open();
                    com = new MySqlCommand("SPDeviation_Process", MyConn);
                    com.CommandType = CommandType.StoredProcedure;

                    foreach (var pp in DeviationFiles)
                    {
                        foreach (var deviationname in DeviationNames)
                        {
                            string file = deviationname.Split(".")[0];
                            string DeviationAttachmentpath = "";
                            if (pp.Split(':')[0].Split("*")[1] == deviationname)
                            {
                                try
                                {
                                    foreach (var p in DeviationFiles)
                                    {
                                        string[] nameSplit = p.Split("*");
                                        string[] words = nameSplit[1].Split(':');
                                        string r = words[0];
                                        string q = words[1];
                                        if (deviationname == r)
                                        {
                                            string name = nameSplit[0];
                                            string filePath = SheetHelper.SQlFileinsertion(q, name);
                                            DeviationAttachmentpath = DeviationAttachmentpath + "," + filePath;
                                            //Attachment attachmentinmain = await ApiCalls.AddAttachmentsToSheet(smartsheet, sheet1, addedRows[0], filePath);
                                        }
                                    }

                                    com.Parameters.AddWithValue("@EventIdEventRequestId", RefID);
                                    com.Parameters.AddWithValue("@EventTopic", formDataList.StallFabrication.EventName);
                                    com.Parameters.AddWithValue("@EventType", formDataList.StallFabrication.EventType);
                                    com.Parameters.AddWithValue("@EventDate", formDataList.StallFabrication.StartDate);
                                    com.Parameters.AddWithValue("@EndDate", formDataList.StallFabrication.EndDate);
                                    com.Parameters.AddWithValue("@StartTime", "");
                                    com.Parameters.AddWithValue("@EndTime", "");
                                    com.Parameters.AddWithValue("@MISCode", "");
                                    com.Parameters.AddWithValue("@HCPName", "");
                                    com.Parameters.AddWithValue("@HonorariumAmount", "");
                                    com.Parameters.AddWithValue("@TravelAccommodationAmount", "");
                                    com.Parameters.AddWithValue("@OtherExpenses", "");


                                    if (file == "30DaysDeviationFile")
                                    {
                                        com.Parameters.AddWithValue("@DeviationType", configuration.GetSection("DeviationNamesInPreEvent:30DaysDeviationFile").Value);
                                        com.Parameters.AddWithValue("@EventOpen45days", "Yes");
                                        com.Parameters.AddWithValue("@OutstandingEvents", SheetHelper.NumCheck(formDataList.StallFabrication.EventOpen30dayscount));
                                    }
                                    else
                                    {
                                        com.Parameters.AddWithValue("@EventOpen45days", "");
                                        com.Parameters.AddWithValue("@OutstandingEvents", "");
                                    }
                                    if (file == "7DaysDeviationFile")
                                    {
                                        com.Parameters.AddWithValue("@DeviationType", configuration.GetSection("DeviationNamesInPreEvent:5DaysDeviationFile").Value);
                                        com.Parameters.AddWithValue("@EventWithin5days", "");
                                    }
                                    else
                                    {
                                        com.Parameters.AddWithValue("@EventWithin5days", "");
                                    }
                                    if (file == "ExpenseExcludingTax")
                                    {
                                        com.Parameters.AddWithValue("@DeviationType", configuration.GetSection("DeviationNamesInPreEvent:ExpenseExcludingTax").Value);
                                        com.Parameters.AddWithValue("@PREExpenseExcludingTax", "Yes");
                                    }
                                    else
                                    {
                                        com.Parameters.AddWithValue("@PREExpenseExcludingTax", "");
                                    }
                                    if (file.Contains("Travel_Accomodation3LExceededFile"))
                                    {
                                        com.Parameters.AddWithValue("@DeviationType", configuration.GetSection("DeviationNamesInPreEvent:Travel_Accomodation3LExceededFile").Value);
                                        com.Parameters.AddWithValue("@TravelAccomodationExceededTrigger", "Yes");
                                    }
                                    else
                                    {
                                        com.Parameters.AddWithValue("@TravelAccomodationExceededTrigger", "");
                                    }
                                    if (file.Contains("TrainerHonorarium12LExceededFile"))
                                    {
                                        com.Parameters.AddWithValue("@DeviationType", configuration.GetSection("DeviationNamesInPreEvent:TrainerHonorarium12LExceededFile").Value);
                                        com.Parameters.AddWithValue("@TrainerHonorariumExceededTrigger", "Yes");
                                    }
                                    else
                                    {
                                        com.Parameters.AddWithValue("@TrainerHonorariumExceededTrigger", "");
                                    }
                                    if (file.Contains("HCPHonorarium6LExceededFile"))
                                    {
                                        com.Parameters.AddWithValue("@DeviationType", configuration.GetSection("DeviationNamesInPreEvent:HCPHonorarium6LExceededFile").Value);
                                        com.Parameters.AddWithValue("@HCPHonorariumExceededTrigger", "Yes");
                                    }
                                    else
                                    {
                                        com.Parameters.AddWithValue("@HCPHonorariumExceededTrigger", "");
                                    }
                                    com.Parameters.AddWithValue("@SalesHead", formDataList.StallFabrication.Sales_Head);
                                    com.Parameters.AddWithValue("@FinanceHead", formDataList.StallFabrication.FinanceHead);
                                    com.Parameters.AddWithValue("@InitiatorName", formDataList.StallFabrication.InitiatorName);
                                    com.Parameters.AddWithValue("@InitiatorEmail", formDataList.StallFabrication.Initiator_Email);
                                    com.Parameters.AddWithValue("@SalesCoordinator", formDataList.StallFabrication.SalesCoordinatorEmail);
                                    com.Parameters.AddWithValue("@AttachmentPaths", DeviationAttachmentpath);
                                    com.ExecuteNonQuery();
                                    com.Parameters.Clear();
                                }
                                catch (Exception ex)
                                {
                                    return BadRequest(new
                                    {
                                        Message = ex.Message + "------" + ex.StackTrace
                                    });

                                }

                            }

                        }

                    }
                }

                #endregion



                //Row addedrow = addedRows[0];
                //long ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role");
                //Cell UpdateB = new Cell { ColumnId = ColumnId, Value = formDataList.Webinar.Role };
                //Row updateRows = new Row { Id = addedrow.Id, Cells = new Cell[] { UpdateB } };
                //Cell? cellsToUpdate = addedrow.Cells.FirstOrDefault(c => c.ColumnId == ColumnId);
                //if (cellsToUpdate != null) { cellsToUpdate.Value = formDataList.Webinar.Role; }

                //strMessage += "==Before adding Role to WebSheet " + "==" + DateTime.Now.ToString() + "==";
                //await Task.Run(() => ApiCalls.UpdateRole(smartsheet, sheet1, updateRows)); //smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRows }));
                //strMessage += "==Before adding Role to WebSheet " + "==" + DateTime.Now.ToString() + "==";
                //Log.Information("End of api " + DateTime.Now);
                MyConn.CloseAsync();
                //return Ok(new
                //{ Message = " Success!" });/* { Message = " Success!" });*/
                DateTime currentDate = DateTime.Now;
                return Ok(new
                {
                    Message = $"Thank you. Your event creation request has been received. " +
                "You should receive a confirmation email with the details of your event after a few minutes."
                });

            }
            catch (Exception ex)
            {
                //return BadRequest($"Could not find {ex.Message}");
                Log.Error($"Error occured on class1 method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                //return BadRequest(ex.Message);
                return BadRequest(new
                { Message = ex.Message + "------" + ex.StackTrace });
            }
            //finally
            //{
            //    semaphore.Release();
            //}
        }


        [HttpPost("MedicalUtilityPreEvent"), DisableRequestSizeLimit]
        public IActionResult MedicalUtilityPreEvent(MedicalUtilityPreEventPayload formDataList)
        {

            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
            Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
            Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
            Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);


            StringBuilder addedBrandsData = new();
            StringBuilder addedHcpData = new();
            StringBuilder addedExpences = new();

            int addedHcpDataNo = 1;
            int addedBrandsDataNo = 1;
            int addedExpencesNo = 1;


            double TotalExpenseAmount = 0;
            string EventOpen30Days = !string.IsNullOrEmpty(formDataList.MedicalUtilityData.EventOpen30daysFile) ? "Yes" : "No";
            string EventWithin7Days = !string.IsNullOrEmpty(formDataList.MedicalUtilityData.EventWithin7daysFile) ? "Yes" : "No";
            string UploadDeviationFile = !string.IsNullOrEmpty(formDataList.MedicalUtilityData.UploadDeviationFile) ? "Yes" : "No";
            string FCPA = "";




            foreach (var formdata in formDataList.ExpenseSheet)
            {
                string rowData = $"{addedExpencesNo}. {formdata.Expense} | TotalAmount: {formdata.TotalExpenseAmount}| {formdata.BTC_BTE}";
                addedExpences.AppendLine(rowData);
                addedExpencesNo++;
                double amount = SheetHelper.NumCheck(formdata.TotalExpenseAmount);
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
                string rowData = $"{addedHcpDataNo}. {formdata.MisCode} |{formdata.HcpName} |Speciality: {formdata.Speciality} |Tier: {formdata.Tier} |Rationale :{formdata.Rationale}";
                addedHcpData.AppendLine(rowData);
                addedHcpDataNo++;
            }
            string HCP = addedHcpData.ToString();

            double total = TotalExpenseAmount;


            StringBuilder addedExpencesBTE = new();
            int addedExpencesNoBTE = 1;
            foreach (var formdata in formDataList.ExpenseSheet)
            {
                if (formdata.BTC_BTE.ToLower() == "bte")
                {
                    string rowData = $"{addedExpencesNoBTE}. {formdata.Expense} | Amount: {formdata.TotalExpenseAmount}";
                    addedExpencesBTE.AppendLine(rowData);
                    addedExpencesNoBTE++;
                }
            }
            string BTEExpense = addedExpencesBTE.ToString();


            try
            {

                Row newRow = new()
                {
                    Cells = new List<Cell>()
                };
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.MedicalUtilityData.EventTopic });

                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Type"), Value = formDataList.MedicalUtilityData.EventType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Date"), Value = formDataList.MedicalUtilityData.EventDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Valid From"), Value = formDataList.MedicalUtilityData.ValidFrom });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Valid To"), Value = formDataList.MedicalUtilityData.ValidTill });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Utility Type"), Value = formDataList.MedicalUtilityData.MedicalUtilityType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Utility Description"), Value = formDataList.MedicalUtilityData.MedicalUtilityDescription });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "IsAdvanceRequired"), Value = formDataList.MedicalUtilityData.IsAdvanceRequired });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Advance Amount"), Value = SheetHelper.NumCheck(formDataList.MedicalUtilityData.AdvanceAmount) });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Brands"), Value = brand });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Expenses"), Value = Expense });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Panelists"), Value = HCP });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Name"), Value = formDataList.MedicalUtilityData.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense"), Value = TotalExpenseAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Budget Amount"), Value = total });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Email"), Value = formDataList.MedicalUtilityData.Initiator_Email });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "RBM/BM"), Value = formDataList.MedicalUtilityData.RBMorBM });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Head"), Value = formDataList.MedicalUtilityData.Sales_Head });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Coordinator"), Value = formDataList.MedicalUtilityData.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Marketing Coordinator"), Value = formDataList.MedicalUtilityData.MarketingCoordinatorEmail });
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
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "BTE Expense Details"), Value = BTEExpense });


                IList<Row> addedRows = smartsheet.SheetResources.RowResources.AddRows(sheet1.Id.Value, new Row[] { newRow });

                long eventIdColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                Cell? eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                string val = eventIdCell.DisplayValue;


                List<string> DeviationNames = new()
                {
                    $"30DaysDeviationFile:{formDataList.MedicalUtilityData.EventOpen30daysFile}",
                    $"7DaysDeviationFile:{formDataList.MedicalUtilityData.EventWithin7daysFile}",
                    $"AgregateSpendDeviationFile:{formDataList.MedicalUtilityData.UploadDeviationFile}"
                };

                if (EventOpen30Days == "Yes" || EventWithin7Days == "Yes" || UploadDeviationFile == "Yes")
                {
                    string eventId = val;
                    foreach (var name in DeviationNames)
                    {

                        string[] y = name.Split(':');
                        string fn = y[0];
                        string bs = y[1];

                        if (bs != "")
                        {
                            try
                            {
                                Row newRow7 = new()
                                {
                                    Cells = new List<Cell>()
                                };
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.MedicalUtilityData.EventTopic });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Type"), Value = formDataList.MedicalUtilityData.EventType });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Date"), Value = formDataList.MedicalUtilityData.EventDate });
                                if (fn == "30DaysDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = EventOpen30Days });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = SheetHelper.NumCheck(formDataList.MedicalUtilityData.EventOpen30dayscount) });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:30DaysDeviationFile").Value });

                                }
                                else if (fn == "7DaysDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = EventWithin7Days });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:5DaysDeviationFile").Value });

                                }
                                else if (fn == "AgregateSpendDeviationFile")
                                {
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HCP exceeds 1,00,000 Trigger"), Value = UploadDeviationFile });
                                    newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:AgregateSpendDeviationFile1L").Value });
                                }
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.MedicalUtilityData.Sales_Head });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = /*"sybil.oommen@menariniapac.com"*/formDataList.MedicalUtilityData.FinanceHead });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Name"), Value = formDataList.MedicalUtilityData.InitiatorName });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.MedicalUtilityData.Initiator_Email });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Coordinator"), Value = formDataList.MedicalUtilityData.SalesCoordinatorEmail });


                                IList<Row> addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });

                                string filename = bs.Split("*")[0].Split(".")[0];
                                string filePath = SheetHelper.testingFile(bs.Split("*")[1], filename);



                                Row addedRow = addeddeviationrow[0];

                                Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                        sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                                Attachment attachmentintoMain = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                        sheet1.Id.Value, addedRows[0].Id.Value, filePath, "application/msword");

                                if (System.IO.File.Exists(filePath))
                                {
                                    System.IO.File.Delete(filePath);
                                }

                            }
                            catch (Exception ex)
                            {
                                // return BadRequest(ex.Message);
                                Log.Error($"Error occured on AllPreEventsController  method {ex.Message} at {DateTime.Now}");
                                Log.Error(ex.StackTrace);
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
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "MISCode"), Value = SheetHelper.MisCodeCheck(formData.MisCode) });
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
                    IList<Row> addeddatarows = smartsheet.SheetResources.RowResources.AddRows(sheet4.Id.Value, new Row[] { newRow1 });

                    string FCPAFile = !string.IsNullOrEmpty(formData.UploadFCPA) ? "Yes" : "No";
                    string UploadWrittenRequestDate = !string.IsNullOrEmpty(formData.UploadWrittenRequestDate) ? "Yes" : "No";
                    string Invoice_Brouchere_Quotation = !string.IsNullOrEmpty(formData.Invoice_Brouchere_Quotation) ? "Yes" : "No";

                    long columnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                    Cell? Cell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == columnId);
                    string value = Cell.DisplayValue;
                    if (FCPAFile == "Yes")
                    {
                        string filename = " FCPA";
                        string filePath = SheetHelper.testingFile(formData.UploadFCPA, filename);
                        Row addedRow = addeddatarows[0];
                        Row webRow = addedRows[0];
                        Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet4.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                        Attachment webattachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet1.Id.Value, webRow.Id.Value, filePath, "application/msword");
                        if (System.IO.File.Exists(filePath))
                        {
                            SheetHelper.DeleteFile(filePath);
                        }
                    }
                    if (UploadWrittenRequestDate == "Yes")
                    {
                        string filename = " UploadWrittenRequestDate";
                        string filePath = SheetHelper.testingFile(formData.UploadWrittenRequestDate, filename);
                        Row addedRow = addeddatarows[0];
                        Row webRow = addedRows[0];
                        Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet4.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                        Attachment webattachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet1.Id.Value, webRow.Id.Value, filePath, "application/msword");
                        if (System.IO.File.Exists(filePath))
                        {
                            SheetHelper.DeleteFile(filePath);
                        }
                    }

                    if (Invoice_Brouchere_Quotation == "Yes")
                    {
                        string filename = " Invoice_Brouchere_Quotation";
                        string filePath = SheetHelper.testingFile(formData.Invoice_Brouchere_Quotation, filename);
                        Row addedRow = addeddatarows[0];
                        Row webRow = addedRows[0];
                        Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet4.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                        Attachment webattachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet1.Id.Value, webRow.Id.Value, filePath, "application/msword");
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

                            new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "MisCode"), Value = SheetHelper.MisCodeCheck(formdata.MisCode) },
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


                Row s = addedRows[0];
                long ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role");
                Cell UpdateB = new Cell { ColumnId = ColumnId, Value = formDataList.MedicalUtilityData.Role };
                Row updateRows = new Row { Id = s.Id, Cells = new Cell[] { UpdateB } };
                Cell? cellsToUpdate = s.Cells.FirstOrDefault(c => c.ColumnId == ColumnId);
                if (cellsToUpdate != null) { cellsToUpdate.Value = formDataList.MedicalUtilityData.Role; }

                smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRows });


                return Ok(new
                { Message = " Success!" });
            }

            catch (Exception ex)
            {
                //return BadRequest($"Could not find {ex.Message}");
                Log.Error($"Error occured on AllPreEventsController method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(ex.Message);
            }
        }

        [HttpPost("HandsOnPreEvent"), DisableRequestSizeLimit]
        public IActionResult HandsOnPreEvent(HandsOnTrainingPreEvet formDataList)
        {
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
            Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
            Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
            Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
            Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
            Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
            Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);
            Sheet sheet8 = SheetHelper.GetSheetById(smartsheet, sheetId8);
            Sheet sheet9 = SheetHelper.GetSheetById(smartsheet, sheetId9);

            StringBuilder addedBrandsData = new();
            StringBuilder addedInviteesData = new();
            StringBuilder addedHcpData = new();
            StringBuilder addedSlideKitData = new();
            StringBuilder addedExpences = new();
            StringBuilder SelectedProduct = new();
            StringBuilder addedMEnariniInviteesData = new();
            StringBuilder addedBEneficirydata = new();
            int addedSlideKitDataNo = 1;
            int addedHcpDataNo = 1;
            int addedInviteesDataNo = 1;
            int addedBrandsDataNo = 1;
            int addedExpencesNo = 1;
            int SelectedProductNo = 1;
            int addedInviteesDataNoforMenarini = 1;
            int addedBEneficirydataNo = 1;

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
                string rowData = $"{addedSlideKitDataNo}. {formdata.MISCode} | {formdata.SlideKitSelection} | Id :{formdata.SlideKitDocument}";
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
            if (formDataList.HandsOnTraining.AnaestheticBenificiaryDetailsData != null)
            {
                string beneficiaryDetails = $"Currency : {formDataList.HandsOnTraining.AnaestheticBenificiaryDetailsData.Currency} | BenificiaryName :  {formDataList.HandsOnTraining.AnaestheticBenificiaryDetailsData.BenificiaryName} | BankAccountNumber : {formDataList.HandsOnTraining.AnaestheticBenificiaryDetailsData.BenificiaryName}";
                addedBEneficirydata.AppendLine(beneficiaryDetails);
                addedBEneficirydataNo++;
            }
            if (formDataList.HandsOnTraining.VenueBenificiaryDetailsData != null)
            {
                string beneficiaryDetails = $"Currency : {formDataList.HandsOnTraining.VenueBenificiaryDetailsData.Currency} | BenificiaryName :  {formDataList.HandsOnTraining.VenueBenificiaryDetailsData.BenificiaryName} | BankAccountNumber : {formDataList.HandsOnTraining.VenueBenificiaryDetailsData.BenificiaryName}";
                addedBEneficirydata.AppendLine(beneficiaryDetails);
                addedBEneficirydataNo++;
            }
            string BeneficiaryData = addedBEneficirydata.ToString();




            StringBuilder addedExpencesBTE = new();
            int addedExpencesNoBTE = 1;
            foreach (var formdata in formDataList.ExpenseData)
            {
                if (formdata.IsBtcorBte.ToLower() == "bte")
                {
                    string rowData = $"{addedExpencesNoBTE}. {formdata.ExpenseType} | Amount: {formdata.ExpenseAmountIncludingTax}";
                    addedExpencesBTE.AppendLine(rowData);
                    addedExpencesNoBTE++;
                }
            }
            string BTEExpense = addedExpencesBTE.ToString();
            try
            {

                Row newRow = new()
                {
                    Cells = new List<Cell>()
                };

                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Date"), Value = formDataList.HandsOnTraining.EventDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Type"), Value = formDataList.HandsOnTraining.EventType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.HandsOnTraining.EventName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Start Time"), Value = formDataList.HandsOnTraining.EventStartTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "End Time"), Value = formDataList.HandsOnTraining.EventEndTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Product Brand"), Value = formDataList.HandsOnTraining.ProductBrandSelection });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Mode of Training"), Value = formDataList.HandsOnTraining.ModeOfTraining });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Venue Name"), Value = formDataList.HandsOnTraining.VenueName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "City"), Value = formDataList.HandsOnTraining.City });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "State"), Value = formDataList.HandsOnTraining.State });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "HOT Webinar Vendor Name"), Value = formDataList.HandsOnTraining.VendorName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "HOT Webinar Type"), Value = formDataList.HandsOnTraining.WebinarType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Venue Selection Checklist"), Value = formDataList.HandsOnTraining.VenueSelectionCheckList });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Emergency Support"), Value = formDataList.HandsOnTraining.IsVenueHasAnyEmergancySupport });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Emergency Contact No"), Value = formDataList.HandsOnTraining.EmerganctContact });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Facility Charges"), Value = formDataList.HandsOnTraining.IsVenueFacilityCharges });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Facility Charges BTC/BTE"), Value = formDataList.HandsOnTraining.VenueFacilityChargesBtc_Bte });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Facility Charges Excluding Tax"), Value = formDataList.HandsOnTraining.FacilityChargesExcludingTax });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Facility Charges Including Tax"), Value = formDataList.HandsOnTraining.FacilityChargesIncludingTax });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Anesthetist Required?"), Value = formDataList.HandsOnTraining.IsAnesthetistRequired });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Anesthetist BTC/BTE"), Value = formDataList.HandsOnTraining.AnesthetistRequiredBtc_Bte });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Anesthetist Excluding Tax"), Value = formDataList.HandsOnTraining.AnesthetistChargesExcludingTax });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Anesthetist Including Tax"), Value = formDataList.HandsOnTraining.AnesthetistChargesIncludingTax });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Name"), Value = formDataList.HandsOnTraining.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Advance Amount"), Value = formDataList.HandsOnTraining.AdvanceAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, " Total Expense BTC"), Value = formDataList.HandsOnTraining.TotalExpenseBTC });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense BTE"), Value = formDataList.HandsOnTraining.TotalExpenseBTE });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Honorarium Amount"), Value = formDataList.HandsOnTraining.TotalHonorariumAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel Amount"), Value = formDataList.HandsOnTraining.TotalTravelAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel & Accommodation Amount"), Value = formDataList.HandsOnTraining.TotalTravelAccommodationAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Accommodation Amount"), Value = formDataList.HandsOnTraining.TotalAccomodationAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Budget Amount"), Value = formDataList.HandsOnTraining.TotalBudget });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Local Conveyance"), Value = formDataList.HandsOnTraining.TotalLocalConveyance });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense"), Value = formDataList.HandsOnTraining.TotalExpense });

                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Email"), Value = formDataList.HandsOnTraining.InitiatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "RBM/BM"), Value = formDataList.HandsOnTraining.RBMorBMEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Head"), Value = formDataList.HandsOnTraining.SalesHeadEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Coordinator"), Value = formDataList.HandsOnTraining.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Marketing Coordinator"), Value = formDataList.HandsOnTraining.MarketingCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Marketing Head"), Value = formDataList.HandsOnTraining.MarketingHeadEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Treasury"), Value = formDataList.HandsOnTraining.FinanceEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Compliance"), Value = formDataList.HandsOnTraining.ComplianceEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Accounts"), Value = formDataList.HandsOnTraining.FinanceAccountsEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Reporting Manager"), Value = formDataList.HandsOnTraining.ReportingManagerEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "1 Up Manager"), Value = formDataList.HandsOnTraining.FirstLevelEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Affairs Head"), Value = formDataList.HandsOnTraining.MedicalAffairsEmail });

                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Panelists"), Value = HCP });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Invitees"), Value = Invitees });

                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "MIPL Invitees"), Value = MenariniInvitees });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Brands"), Value = brand });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Expenses"), Value = Expense });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "SlideKits"), Value = slideKit });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Beneficiary Details"), Value = BeneficiaryData });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Selected Products"), Value = SelectedProductData });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "BTE Expense Details"), Value = BTEExpense });

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
                    string filePath = SheetHelper.testingFile(q, name);
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
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Type"), Value = formDataList.HandsOnTraining.EventType });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Even tDate"), Value = formDataList.HandsOnTraining.EventDate });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Start Time"), Value = formDataList.HandsOnTraining.EventStartTime });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "End Time"), Value = formDataList.HandsOnTraining.EventEndTime });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Venue Name"), Value = formDataList.HandsOnTraining.VenueName });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "City"), Value = formDataList.HandsOnTraining.City });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "State"), Value = formDataList.HandsOnTraining.State });


                            if (file == "30DaysDeviationFile")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:30DaysDeviationFile").Value /*"Outstanding with initiator for more than 45 days"*/ });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = "Yes" });//formDataList.HandsOnTraining.EventOpen30days });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = formDataList.HandsOnTraining.EventOpen30dayscount });
                            }
                            else if (file == "7DaysDeviationFile")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:5DaysDeviationFile").Value /*"5 days from the Event Date"*/ });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = "Yes" });//formDataList.HandsOnTraining.EventWithin7days });

                            }
                            else if (file == "ExpenseExcludingTax")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:ExpenseExcludingTax").Value /* "Food and Beverages expense exceeds 1500" */});
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "PRE-F&B Expense Excluding Tax"), Value = "Yes" });
                            }
                            else if (file.Contains("Travel_Accomodation3LExceededFile"))
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:Travel_Accomodation3LExceededFile").Value /*"Travel/AccomodationAggregate Limit of 3,00,000 is Exceeded"*/ });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Travel/Accomodation 3,00,000 Exceeded Trigger"), Value = "Yes" });//formDataList.class1.FB_Expense_Excluding_Tax });
                            }
                            else if (file.Contains("TrainerHonorarium12LExceededFile"))
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:TrainerHonorarium12LExceededFile").Value  /*"Honorarium Aggregate Limit of 12,00,000 is Exceeded" */});
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Trainer Honorarium 12,00,000 Exceeded Trigger"), Value = "Yes" }); //formDataList.class1.FB_Expense_Excluding_Tax });
                            }
                            else if (file.Contains("HCPHonorarium6LExceededFile"))
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:HCPHonorarium6LExceededFile").Value /* "Honorarium Aggregate Limit of 6,00,000 is Exceeded"*/ });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HCP Honorarium 6,00,000 Exceeded Trigger"), Value = "Yes" });
                            }
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.HandsOnTraining.SalesHeadEmail });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formDataList.HandsOnTraining.FinanceEmail });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Name"), Value = formDataList.HandsOnTraining.InitiatorName });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.HandsOnTraining.InitiatorEmail });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Coordinator"), Value = formDataList.HandsOnTraining.SalesCoordinatorEmail });

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
                                    string filePath = SheetHelper.testingFile(q, name);
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
                            //return BadRequest(ex.Message);
                            Log.Error($"Error occured on AllPreEventsController method {ex.Message} at {DateTime.Now}");
                            Log.Error(ex.StackTrace);
                            return BadRequest(ex.Message);
                        }
                    }
                }

                foreach (var formData in formDataList.TrainerDetails)
                {
                    Row newRow1 = new()
                    {
                        Cells = new List<Cell>()
                    };

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
                        string filePath = SheetHelper.testingFile(q, name);
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


                foreach (var formdata in formDataList.AttenderSelections)
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
                        string filePath = SheetHelper.testingFile(q, name);
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
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "MIS"), Value = SheetHelper.MisCodeCheck(formdata.MISCode) });
                    newRow5.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet5, "HCP Name"), Value = formdata.TrainerName });
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
                    if (formDataList.HandsOnTraining.VenueBenificiaryDetailsData != null)
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
                        newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Other Currency"), Value = formDataList.HandsOnTraining.VenueBenificiaryDetailsData.EnterCurrencyType });
                        newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Beneficiary Name"), Value = formDataList.HandsOnTraining.VenueBenificiaryDetailsData.BenificiaryName });
                        newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Bank Account Number"), Value = formDataList.HandsOnTraining.VenueBenificiaryDetailsData.BankAccountNumber });
                        newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Facility Charges"), Value = formDataList.HandsOnTraining.IsVenueFacilityCharges });
                        newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Anesthetist Required?"), Value = formDataList.HandsOnTraining.IsAnesthetistRequired });
                        newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Currency"), Value = formDataList.HandsOnTraining.VenueBenificiaryDetailsData.Currency });
                        newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Bank Name"), Value = formDataList.HandsOnTraining.VenueBenificiaryDetailsData.BankName });
                        newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "PAN card name"), Value = formDataList.HandsOnTraining.VenueBenificiaryDetailsData.NameasPerPAN });
                        newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Pan Number"), Value = formDataList.HandsOnTraining.VenueBenificiaryDetailsData.PANCardNumber });
                        newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "IFSC Code"), Value = formDataList.HandsOnTraining.VenueBenificiaryDetailsData.IFSCCode });
                        //newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, ""), Value = formDataList.HandsOnTraining.BenificiaryDetailsData.IbnNumber });
                        //newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formDataList.HandsOnTraining.BenificiaryDetailsData.SwiftCode });
                        //newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formDataList.HandsOnTraining.BenificiaryDetailsData.TaxResidenceCertificateDate });
                        newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Email Id"), Value = formDataList.HandsOnTraining.VenueBenificiaryDetailsData.EmailID });

                        smartsheet.SheetResources.RowResources.AddRows(sheet8.Id.Value, new Row[] { newRow8 });
                    }

                    if (formDataList.HandsOnTraining.AnaestheticBenificiaryDetailsData != null)
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
                        newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Other Currency"), Value = formDataList.HandsOnTraining.AnaestheticBenificiaryDetailsData.EnterCurrencyType });
                        newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Beneficiary Name"), Value = formDataList.HandsOnTraining.AnaestheticBenificiaryDetailsData.BenificiaryName });
                        newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Bank Account Number"), Value = formDataList.HandsOnTraining.AnaestheticBenificiaryDetailsData.BankAccountNumber });
                        newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Facility Charges"), Value = formDataList.HandsOnTraining.IsVenueFacilityCharges });
                        newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Anesthetist Required?"), Value = formDataList.HandsOnTraining.IsAnesthetistRequired });
                        newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Currency"), Value = formDataList.HandsOnTraining.AnaestheticBenificiaryDetailsData.Currency });
                        newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Bank Name"), Value = formDataList.HandsOnTraining.AnaestheticBenificiaryDetailsData.BankName });
                        newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "PAN card name"), Value = formDataList.HandsOnTraining.AnaestheticBenificiaryDetailsData.NameasPerPAN });
                        newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Pan Number"), Value = formDataList.HandsOnTraining.AnaestheticBenificiaryDetailsData.PANCardNumber });
                        newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "IFSC Code"), Value = formDataList.HandsOnTraining.AnaestheticBenificiaryDetailsData.IFSCCode });
                        //newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, ""), Value = formDataList.HandsOnTraining.BenificiaryDetailsData.IbnNumber });
                        //newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formDataList.HandsOnTraining.BenificiaryDetailsData.SwiftCode });
                        //newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet6, "BTC/BTE"), Value = formDataList.HandsOnTraining.BenificiaryDetailsData.TaxResidenceCertificateDate });
                        newRow8.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet8, "Email Id"), Value = formDataList.HandsOnTraining.AnaestheticBenificiaryDetailsData.EmailID });

                        smartsheet.SheetResources.RowResources.AddRows(sheet8.Id.Value, new Row[] { newRow8 });
                    }

                }
                catch (Exception ex)
                {
                    //return BadRequest($"Could not find {ex.Message}");
                    Log.Error($"Error occured on AllPreEventsController Attachementfile method {ex.Message} at {DateTime.Now}");
                    Log.Error(ex.StackTrace);
                    return BadRequest(ex.Message);
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

                    IList<Row> n = smartsheet.SheetResources.RowResources.AddRows(sheet9.Id.Value, new Row[] { newRow9 });
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
                //return BadRequest($"Could not find {ex.Message}");
                Log.Error($"Error occured on AllPreEventsController Attachementfile method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(ex.Message);
            }
        }

        [HttpPost("DemoMeetingsPreEvent"), DisableRequestSizeLimit]
        public IActionResult DemoMeetingsPreEvent(DemoMeetingsPreEvent formDataList)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

                Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
                Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
                Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
                Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
                Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
                Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
                Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);
                Sheet sheet8 = SheetHelper.GetSheetById(smartsheet, sheetId8);
                Sheet sheet9 = SheetHelper.GetSheetById(smartsheet, sheetId9);

                StringBuilder addedBrandsData = new();
                StringBuilder addedInviteesData = new();
                StringBuilder addedHcpData = new();
                StringBuilder addedSlideKitData = new();
                StringBuilder addedExpences = new();
                StringBuilder SelectedProduct = new();
                StringBuilder addedMEnariniInviteesData = new();
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
                    string rowData = $"{addedSlideKitDataNo}. {formdata.MISCode} | {formdata.SlideKitSelection} | Id :{formdata.SlideKitDocument}";
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

                    string rowData = $"{addedHcpDataNo}. {formdata.HcpType} |{formdata.HcpName} | Honr.Amt: {formdata.HonorariumAmountincludingTax} |Trav.&Acc.Amt: {formdata.TravelAmountIncludingTax + formdata.AccomodationAmountIncludingTax} ";
                    addedHcpData.AppendLine(rowData);
                    addedHcpDataNo++;

                }
                string HCP = addedHcpData.ToString();
                string beneficiaryDetails = $"Currency : {formDataList.DemoMeetings.BenificiaryDetailsData.Currency} | BenificiaryName :  {formDataList.DemoMeetings.BenificiaryDetailsData.BenificiaryName} | BankAccountNumber : {formDataList.DemoMeetings.BenificiaryDetailsData.BenificiaryName}";



                StringBuilder addedExpencesBTE = new();
                int addedExpencesNoBTE = 1;
                foreach (var formdata in formDataList.ExpenseData)
                {
                    if (formdata.IsBtcorBte.ToLower() == "bte")
                    {
                        string rowData = $"{addedExpencesNoBTE}. {formdata.ExpenseType} | Amount: {formdata.ExpenseAmountIncludingTax}";
                        addedExpencesBTE.AppendLine(rowData);
                        addedExpencesNoBTE++;
                    }
                }
                string BTEExpense = addedExpencesBTE.ToString();
                Row newRow = new()
                {
                    Cells = new List<Cell>()
                };

                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Date"), Value = formDataList.DemoMeetings.EventDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Type"), Value = formDataList.DemoMeetings.EventType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Event Topic"), Value = formDataList.DemoMeetings.EventName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Start Time"), Value = formDataList.DemoMeetings.EventStartTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "End Time"), Value = formDataList.DemoMeetings.EventEndTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Product Brand"), Value = formDataList.DemoMeetings.ProductBrandSelection });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Mode of Training"), Value = formDataList.DemoMeetings.ModeOfTraining });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Venue Name"), Value = formDataList.DemoMeetings.VenueName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "City"), Value = formDataList.DemoMeetings.City });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "State"), Value = formDataList.DemoMeetings.State });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "HOT Webinar Vendor Name"), Value = formDataList.DemoMeetings.VendorName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "HOT Webinar Type"), Value = formDataList.DemoMeetings.WebinarType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Venue Selection Checklist"), Value = formDataList.DemoMeetings.VenueSelectionCheckList });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Emergency Support"), Value = formDataList.DemoMeetings.IsVenueHasAnyEmergancySupport });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Emergency Contact No"), Value = formDataList.DemoMeetings.EmerganctContact });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Facility Charges"), Value = formDataList.DemoMeetings.IsVenueFacilityCharges });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Facility Charges BTC/BTE"), Value = formDataList.DemoMeetings.VenueFacilityChargesBtc_Bte });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Facility Charges Excluding Tax"), Value = formDataList.DemoMeetings.FacilityChargesExcludingTax });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Facility Charges Including Tax"), Value = formDataList.DemoMeetings.FacilityChargesIncludingTax });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Anesthetist Required?"), Value = formDataList.DemoMeetings.IsAnesthetistRequired });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Anesthetist BTC/BTE"), Value = formDataList.DemoMeetings.AnesthetistRequiredBtc_Bte });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Anesthetist Excluding Tax"), Value = formDataList.DemoMeetings.AnesthetistChargesExcludingTax });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Anesthetist Including Tax"), Value = formDataList.DemoMeetings.AnesthetistChargesIncludingTax });

                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Name"), Value = formDataList.DemoMeetings.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Advance Amount"), Value = formDataList.DemoMeetings.AdvanceAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, " Total Expense BTC"), Value = formDataList.DemoMeetings.TotalExpenseBTC });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense BTE"), Value = formDataList.DemoMeetings.TotalExpenseBTE });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Honorarium Amount"), Value = formDataList.DemoMeetings.TotalHonorariumAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel Amount"), Value = formDataList.DemoMeetings.TotalTravelAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Travel & Accommodation Amount"), Value = formDataList.DemoMeetings.TotalTravelAccommodationAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Accommodation Amount"), Value = formDataList.DemoMeetings.TotalAccomodationAmount });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Budget Amount"), Value = formDataList.DemoMeetings.TotalBudget });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Local Conveyance"), Value = formDataList.DemoMeetings.TotalLocalConveyance });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Total Expense"), Value = formDataList.DemoMeetings.TotalExpense });

                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Initiator Email"), Value = formDataList.DemoMeetings.InitiatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "RBM/BM"), Value = formDataList.DemoMeetings.RBMorBMEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Head"), Value = formDataList.DemoMeetings.SalesHeadEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Sales Coordinator"), Value = formDataList.DemoMeetings.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Marketing Coordinator"), Value = formDataList.DemoMeetings.MarketingCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Marketing Head"), Value = formDataList.DemoMeetings.MarketingHeadEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Treasury"), Value = formDataList.DemoMeetings.FinanceEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Compliance"), Value = formDataList.DemoMeetings.ComplianceEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Accounts"), Value = formDataList.DemoMeetings.FinanceAccountsEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Reporting Manager"), Value = formDataList.DemoMeetings.ReportingManagerEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "1 Up Manager"), Value = formDataList.DemoMeetings.FirstLevelEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Medical Affairs Head"), Value = formDataList.DemoMeetings.MedicalAffairsEmail });

                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Panelists"), Value = HCP });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Invitees"), Value = Invitees });

                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "MIPL Invitees"), Value = MenariniInvitees });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Brands"), Value = brand });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Expenses"), Value = Expense });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "SlideKits"), Value = slideKit });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Beneficiary Details"), Value = beneficiaryDetails });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Selected Products"), Value = SelectedProductData });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "BTE Expense Details"), Value = BTEExpense });

                IList<Row> addedRows = smartsheet.SheetResources.RowResources.AddRows(sheet1.Id.Value, new Row[] { newRow });

                long eventIdColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                Cell? eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                string val = eventIdCell.DisplayValue;
                int x = 1;
                foreach (string p in formDataList.DemoMeetings.Files)
                {
                    string[] words = p.Split(':');
                    string r = words[0];
                    string q = words[1];
                    string name = r.Split(".")[0];
                    string filePath = SheetHelper.testingFile(q, name);
                    Row addedRow = addedRows[0];
                    Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                           sheet1.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                    x++;


                    if (System.IO.File.Exists(filePath))
                    {
                        SheetHelper.DeleteFile(filePath);
                    }
                }
                if (formDataList.DemoMeetings.IsDeviationUpload == "Yes")
                {
                    List<string> DeviationNames = new();
                    foreach (string p in formDataList.DemoMeetings.DeviationFiles)
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
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formDataList.DemoMeetings.EventName });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Type"), Value = formDataList.DemoMeetings.EventType });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Date"), Value = formDataList.DemoMeetings.EventDate });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Start Time"), Value = formDataList.DemoMeetings.EventStartTime });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "End Time"), Value = formDataList.DemoMeetings.EventEndTime });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Venue Name"), Value = formDataList.DemoMeetings.VenueName });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "City"), Value = formDataList.DemoMeetings.City });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "State"), Value = formDataList.DemoMeetings.State });


                            if (file == "30DaysDeviationFile")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:30DaysDeviationFile").Value /*"Outstanding with initiator for more than 45 days"*/ });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventOpen45days"), Value = "Yes" });//formDataList.HandsOnTraining.EventOpen30days });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Outstanding Events"), Value = formDataList.DemoMeetings.EventOpen30dayscount });
                            }
                            else if (file == "7DaysDeviationFile")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:5DaysDeviationFile").Value /*"5 days from the Event Date"*/ });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventWithin5days"), Value = "Yes" });//formDataList.HandsOnTraining.EventWithin7days });

                            }
                            else if (file == "ExpenseExcludingTax")
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:ExpenseExcludingTax").Value /* "Food and Beverages expense exceeds 1500" */});
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "PRE-F&B Expense Excluding Tax"), Value = "Yes" });
                            }
                            else if (file.Contains("Travel_Accomodation3LExceededFile"))
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:Travel_Accomodation3LExceededFile").Value /*"Travel/AccomodationAggregate Limit of 3,00,000 is Exceeded"*/ });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Travel/Accomodation 3,00,000 Exceeded Trigger"), Value = "Yes" });//formDataList.class1.FB_Expense_Excluding_Tax });
                            }
                            else if (file.Contains("TrainerHonorarium12LExceededFile"))
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:TrainerHonorarium12LExceededFile").Value  /*"Honorarium Aggregate Limit of 12,00,000 is Exceeded" */});
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Trainer Honorarium 12,00,000 Exceeded Trigger"), Value = "Yes" }); //formDataList.class1.FB_Expense_Excluding_Tax });
                            }
                            else if (file.Contains("HCPHonorarium6LExceededFile"))
                            {
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInPreEvent:HCPHonorarium6LExceededFile").Value /* "Honorarium Aggregate Limit of 6,00,000 is Exceeded"*/ });
                                newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HCP Honorarium 6,00,000 Exceeded Trigger"), Value = "Yes" });
                            }
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formDataList.DemoMeetings.SalesHeadEmail });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formDataList.DemoMeetings.FinanceEmail });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Name"), Value = formDataList.DemoMeetings.InitiatorName });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formDataList.DemoMeetings.InitiatorEmail });
                            newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Coordinator"), Value = formDataList.DemoMeetings.SalesCoordinatorEmail });

                            IList<Row> addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });

                            int j = 1;
                            foreach (var p in formDataList.DemoMeetings.DeviationFiles)
                            {
                                string[] words = p.Split(':');
                                string r = words[0];
                                string q = words[1];
                                if (deviationname == r)
                                {
                                    string name = r.Split(".")[0];
                                    string filePath = SheetHelper.testingFile(q, name);
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
                            //return BadRequest(ex.Message);
                            Log.Error($"Error occured on AllPreEventsController method {ex.Message} at {DateTime.Now}");
                            Log.Error(ex.StackTrace);
                            return BadRequest(ex.Message);
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

                foreach (var formData in formDataList.TrainerDetails)
                {
                    Row newRow1 = new()
                    {
                        Cells = new List<Cell>()
                    };

                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "MISCode"), Value = SheetHelper.MisCodeCheck(formData.MISCode) });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HcpRole"), Value = formData.HCPRole });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCPName"), Value = formData.HcpName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TrainerCode"), Value = formData.HcpCode });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Qualification"), Value = formData.HcpQualification });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Country"), Value = formData.HcpCountry });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Speciality"), Value = formData.Speciality });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Tier"), Value = formData.HcpCategory });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "HCP Type"), Value = formData.HcpType });
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
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Travel BTC/BTE"), Value = formData.IsTravelBTC_BTE });
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
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Topic"), Value = formDataList.DemoMeetings.EventName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Type"), Value = formDataList.DemoMeetings.EventType });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Venue name"), Value = formDataList.DemoMeetings.VenueName });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event Date Start"), Value = formDataList.DemoMeetings.EventDate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Event End Date"), Value = formDataList.DemoMeetings.EventDate });
                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "TotalSpend"), Value = formData.FinalAmount });

                    newRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "EventId/EventRequestId"), Value = val });

                    IList<Row> row = smartsheet.SheetResources.RowResources.AddRows(sheet4.Id.Value, new Row[] { newRow1 });

                    foreach (string p in formData.HcpFiles)
                    {
                        string[] words = p.Split(':');
                        string r = words[0];
                        string q = words[1];
                        string name = r.Split(".")[0];
                        string filePath = SheetHelper.testingFile(q, name);
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


                foreach (var formdata in formDataList.AttenderSelections)
                {
                    Row newRow3 = new()
                    {
                        Cells = new List<Cell>()
                    };
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "MISCode"), Value = SheetHelper.MisCodeCheck(formdata.MisCode) });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "HCPName"), Value = formdata.InviteeName });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "HCP Type"), Value = formdata.InviteeType });
                    //newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Designation"), Value = formdata.Designation });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Employee Code"), Value = formdata.EmployeeCode });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "LocalConveyance"), Value = formdata.IsLocalConveyance });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "BTC/BTE"), Value = formdata.IsBtcorBte });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "LcAmount"), Value = formdata.LocalConveyanceAmountIncludingTax });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Lc Amount Excluding Tax"), Value = formdata.LocalConveyanceAmountExcludingTax });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "EventId/EventRequestId"), Value = val });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Invitee Source"), Value = formdata.InviteedFrom });
                    //newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Qualification"), Value = formdata.Qualification });
                    //newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Experience"), Value = formdata.Experiance });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Topic"), Value = formDataList.DemoMeetings.EventName });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Type"), Value = formDataList.DemoMeetings.EventType });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Venue name"), Value = formDataList.DemoMeetings.VenueName });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event Date Start"), Value = formDataList.DemoMeetings.EventDate });
                    newRow3.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet3, "Event End Date"), Value = formDataList.DemoMeetings.EventDate });

                    IList<Row> row = smartsheet.SheetResources.RowResources.AddRows(sheet3.Id.Value, new Row[] { newRow3 });
                }
            }

            catch (Exception ex)
            {
                Log.Error($"Error occured on AllPreEventsController method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(ex.Message);
            }
            return Ok();
        }

    }
}