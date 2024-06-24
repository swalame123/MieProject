using LazyCache;
using Microsoft.AspNetCore.Mvc;
using MySqlConnector;
using Smartsheet.Api;
using Serilog;
using System.Data;
using Smartsheet.Api.Models;
using IndiaEventsWebApi.Helper;
using Microsoft.Extensions.Logging;
using System.Text;

namespace IndiaEventsWebApi.Controllers.Scheduler
{
    [Route("api/[controller]")]
    [ApiController]
    //[Authorize]
    public class SchedulerSync : ControllerBase
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

        public SchedulerSync(IConfiguration configuration, SemaphoreSlim externalApiSemaphore)
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
        }

        [HttpGet("Getsyncdata")]
        public async Task<IActionResult> GetSyncData()
        {
            try
            {
                string MyConnection = configuration.GetSection("ConnectionStrings:mysql").Value;
                MySqlConnection MyConn = new MySqlConnection(MyConnection);
                MySqlConnection MyConn2 = new MySqlConnection(MyConnection);
                MySqlCommand MyCommand2 = new MySqlCommand("GetAllEventSyncData", MyConn2);
                MyCommand2.CommandType = CommandType.StoredProcedure;
                await MyConn2.OpenAsync();
                MySqlDataAdapter MyAdapter = new MySqlDataAdapter();
                MyAdapter.SelectCommand = MyCommand2;
                DataSet ds = new DataSet();
                MyAdapter.Fill(ds);
                DataTable EventRequestsWebdt = ds.Tables[0];
                DataTable UpdateEventId = new DataTable();
                UpdateEventId.Columns.Add("dbid", typeof(Int64));
                UpdateEventId.Columns.Add("Eventid", typeof(string));
                await MyConn2.CloseAsync();

                if (EventRequestsWebdt.Rows.Count > 0)
                {
                    DataTable EventRequestPanelDetailsdt = new DataTable();
                    DataTable EventRequestsBrandsListdt = new DataTable();
                    DataTable EventRequestInviteesdt = new DataTable();
                    DataTable EventRequestHCPSlideKitDetailsdt = new DataTable();
                    DataTable EventRequestExpensesSheetdt = new DataTable();
                    DataTable Deviation_Processdt = new DataTable();
                    SmartsheetClient smartsheet = await Task.Run(() => SmartSheetBuilder.AccessClient(accessToken, _externalApiSemaphore));
                    Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
                    Sheet sheet2 = SheetHelper.GetSheetById(smartsheet, sheetId2);
                    Sheet sheet3 = SheetHelper.GetSheetById(smartsheet, sheetId3);
                    Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
                    Sheet sheet5 = SheetHelper.GetSheetById(smartsheet, sheetId5);
                    Sheet sheet6 = SheetHelper.GetSheetById(smartsheet, sheetId6);
                    Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);

                    Dictionary<string, long> Sheet1columns = new();
                    foreach (var column in sheet1.Columns)
                    {
                        Sheet1columns.Add(column.Title, (long)column.Id);
                    }

                    Dictionary<string, long> Sheet2columns = new();
                    foreach (var column in sheet2.Columns)
                    {
                        Sheet2columns.Add(column.Title, (long)column.Id);
                    }

                    Dictionary<string, long> Sheet3columns = new();
                    foreach (var column in sheet3.Columns)
                    {
                        Sheet3columns.Add(column.Title, (long)column.Id);
                    }

                    Dictionary<string, long> Sheet4columns = new();
                    foreach (var column in sheet4.Columns)
                    {
                        Sheet4columns.Add(column.Title, (long)column.Id);
                    }

                    Dictionary<string, long> Sheet5columns = new();
                    foreach (var column in sheet5.Columns)
                    {
                        Sheet5columns.Add(column.Title, (long)column.Id);
                    }

                    Dictionary<string, long> Sheet6columns = new();
                    foreach (var column in sheet6.Columns)
                    {
                        Sheet6columns.Add(column.Title, (long)column.Id);
                    }

                    Dictionary<string, long> Sheet7columns = new();
                    foreach (var column in sheet7.Columns)
                    {
                        Sheet7columns.Add(column.Title, (long)column.Id);
                    }

                    for (int loopcount = 0; loopcount < EventRequestsWebdt.Rows.Count; loopcount++)
                    {

                        string expresssion = "`EventId/EventRequestId` =" + EventRequestsWebdt.Rows[loopcount]["ID"];
                        DataRow[] dr1 = ds.Tables[1].Select(expresssion);
                        DataRow[] dr2 = ds.Tables[2].Select(expresssion);
                        DataRow[] dr3 = ds.Tables[3].Select(expresssion);
                        DataRow[] dr4 = ds.Tables[4].Select(expresssion);
                        DataRow[] dr5 = ds.Tables[5].Select(expresssion);
                        DataRow[] dr6 = ds.Tables[6].Select(expresssion);

                        if (dr1.Any())
                        {
                            EventRequestPanelDetailsdt = dr1.CopyToDataTable();
                        }
                        if (dr2.Any())
                        {
                            EventRequestsBrandsListdt = dr2.CopyToDataTable();
                        }
                        if (dr3.Any())
                        {
                            EventRequestInviteesdt = dr3.CopyToDataTable();
                        }
                        if (dr4.Any())
                        {
                            EventRequestHCPSlideKitDetailsdt = dr4.CopyToDataTable();
                        }
                        if (dr5.Any())
                        {
                            EventRequestExpensesSheetdt = dr5.CopyToDataTable();
                        }
                        if (dr6.Any())
                        {
                            Deviation_Processdt = dr6.CopyToDataTable();
                        }

                        if (EventRequestsWebdt.Rows[loopcount]["Event Type"].ToString() == "Stall Fabrication")
                        {
                            try
                            {
                                Cell[] cellsToInsertRequestWeb = new Cell[]
                                   {
                    new Cell { ColumnId =Sheet1columns["Approver Pre Event URL"],Value = EventRequestsWebdt.Rows[loopcount]["Approver Pre Event URL"]},
                    new Cell { ColumnId =Sheet1columns["Finance Treasury URL"],Value = EventRequestsWebdt.Rows[loopcount]["Finance Treasury URL"]},
                    new Cell { ColumnId =Sheet1columns["Initiator URL"],Value = EventRequestsWebdt.Rows[loopcount]["Initiator URL"]},
                    new Cell { ColumnId =Sheet1columns["Event Topic"],Value = EventRequestsWebdt.Rows[loopcount]["Event Topic"]},
                    new Cell { ColumnId =Sheet1columns["Event Type"],Value = EventRequestsWebdt.Rows[loopcount]["Event Type"]},
                    new Cell { ColumnId =Sheet1columns["Event Date"],Value = EventRequestsWebdt.Rows[loopcount]["Event Date"]},
                    new Cell { ColumnId =Sheet1columns["Start Time"],Value = EventRequestsWebdt.Rows[loopcount]["Start Time"]},
                    new Cell { ColumnId =Sheet1columns["End Time"],Value = EventRequestsWebdt.Rows[loopcount]["End Time"]},
                    new Cell { ColumnId =Sheet1columns["Meeting Type"],Value = EventRequestsWebdt.Rows[loopcount]["Meeting Type"]},
                    new Cell { ColumnId =Sheet1columns["Brands"],Value = EventRequestsWebdt.Rows[loopcount]["Brands"]},
                    new Cell { ColumnId =Sheet1columns["Expenses"],Value = EventRequestsWebdt.Rows[loopcount]["Expenses"]},
                    new Cell { ColumnId =Sheet1columns["Panelists"],Value = EventRequestsWebdt.Rows[loopcount]["Panelists"]},
                    new Cell { ColumnId =Sheet1columns["Invitees"],Value = EventRequestsWebdt.Rows[loopcount]["Invitees"]},
                    new Cell { ColumnId =Sheet1columns["MIPL Invitees"],Value = EventRequestsWebdt.Rows[loopcount]["MIPL Invitees"]},
                    new Cell { ColumnId =Sheet1columns["SlideKits"],Value = EventRequestsWebdt.Rows[loopcount]["SlideKits"]},
                    new Cell { ColumnId =Sheet1columns["IsAdvanceRequired"],Value = EventRequestsWebdt.Rows[loopcount]["IsAdvanceRequired"]},
                    new Cell { ColumnId =Sheet1columns["EventOpen30days"],Value = EventRequestsWebdt.Rows[loopcount]["EventOpen30days"]},
                    new Cell { ColumnId =Sheet1columns["EventWithin7days"],Value = EventRequestsWebdt.Rows[loopcount]["EventWithin7days"]},
                    new Cell { ColumnId =Sheet1columns["Initiator Name"],Value = EventRequestsWebdt.Rows[loopcount]["Initiator Name"]},
                    new Cell { ColumnId =Sheet1columns["Advance Amount"],Value = EventRequestsWebdt.Rows[loopcount]["Advance Amount"]},
                    new Cell { ColumnId =Sheet1columns[" Total Expense BTC"],Value = EventRequestsWebdt.Rows[loopcount]["Total Expense BTC"]},
                    new Cell { ColumnId =Sheet1columns["Total Expense BTE"],Value = EventRequestsWebdt.Rows[loopcount]["Total Expense BTE"]},
                    new Cell { ColumnId =Sheet1columns["Total Honorarium Amount"],Value = EventRequestsWebdt.Rows[loopcount]["Total Honorarium Amount"]},
                    new Cell { ColumnId =Sheet1columns["Total Travel Amount"],Value = EventRequestsWebdt.Rows[loopcount]["Total Travel Amount"]},
                    new Cell { ColumnId =Sheet1columns["Total Travel & Accommodation Amount"],Value = EventRequestsWebdt.Rows[loopcount]["Total Travel & Accommodation Amount"]},
                    new Cell { ColumnId =Sheet1columns["Total Accommodation Amount"],Value = EventRequestsWebdt.Rows[loopcount]["Total Accommodation Amount"]},
                    new Cell { ColumnId =Sheet1columns["Budget Amount"],Value = EventRequestsWebdt.Rows[loopcount]["Budget Amount"]},
                    new Cell { ColumnId =Sheet1columns["Total Local Conveyance"],Value = EventRequestsWebdt.Rows[loopcount]["Total Local Conveyance"]},
                    new Cell { ColumnId =Sheet1columns["Total Expense"],Value = EventRequestsWebdt.Rows[loopcount]["Total Expense"]},
                    new Cell { ColumnId =Sheet1columns["Initiator Email"],Value = EventRequestsWebdt.Rows[loopcount]["Initiator Email"]},
                    new Cell { ColumnId =Sheet1columns["RBM/BM"],Value = EventRequestsWebdt.Rows[loopcount]["RBM/BM"]},
                    new Cell { ColumnId =Sheet1columns["Sales Head"],Value = EventRequestsWebdt.Rows[loopcount]["Sales Head"]},
                    new Cell { ColumnId =Sheet1columns["Sales Coordinator"],Value = EventRequestsWebdt.Rows[loopcount]["Sales Coordinator"]},
                    new Cell { ColumnId =Sheet1columns["Marketing Coordinator"],Value = EventRequestsWebdt.Rows[loopcount]["Marketing Coordinator"]},
                    new Cell { ColumnId =Sheet1columns["Marketing Head"],Value = EventRequestsWebdt.Rows[loopcount]["Marketing Head"]},
                    new Cell { ColumnId =Sheet1columns["Compliance"],Value = EventRequestsWebdt.Rows[loopcount]["Compliance"]},
                    new Cell { ColumnId =Sheet1columns["Finance Accounts"],Value = EventRequestsWebdt.Rows[loopcount]["Finance Accounts"]},
                    new Cell { ColumnId =Sheet1columns["Finance Treasury"],Value = EventRequestsWebdt.Rows[loopcount]["Finance Treasury"]},
                    new Cell { ColumnId =Sheet1columns["Reporting Manager"],Value = EventRequestsWebdt.Rows[loopcount]["Reporting Manager"]},
                    new Cell { ColumnId =Sheet1columns["1 Up Manager"],Value = EventRequestsWebdt.Rows[loopcount]["1 Up Manager"]},
                    new Cell { ColumnId =Sheet1columns["Medical Affairs Head"],Value = EventRequestsWebdt.Rows[loopcount]["Medical Affairs Head"]},
                    new Cell { ColumnId =Sheet1columns["BTE Expense Details"],Value = EventRequestsWebdt.Rows[loopcount]["BTE Expense Details"]},
                    new Cell { ColumnId =Sheet1columns["Class III Event Code"],Value = EventRequestsWebdt.Rows[loopcount]["Class III Event Code"]},
                    new Cell { ColumnId=Sheet1columns["End Date"],Value = EventRequestsWebdt.Rows[loopcount]["End Date"]}
                                               };
                                Row row = new Row
                                {
                                    ToBottom = true,
                                    Cells = cellsToInsertRequestWeb,
                                };
                                IList<Row> addedRows = ApiCalls.WebDetails(smartsheet, sheet1, row);

                                long eventIdColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                                Cell? eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == Sheet1columns["EventId/EventRequestId"]);
                                string val = eventIdCell.DisplayValue;

                                string[] eventattachment = EventRequestsWebdt.Rows[loopcount]["AttachmentPaths"].ToString().Split(',');

                                for (int attachcount = 1; attachcount < eventattachment.Length; attachcount++)
                                {
                                    Row addedRow = addedRows[0];
                                    Attachment attachment = await ApiCalls.AddAttachmentsToSheetSync(smartsheet, sheet1, addedRow, eventattachment[attachcount]);
                                    if (System.IO.File.Exists(eventattachment[attachcount]))
                                    {
                                        SheetHelper.DeleteFile(eventattachment[attachcount]);
                                    }
                                }


                                for (int Deviationcount = 0; Deviationcount < Deviation_Processdt.Rows.Count; Deviationcount++)
                                {
                                    Row newRow7 = new()
                                    {
                                        Cells = new List<Cell>()
                                    };
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["EventId/EventRequestId"], Value = val });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Event Topic"], Value = Deviation_Processdt.Rows[Deviationcount]["Event Topic"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Event Type"], Value = Deviation_Processdt.Rows[Deviationcount]["Event Type"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Event Date"], Value = Deviation_Processdt.Rows[Deviationcount]["Event Date"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["End Date"], Value = Deviation_Processdt.Rows[Deviationcount]["End Date"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["EventOpen45days"], Value = Deviation_Processdt.Rows[Deviationcount]["EventOpen45days"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Sales Head"], Value = Deviation_Processdt.Rows[Deviationcount]["Sales Head"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Initiator Name"], Value = Deviation_Processdt.Rows[Deviationcount]["Initiator Name"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Initiator Email"], Value = Deviation_Processdt.Rows[Deviationcount]["Initiator Email"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Deviation Type"], Value = Deviation_Processdt.Rows[Deviationcount]["Deviation Type"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Outstanding Events"], Value = Deviation_Processdt.Rows[Deviationcount]["Outstanding Events"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Sales Coordinator"], Value = Deviation_Processdt.Rows[Deviationcount]["Sales Coordinator"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["EventWithin5days"], Value = Deviation_Processdt.Rows[Deviationcount]["EventWithin5days"] });

                                    IList<Row> addeddeviationrow = ApiCalls.DeviationData(smartsheet, sheet7, newRow7);

                                    string[] Deviationattachment = Deviation_Processdt.Rows[Deviationcount]["AttachmentPaths"].ToString().Split(',');
                                    for (int Deviationattachcount = 1; Deviationattachcount < Deviationattachment.Length; Deviationattachcount++)
                                    {
                                        Row addedRow = addeddeviationrow[0];
                                        Attachment attachment = await ApiCalls.AddAttachmentsToSheetSync(smartsheet, sheet7, addedRow, Deviationattachment[Deviationattachcount]);
                                        Attachment attachmentinmain = await ApiCalls.AddAttachmentsToSheetSync(smartsheet, sheet1, addedRows[0], Deviationattachment[Deviationattachcount]);
                                        if (System.IO.File.Exists(Deviationattachment[Deviationattachcount]))
                                        {
                                            SheetHelper.DeleteFile(Deviationattachment[Deviationattachcount]);
                                        }
                                    }
                                }

                                //Brand sheet insertion

                                List<Row> newRows2 = new();
                                for (int brandcount = 0; brandcount < EventRequestsBrandsListdt.Rows.Count; brandcount++)
                                {
                                    Row newRow2 = new()
                                    {
                                        Cells = new List<Cell>()
                                        {
                                        new(){ ColumnId = Sheet2columns[ "% Allocation"], Value = EventRequestsBrandsListdt.Rows[brandcount]["% Allocation"]},
                                        new(){ ColumnId = Sheet2columns[ "Brands"], Value = EventRequestsBrandsListdt.Rows[brandcount]["Brands"] },
                                        new(){ ColumnId = Sheet2columns[ "Project ID"], Value = EventRequestsBrandsListdt.Rows[brandcount]["Project ID"]},
                                        new(){ ColumnId = Sheet2columns[ "EventId/EventRequestId"], Value = val }
                                        }
                                    };

                                    newRows2.Add(newRow2);
                                }
                                await Task.Run(() => ApiCalls.BrandsDetails(smartsheet, sheet2, newRows2));



                                //ExpenseSheet Insertion
                                List<Row> newRows6 = new();
                                for (int Expensecount = 0; Expensecount < EventRequestExpensesSheetdt.Rows.Count; Expensecount++)
                                {
                                    Row newRow6 = new()
                                    {
                                        Cells = new List<Cell>()
                                        {
                                new () { ColumnId = Sheet6columns[ "Expense"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["Expense"] },
                                new () { ColumnId = Sheet6columns[ "EventId/EventRequestID"], Value = val },
                                new () { ColumnId = Sheet6columns[ "AmountExcludingTax?"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["AmountExcludingTax?"] },
                                new () { ColumnId = Sheet6columns[ "Amount Excluding Tax"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["Amount Excluding Tax"] },
                                new () { ColumnId = Sheet6columns[ "Amount"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["Amount"] },
                                new () { ColumnId = Sheet6columns[ "BTC/BTE"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["BTC/BTE"] },
                                new () { ColumnId = Sheet6columns[ "BudgetAmount"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["BudgetAmount"] },
                                new () { ColumnId = Sheet6columns[ "BTCAmount"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["BTCAmount"] },
                                new () { ColumnId = Sheet6columns[ "BTEAmount"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["BTEAmount"] },
                                new () { ColumnId = Sheet6columns[ "Event Topic"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["Event Topic"] },
                                new () { ColumnId = Sheet6columns[ "Event Type"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["Event Type"] },
                                new () { ColumnId = Sheet6columns[ "Event Date Start"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["Event Date Start"] },
                                new () { ColumnId = Sheet6columns[ "Event End Date"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["Event End Date"] }
                                         }
                                    };
                                    newRows6.Add(newRow6);
                                }
                                await Task.Run(() => ApiCalls.ExpenseDetails(smartsheet, sheet6, newRows6));



                                Row addedrow = addedRows[0];
                                long ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role");
                                Cell UpdateB = new Cell
                                {
                                    ColumnId = ColumnId,
                                    Value = EventRequestsWebdt.Rows[loopcount]["Role"]
                                };
                                Row updateRows = new Row
                                {
                                    Id = addedrow.Id,
                                    Cells = new Cell[]
                                  { UpdateB }
                                };
                                //Cell? cellsToUpdate = addedrow.Cells.FirstOrDefault(c => c.ColumnId == ColumnId);
                                //if (cellsToUpdate != null) 
                                //{ 
                                //                cellsToUpdate.Value = formDataList.Webinar.Role; 
                                //            }
                                await Task.Run(() => ApiCalls.UpdateRole(smartsheet, sheet1, updateRows));

                                UpdateEventId.Rows.Add(new Object[] { EventRequestsWebdt.Rows[loopcount]["ID"], val });

                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.StackTrace);
                                Console.WriteLine(ex.Message);
                                Log.Error($"Error occured on Stall Fabrication method {ex.Message} at {DateTime.Now}");
                                Log.Error(ex.StackTrace);
                            }

                        }
                        else if (EventRequestsWebdt.Rows[loopcount]["Event Type"].ToString() == "Webinar")
                        {
                            try
                            {
                                Cell[] cellsToInsertRequestWeb = new Cell[]
                            {
                    new Cell { ColumnId =Sheet1columns["Approver Pre Event URL"],Value = EventRequestsWebdt.Rows[loopcount]["Approver Pre Event URL"]},
                    new Cell { ColumnId =Sheet1columns["Finance Treasury URL"],Value = EventRequestsWebdt.Rows[loopcount]["Finance Treasury URL"]},
                    new Cell { ColumnId =Sheet1columns["Initiator URL"],Value = EventRequestsWebdt.Rows[loopcount]["Initiator URL"]},
                    new Cell { ColumnId =Sheet1columns["Event Topic"],Value = EventRequestsWebdt.Rows[loopcount]["Event Topic"]},
                    new Cell { ColumnId =Sheet1columns["Event Type"],Value = EventRequestsWebdt.Rows[loopcount]["Event Type"]},
                    new Cell { ColumnId =Sheet1columns["Event Date"],Value = EventRequestsWebdt.Rows[loopcount]["Event Date"]},
                    new Cell { ColumnId =Sheet1columns["Start Time"],Value = EventRequestsWebdt.Rows[loopcount]["Start Time"]},
                    new Cell { ColumnId =Sheet1columns["End Time"],Value = EventRequestsWebdt.Rows[loopcount]["End Time"]},
                    new Cell { ColumnId =Sheet1columns["Meeting Type"],Value = EventRequestsWebdt.Rows[loopcount]["Meeting Type"]},
                    new Cell { ColumnId =Sheet1columns["Brands"],Value = EventRequestsWebdt.Rows[loopcount]["Brands"]},
                    new Cell { ColumnId =Sheet1columns["Expenses"],Value = EventRequestsWebdt.Rows[loopcount]["Expenses"]},
                    new Cell { ColumnId =Sheet1columns["Panelists"],Value = EventRequestsWebdt.Rows[loopcount]["Panelists"]},
                    new Cell { ColumnId =Sheet1columns["Invitees"],Value = EventRequestsWebdt.Rows[loopcount]["Invitees"]},
                    new Cell { ColumnId =Sheet1columns["MIPL Invitees"],Value = EventRequestsWebdt.Rows[loopcount]["MIPL Invitees"]},
                    new Cell { ColumnId =Sheet1columns["SlideKits"],Value = EventRequestsWebdt.Rows[loopcount]["SlideKits"]},
                    new Cell { ColumnId =Sheet1columns["IsAdvanceRequired"],Value = EventRequestsWebdt.Rows[loopcount]["IsAdvanceRequired"]},
                    new Cell { ColumnId =Sheet1columns["EventOpen30days"],Value = EventRequestsWebdt.Rows[loopcount]["EventOpen30days"]},
                    new Cell { ColumnId =Sheet1columns["EventWithin7days"],Value = EventRequestsWebdt.Rows[loopcount]["EventWithin7days"]},
                    new Cell { ColumnId =Sheet1columns["Initiator Name"],Value = EventRequestsWebdt.Rows[loopcount]["Initiator Name"]},
                    new Cell { ColumnId =Sheet1columns["Advance Amount"],Value = EventRequestsWebdt.Rows[loopcount]["Advance Amount"]},
                    new Cell { ColumnId =Sheet1columns[" Total Expense BTC"],Value = EventRequestsWebdt.Rows[loopcount]["Total Expense BTC"]},
                    new Cell { ColumnId =Sheet1columns["Total Expense BTE"],Value = EventRequestsWebdt.Rows[loopcount]["Total Expense BTE"]},
                    new Cell { ColumnId =Sheet1columns["Total Honorarium Amount"],Value = EventRequestsWebdt.Rows[loopcount]["Total Honorarium Amount"]},
                    new Cell { ColumnId =Sheet1columns["Total Travel Amount"],Value = EventRequestsWebdt.Rows[loopcount]["Total Travel Amount"]},
                    new Cell { ColumnId =Sheet1columns["Total Travel & Accommodation Amount"],Value = EventRequestsWebdt.Rows[loopcount]["Total Travel & Accommodation Amount"]},
                    new Cell { ColumnId =Sheet1columns["Total Accommodation Amount"],Value = EventRequestsWebdt.Rows[loopcount]["Total Accommodation Amount"]},
                    new Cell { ColumnId =Sheet1columns["Budget Amount"],Value = EventRequestsWebdt.Rows[loopcount]["Budget Amount"]},
                    new Cell { ColumnId =Sheet1columns["Total Local Conveyance"],Value = EventRequestsWebdt.Rows[loopcount]["Total Local Conveyance"]},
                    new Cell { ColumnId =Sheet1columns["Total Expense"],Value = EventRequestsWebdt.Rows[loopcount]["Total Expense"]},
                    new Cell { ColumnId =Sheet1columns["Initiator Email"],Value = EventRequestsWebdt.Rows[loopcount]["Initiator Email"]},
                    new Cell { ColumnId =Sheet1columns["RBM/BM"],Value = EventRequestsWebdt.Rows[loopcount]["RBM/BM"]},
                    new Cell { ColumnId =Sheet1columns["Sales Head"],Value = EventRequestsWebdt.Rows[loopcount]["Sales Head"]},
                    new Cell { ColumnId =Sheet1columns["Sales Coordinator"],Value = EventRequestsWebdt.Rows[loopcount]["Sales Coordinator"]},
                    new Cell { ColumnId =Sheet1columns["Marketing Coordinator"],Value = EventRequestsWebdt.Rows[loopcount]["Marketing Coordinator"]},
                    new Cell { ColumnId =Sheet1columns["Marketing Head"],Value = EventRequestsWebdt.Rows[loopcount]["Marketing Head"]},
                    new Cell { ColumnId =Sheet1columns["Compliance"],Value = EventRequestsWebdt.Rows[loopcount]["Compliance"]},
                    new Cell { ColumnId =Sheet1columns["Finance Accounts"],Value = EventRequestsWebdt.Rows[loopcount]["Finance Accounts"]},
                    new Cell { ColumnId =Sheet1columns["Finance Treasury"],Value = EventRequestsWebdt.Rows[loopcount]["Finance Treasury"]},
                    new Cell { ColumnId =Sheet1columns["Reporting Manager"],Value = EventRequestsWebdt.Rows[loopcount]["Reporting Manager"]},
                    new Cell { ColumnId =Sheet1columns["1 Up Manager"],Value = EventRequestsWebdt.Rows[loopcount]["1 Up Manager"]},
                    new Cell { ColumnId =Sheet1columns["Medical Affairs Head"],Value = EventRequestsWebdt.Rows[loopcount]["Medical Affairs Head"]},
                    new Cell { ColumnId =Sheet1columns["BTE Expense Details"],Value = EventRequestsWebdt.Rows[loopcount]["BTE Expense Details"]}
                                        };
                                Row row = new Row
                                {
                                    ToBottom = true,
                                    Cells = cellsToInsertRequestWeb,
                                };

                                IList<Row> addedRows = ApiCalls.WebDetails(smartsheet, sheet1, row);

                                long eventIdColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                                Cell? eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == Sheet1columns["EventId/EventRequestId"]);
                                string val = eventIdCell.DisplayValue;

                                int i = 1;
                                string[] eventattachment = EventRequestsWebdt.Rows[loopcount]["AttachmentPaths"].ToString().Split(',');

                                for (int attachcount = 1; attachcount < eventattachment.Length; attachcount++)
                                {
                                    Row addedRow = addedRows[0];
                                    Attachment attachment = await ApiCalls.AddAttachmentsToSheetSync(smartsheet, sheet1, addedRow, eventattachment[attachcount]);
                                    if (System.IO.File.Exists(eventattachment[attachcount]))
                                    {
                                        SheetHelper.DeleteFile(eventattachment[attachcount]);
                                    }
                                }

                                //Panel Sheet Insertion

                                for (int panelcount = 0; panelcount < EventRequestPanelDetailsdt.Rows.Count; panelcount++)
                                {
                                    Row newRow1 = new()
                                    {
                                        Cells = new List<Cell>()
                                    };
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["HcpRole"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["HcpRole"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["MISCode"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["MISCode"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Travel"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Travel"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["TotalSpend"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["TotalSpend"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Accomodation"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Accomodation"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["LocalConveyance"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["LocalConveyance"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["SpeakerCode"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["SpeakerCode"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["TrainerCode"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["TrainerCode"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["HonorariumRequired"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["HonorariumRequired"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["AgreementAmount"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["AgreementAmount"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["HonorariumAmount"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["HonorariumAmount"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Speciality"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Speciality"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Event Topic"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Event Topic"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Event Type"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Event Type"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Event Date Start"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Event Date Start"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Event End Date"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Event End Date"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["HCPName"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["HCPName"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["PAN card name"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["PAN card name"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["ExpenseType"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["ExpenseType"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Bank Account Number"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Bank Account Number"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Bank Name"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Bank Name"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["IFSC Code"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["IFSC Code"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["FCPA Date"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["FCPA Date"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Currency"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Currency"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Honorarium Amount Excluding Tax"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Honorarium Amount Excluding Tax"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Travel Excluding Tax"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Travel Excluding Tax"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Accomodation Excluding Tax"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Accomodation Excluding Tax"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Local Conveyance Excluding Tax"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Local Conveyance Excluding Tax"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["LC BTC/BTE"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["LC BTC/BTE"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Travel BTC/BTE"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Travel BTC/BTE"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Accomodation BTC/BTE"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Accomodation BTC/BTE"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Mode of Travel"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Mode of Travel"] });
                                    if (EventRequestPanelDetailsdt.Rows[panelcount]["Currency"] == "Others")
                                    {
                                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Other Currency"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Other Currency"] });
                                    }
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Beneficiary Name"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Beneficiary Name"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Pan Number"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Pan Number"] });

                                    if (EventRequestPanelDetailsdt.Rows[panelcount]["HcpRole"] == "Others")
                                    {

                                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Other Type"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Other Type"] });
                                    }

                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Tier"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Tier"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["HCP Type"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["HCP Type"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["PresentationDuration"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["PresentationDuration"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["PanelSessionPreparationDuration"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["PanelSessionPreparationDuration"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["PanelDiscussionDuration"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["PanelDiscussionDuration"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["QASessionDuration"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["QASessionDuration"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["BriefingSession"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["BriefingSession"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["TotalSessionHours"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["TotalSessionHours"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Rationale"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Rationale "] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["EventId/EventRequestId"], Value = val });

                                    IList<Row> panelrow = await Task.Run(() => ApiCalls.PanelDetails(smartsheet, sheet4, newRow1));

                                    string[] panelattachment = EventRequestPanelDetailsdt.Rows[panelcount]["AttachmentPaths"].ToString().Split(',');
                                    for (int panelattachcount = 1; panelattachcount < panelattachment.Length; panelattachcount++)
                                    {
                                        Row addedRow = panelrow[0];
                                        Attachment attachment = await ApiCalls.AddAttachmentsToSheetSync(smartsheet, sheet4, addedRow, panelattachment[panelattachcount]);
                                        if (System.IO.File.Exists(panelattachment[panelattachcount]))
                                        {
                                            SheetHelper.DeleteFile(panelattachment[panelattachcount]);
                                        }
                                    }
                                }

                                //Brand sheet insertion
                                List<Row> newRows2 = new();
                                for (int brandcount = 0; brandcount < EventRequestsBrandsListdt.Rows.Count; brandcount++)
                                {
                                    Row newRow2 = new()
                                    {
                                        Cells = new List<Cell>()
                            {
                            new(){ ColumnId = Sheet2columns[ "% Allocation"], Value = EventRequestsBrandsListdt.Rows[brandcount]["% Allocation"]},
                            new(){ ColumnId = Sheet2columns[ "Brands"], Value = EventRequestsBrandsListdt.Rows[brandcount]["Brands"] },
                            new(){ ColumnId = Sheet2columns[ "Project ID"], Value = EventRequestsBrandsListdt.Rows[brandcount]["Project ID"]},
                            new(){ ColumnId = Sheet2columns[ "EventId/EventRequestId"], Value = val }
                            }
                                    };

                                    newRows2.Add(newRow2);
                                }
                                await Task.Run(() => ApiCalls.BrandsDetails(smartsheet, sheet2, newRows2));

                                //Invitees sheet insertion 
                                List<Row> newRows3 = new();
                                for (int Inviteescount = 0; Inviteescount < EventRequestInviteesdt.Rows.Count; Inviteescount++)
                                {
                                    Row newRow3 = new()
                                    {
                                        Cells = new List<Cell>()
                            {
                            new () { ColumnId = Sheet3columns[ "HCPName"], Value = EventRequestInviteesdt.Rows[Inviteescount]["HCPName"] },
                            new () { ColumnId = Sheet3columns[ "Designation"], Value = EventRequestInviteesdt.Rows[Inviteescount]["Designation"] },
                            new () { ColumnId = Sheet3columns[ "Employee Code"], Value = EventRequestInviteesdt.Rows[Inviteescount]["Employee Code"] },
                            new () { ColumnId = Sheet3columns[ "LocalConveyance"], Value = EventRequestInviteesdt.Rows[Inviteescount]["LocalConveyance"] },
                            new () { ColumnId = Sheet3columns[ "BTC/BTE"], Value = EventRequestInviteesdt.Rows[Inviteescount]["BTC/BTE"] },
                            new () { ColumnId = Sheet3columns[ "LcAmount"], Value = EventRequestInviteesdt.Rows[Inviteescount]["LcAmount"] },
                            new () { ColumnId = Sheet3columns[ "Lc Amount Excluding Tax"], Value = EventRequestInviteesdt.Rows[Inviteescount]["Lc Amount Excluding Tax"] },
                            new () { ColumnId = Sheet3columns[ "EventId/EventRequestId"], Value = val },
                            new () { ColumnId = Sheet3columns[ "Invitee Source"], Value = EventRequestInviteesdt.Rows[Inviteescount]["Invitee Source"] },
                            new () { ColumnId = Sheet3columns[ "HCP Type"], Value = EventRequestInviteesdt.Rows[Inviteescount]["HCP Type"] },
                            new () { ColumnId = Sheet3columns[ "Speciality"], Value = EventRequestInviteesdt.Rows[Inviteescount]["Speciality"] },
                            new () { ColumnId = Sheet3columns[ "MISCode"], Value = EventRequestInviteesdt.Rows[Inviteescount]["MISCode"] },
                            new () { ColumnId = Sheet3columns[ "Event Topic"], Value = EventRequestInviteesdt.Rows[Inviteescount]["Event Topic"] },
                            new () { ColumnId = Sheet3columns[ "Event Type"], Value = EventRequestInviteesdt.Rows[Inviteescount]["Event Type"] },
                            new () { ColumnId = Sheet3columns[ "Event Date Start"], Value = EventRequestInviteesdt.Rows[Inviteescount]["Event Date Start"] },
                            new () { ColumnId = Sheet3columns[ "Event End Date"], Value = EventRequestInviteesdt.Rows[Inviteescount]["Event End Date"] }
                            }
                                    };
                                    newRows3.Add(newRow3);
                                }

                                await Task.Run(() => ApiCalls.InviteesDetails(smartsheet, sheet3, newRows3));


                                // Slidekits insertion
                                for (int SlideKitcount = 0; SlideKitcount < EventRequestHCPSlideKitDetailsdt.Rows.Count; SlideKitcount++)
                                {
                                    Row newRow5 = new()
                                    {
                                        Cells = new List<Cell>()
                                    };

                                    newRow5.Cells.Add(new Cell { ColumnId = Sheet5columns["MIS"], Value = EventRequestHCPSlideKitDetailsdt.Rows[SlideKitcount]["MIS"] });
                                    newRow5.Cells.Add(new Cell { ColumnId = Sheet5columns["Slide Kit Type"], Value = EventRequestHCPSlideKitDetailsdt.Rows[SlideKitcount]["Slide Kit Type"] });
                                    newRow5.Cells.Add(new Cell { ColumnId = Sheet5columns["SlideKit Document"], Value = EventRequestHCPSlideKitDetailsdt.Rows[SlideKitcount]["SlideKit Document"] });
                                    newRow5.Cells.Add(new Cell { ColumnId = Sheet5columns["EventId/EventRequestId"], Value = val });

                                    IList<Row> SlideKitrow = await Task.Run(() => ApiCalls.SlideKitDetails(smartsheet, sheet5, newRow5));

                                    string[] SlideKitattachment = EventRequestHCPSlideKitDetailsdt.Rows[SlideKitcount]["AttachmentPaths"].ToString().Split(',');
                                    for (int SlideKitattachcount = 1; SlideKitattachcount < SlideKitattachment.Length; SlideKitattachcount++)
                                    {
                                        Row addedRow = SlideKitrow[0];
                                        Attachment attachment = await ApiCalls.AddAttachmentsToSheetSync(smartsheet, sheet5, addedRow, SlideKitattachment[SlideKitattachcount]);
                                        if (System.IO.File.Exists(SlideKitattachment[SlideKitattachcount]))
                                        {
                                            SheetHelper.DeleteFile(SlideKitattachment[SlideKitattachcount]);
                                        }
                                    }
                                }

                                //ExpenseSheet Insertion
                                List<Row> newRows6 = new();
                                for (int Expensecount = 0; Expensecount < EventRequestExpensesSheetdt.Rows.Count; Expensecount++)
                                {
                                    Row newRow6 = new()
                                    {
                                        Cells = new List<Cell>()
                            {
                                new () { ColumnId = Sheet6columns[ "Expense"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["Expense"] },
                                new () { ColumnId = Sheet6columns[ "EventId/EventRequestID"], Value = val },
                                new () { ColumnId = Sheet6columns[ "AmountExcludingTax?"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["AmountExcludingTax?"] },
                                new () { ColumnId = Sheet6columns[ "Amount Excluding Tax"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["Amount Excluding Tax"] },
                                new () { ColumnId = Sheet6columns[ "Amount"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["Amount"] },
                                new () { ColumnId = Sheet6columns[ "BTC/BTE"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["BTC/BTE"] },
                                new () { ColumnId = Sheet6columns[ "BudgetAmount"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["BudgetAmount"] },
                                new () { ColumnId = Sheet6columns[ "BTCAmount"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["BTCAmount"] },
                                new () { ColumnId = Sheet6columns[ "BTEAmount"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["BTEAmount"] },
                                new () { ColumnId = Sheet6columns[ "Event Topic"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["Event Topic"] },
                                new () { ColumnId = Sheet6columns[ "Event Type"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["Event Type"] },
                                new () { ColumnId = Sheet6columns[ "Event Date Start"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["Event Date Start"] },
                                new () { ColumnId = Sheet6columns[ "Event End Date"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["Event End Date"] }
                            }
                                    };
                                    newRows6.Add(newRow6);
                                }
                                await Task.Run(() => ApiCalls.ExpenseDetails(smartsheet, sheet6, newRows6));

                                //Deviation insertion
                                for (int Deviationcount = 0; Deviationcount < Deviation_Processdt.Rows.Count; Deviationcount++)
                                {
                                    Row newRow7 = new()
                                    {
                                        Cells = new List<Cell>()
                                    };
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["EventId/EventRequestId"], Value = val });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Event Topic"], Value = Deviation_Processdt.Rows[Deviationcount]["Event Topic"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Event Type"], Value = Deviation_Processdt.Rows[Deviationcount]["Event Type"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Event Date"], Value = Deviation_Processdt.Rows[Deviationcount]["Event Date"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Start Time"], Value = Deviation_Processdt.Rows[Deviationcount]["Start Time"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["End Time"], Value = Deviation_Processdt.Rows[Deviationcount]["End Time"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["MIS Code"], Value = Deviation_Processdt.Rows[Deviationcount]["MIS Code"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["HCP Name"], Value = Deviation_Processdt.Rows[Deviationcount]["HCP Name"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Honorarium Amount"], Value = Deviation_Processdt.Rows[Deviationcount]["Honorarium Amount"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Travel & Accommodation Amount"], Value = Deviation_Processdt.Rows[Deviationcount]["Travel & Accommodation Amount"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Other Expenses"], Value = Deviation_Processdt.Rows[Deviationcount]["Other Expenses"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Deviation Type"], Value = Deviation_Processdt.Rows[Deviationcount]["Deviation Type"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["EventOpen45days"], Value = Deviation_Processdt.Rows[Deviationcount]["EventOpen45days"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Outstanding Events"], Value = Deviation_Processdt.Rows[Deviationcount]["Outstanding Events"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["EventWithin5days"], Value = Deviation_Processdt.Rows[Deviationcount]["EventWithin5days"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["PRE-F&B Expense Excluding Tax"], Value = Deviation_Processdt.Rows[Deviationcount]["PRE-F&B Expense Excluding Tax"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Travel/Accomodation 3,00,000 Exceeded Trigger"], Value = "Yes" });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Trainer Honorarium 12,00,000 Exceeded Trigger"], Value = "Yes" });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["HCP Honorarium 6,00,000 Exceeded Trigger"], Value = "Yes" });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Sales Head"], Value = Deviation_Processdt.Rows[Deviationcount]["Sales Head"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Finance Head"], Value = Deviation_Processdt.Rows[Deviationcount]["Finance Head"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Initiator Name"], Value = Deviation_Processdt.Rows[Deviationcount]["Initiator Name"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Initiator Email"], Value = Deviation_Processdt.Rows[Deviationcount]["Initiator Email"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Sales Coordinator"], Value = Deviation_Processdt.Rows[Deviationcount]["Sales Coordinator"] });

                                    IList<Row> addeddeviationrow = ApiCalls.DeviationData(smartsheet, sheet7, newRow7);

                                    string[] Deviationattachment = Deviation_Processdt.Rows[Deviationcount]["AttachmentPaths"].ToString().Split(',');
                                    for (int Deviationattachcount = 1; Deviationattachcount < Deviationattachment.Length; Deviationattachcount++)
                                    {
                                        Row addedRow = addeddeviationrow[0];
                                        Attachment attachment = await ApiCalls.AddAttachmentsToSheetSync(smartsheet, sheet7, addedRow, Deviationattachment[Deviationattachcount]);
                                        Attachment attachmentinmain = await ApiCalls.AddAttachmentsToSheetSync(smartsheet, sheet1, addedRows[0], Deviationattachment[Deviationattachcount]);
                                        if (System.IO.File.Exists(Deviationattachment[Deviationattachcount]))
                                        {
                                            SheetHelper.DeleteFile(Deviationattachment[Deviationattachcount]);
                                        }
                                    }
                                }

                                Row addedrow = addedRows[0];
                                long ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role");
                                Cell UpdateB = new Cell
                                {
                                    ColumnId = ColumnId,
                                    Value = EventRequestsWebdt.Rows[loopcount]["Role"]
                                };
                                Row updateRows = new Row
                                {
                                    Id = addedrow.Id,
                                    Cells = new Cell[]
                                  { UpdateB }
                                };
                                //Cell? cellsToUpdate = addedrow.Cells.FirstOrDefault(c => c.ColumnId == ColumnId);
                                //if (cellsToUpdate != null) 
                                //{ 
                                //                cellsToUpdate.Value = formDataList.Webinar.Role; 
                                //            }
                                await Task.Run(() => ApiCalls.UpdateRole(smartsheet, sheet1, updateRows));

                                UpdateEventId.Rows.Add(new Object[] { EventRequestsWebdt.Rows[loopcount]["ID"], val });

                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.StackTrace);
                                Console.WriteLine(ex.Message);
                                Log.Error($"Error occured on webinar method {ex.Message} at {DateTime.Now}");
                                Log.Error(ex.StackTrace);
                            }
                        }
                        else
                        {
                            try
                            {
                                Cell[] cellsToInsertRequestWeb = new Cell[]
                                {
                    new Cell { ColumnId =Sheet1columns["Approver Pre Event URL"],Value = EventRequestsWebdt.Rows[loopcount]["Approver Pre Event URL"]},
                    new Cell { ColumnId =Sheet1columns["Finance Treasury URL"],Value = EventRequestsWebdt.Rows[loopcount]["Finance Treasury URL"]},
                    new Cell { ColumnId =Sheet1columns["Initiator URL"],Value = EventRequestsWebdt.Rows[loopcount]["Initiator URL"]},
                    new Cell { ColumnId =Sheet1columns["Event Topic"],Value = EventRequestsWebdt.Rows[loopcount]["Event Topic"]},
                    new Cell { ColumnId =Sheet1columns["Event Type"],Value = EventRequestsWebdt.Rows[loopcount]["Event Type"]},
                    new Cell { ColumnId =Sheet1columns["Event Date"],Value = EventRequestsWebdt.Rows[loopcount]["Event Date"]},
                    new Cell { ColumnId =Sheet1columns["Start Time"],Value = EventRequestsWebdt.Rows[loopcount]["Start Time"]},
                    new Cell { ColumnId =Sheet1columns["End Time"],Value = EventRequestsWebdt.Rows[loopcount]["End Time"]},
                    new Cell { ColumnId =Sheet1columns["Meeting Type"],Value = EventRequestsWebdt.Rows[loopcount]["Meeting Type"]},
                    new Cell { ColumnId =Sheet1columns["Brands"],Value = EventRequestsWebdt.Rows[loopcount]["Brands"]},
                    new Cell { ColumnId =Sheet1columns["Expenses"],Value = EventRequestsWebdt.Rows[loopcount]["Expenses"]},
                    new Cell { ColumnId =Sheet1columns["Panelists"],Value = EventRequestsWebdt.Rows[loopcount]["Panelists"]},
                    new Cell { ColumnId =Sheet1columns["Invitees"],Value = EventRequestsWebdt.Rows[loopcount]["Invitees"]},
                    new Cell { ColumnId =Sheet1columns["MIPL Invitees"],Value = EventRequestsWebdt.Rows[loopcount]["MIPL Invitees"]},
                    new Cell { ColumnId =Sheet1columns["SlideKits"],Value = EventRequestsWebdt.Rows[loopcount]["SlideKits"]},
                    new Cell { ColumnId =Sheet1columns["IsAdvanceRequired"],Value = EventRequestsWebdt.Rows[loopcount]["IsAdvanceRequired"]},
                    new Cell { ColumnId =Sheet1columns["EventOpen30days"],Value = EventRequestsWebdt.Rows[loopcount]["EventOpen30days"]},
                    new Cell { ColumnId =Sheet1columns["EventWithin7days"],Value = EventRequestsWebdt.Rows[loopcount]["EventWithin7days"]},
                    new Cell { ColumnId =Sheet1columns["Initiator Name"],Value = EventRequestsWebdt.Rows[loopcount]["Initiator Name"]},
                    new Cell { ColumnId =Sheet1columns["Advance Amount"],Value = EventRequestsWebdt.Rows[loopcount]["Advance Amount"]},
                    new Cell { ColumnId =Sheet1columns[" Total Expense BTC"],Value = EventRequestsWebdt.Rows[loopcount]["Total Expense BTC"]},
                    new Cell { ColumnId =Sheet1columns["Total Expense BTE"],Value = EventRequestsWebdt.Rows[loopcount]["Total Expense BTE"]},
                    new Cell { ColumnId =Sheet1columns["Total Honorarium Amount"],Value = EventRequestsWebdt.Rows[loopcount]["Total Honorarium Amount"]},
                    new Cell { ColumnId =Sheet1columns["Total Travel Amount"],Value = EventRequestsWebdt.Rows[loopcount]["Total Travel Amount"]},
                    new Cell { ColumnId =Sheet1columns["Total Travel & Accommodation Amount"],Value = EventRequestsWebdt.Rows[loopcount]["Total Travel & Accommodation Amount"]},
                    new Cell { ColumnId =Sheet1columns["Total Accommodation Amount"],Value = EventRequestsWebdt.Rows[loopcount]["Total Accommodation Amount"]},
                    new Cell { ColumnId =Sheet1columns["Budget Amount"],Value = EventRequestsWebdt.Rows[loopcount]["Budget Amount"]},
                    new Cell { ColumnId =Sheet1columns["Total Local Conveyance"],Value = EventRequestsWebdt.Rows[loopcount]["Total Local Conveyance"]},
                    new Cell { ColumnId =Sheet1columns["Total Expense"],Value = EventRequestsWebdt.Rows[loopcount]["Total Expense"]},
                    new Cell { ColumnId =Sheet1columns["Initiator Email"],Value = EventRequestsWebdt.Rows[loopcount]["Initiator Email"]},
                    new Cell { ColumnId =Sheet1columns["RBM/BM"],Value = EventRequestsWebdt.Rows[loopcount]["RBM/BM"]},
                    new Cell { ColumnId =Sheet1columns["Sales Head"],Value = EventRequestsWebdt.Rows[loopcount]["Sales Head"]},
                    new Cell { ColumnId =Sheet1columns["Sales Coordinator"],Value = EventRequestsWebdt.Rows[loopcount]["Sales Coordinator"]},
                    new Cell { ColumnId =Sheet1columns["Marketing Coordinator"],Value = EventRequestsWebdt.Rows[loopcount]["Marketing Coordinator"]},
                    new Cell { ColumnId =Sheet1columns["Marketing Head"],Value = EventRequestsWebdt.Rows[loopcount]["Marketing Head"]},
                    new Cell { ColumnId =Sheet1columns["Compliance"],Value = EventRequestsWebdt.Rows[loopcount]["Compliance"]},
                    new Cell { ColumnId =Sheet1columns["Finance Accounts"],Value = EventRequestsWebdt.Rows[loopcount]["Finance Accounts"]},
                    new Cell { ColumnId =Sheet1columns["Finance Treasury"],Value = EventRequestsWebdt.Rows[loopcount]["Finance Treasury"]},
                    new Cell { ColumnId =Sheet1columns["Reporting Manager"],Value = EventRequestsWebdt.Rows[loopcount]["Reporting Manager"]},
                    new Cell { ColumnId =Sheet1columns["1 Up Manager"],Value = EventRequestsWebdt.Rows[loopcount]["1 Up Manager"]},
                    new Cell { ColumnId =Sheet1columns["Medical Affairs Head"],Value = EventRequestsWebdt.Rows[loopcount]["Medical Affairs Head"]},
                    new Cell { ColumnId =Sheet1columns["BTE Expense Details"],Value = EventRequestsWebdt.Rows[loopcount]["BTE Expense Details"]},
                    new Cell { ColumnId =Sheet1columns["Venue Name"], Value =EventRequestsWebdt.Rows[loopcount]["Venue Name"]},
                    new Cell { ColumnId =Sheet1columns["City"], Value = EventRequestsWebdt.Rows[loopcount]["City"]},
                    new Cell { ColumnId =Sheet1columns["State"], Value = EventRequestsWebdt.Rows[loopcount]["State"]}
                            };
                                Row row = new Row
                                {
                                    ToBottom = true,
                                    Cells = cellsToInsertRequestWeb,
                                };

                                IList<Row> addedRows = ApiCalls.WebDetails(smartsheet, sheet1, row);

                                long eventIdColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId");
                                Cell? eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == Sheet1columns["EventId/EventRequestId"]);
                                string val = eventIdCell.DisplayValue;

                                int i = 1;
                                string[] eventattachment = EventRequestsWebdt.Rows[loopcount]["AttachmentPaths"].ToString().Split(',');

                                for (int attachcount = 1; attachcount < eventattachment.Length; attachcount++)
                                {
                                    Row addedRow = addedRows[0];
                                    Attachment attachment = await ApiCalls.AddAttachmentsToSheetSync(smartsheet, sheet1, addedRow, eventattachment[attachcount]);
                                    if (System.IO.File.Exists(eventattachment[attachcount]))
                                    {
                                        SheetHelper.DeleteFile(eventattachment[attachcount]);
                                    }
                                }

                                //Panel Sheet Insertion

                                for (int panelcount = 0; panelcount < EventRequestPanelDetailsdt.Rows.Count; panelcount++)
                                {
                                    Row newRow1 = new()
                                    {
                                        Cells = new List<Cell>()
                                    };
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["HcpRole"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["HcpRole"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["MISCode"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["MISCode"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Travel"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Travel"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["TotalSpend"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["TotalSpend"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Accomodation"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Accomodation"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["LocalConveyance"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["LocalConveyance"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["SpeakerCode"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["SpeakerCode"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["TrainerCode"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["TrainerCode"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["HonorariumRequired"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["HonorariumRequired"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["AgreementAmount"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["AgreementAmount"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["HonorariumAmount"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["HonorariumAmount"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Speciality"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Speciality"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Event Topic"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Event Topic"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Event Type"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Event Type"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Event Date Start"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Event Date Start"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Event End Date"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Event End Date"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["HCPName"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["HCPName"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["PAN card name"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["PAN card name"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["ExpenseType"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["ExpenseType"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Bank Account Number"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Bank Account Number"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Bank Name"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Bank Name"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["IFSC Code"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["IFSC Code"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["FCPA Date"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["FCPA Date"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Currency"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Currency"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Honorarium Amount Excluding Tax"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Honorarium Amount Excluding Tax"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Travel Excluding Tax"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Travel Excluding Tax"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Accomodation Excluding Tax"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Accomodation Excluding Tax"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Local Conveyance Excluding Tax"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Local Conveyance Excluding Tax"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["LC BTC/BTE"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["LC BTC/BTE"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Travel BTC/BTE"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Travel BTC/BTE"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Accomodation BTC/BTE"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Accomodation BTC/BTE"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Mode of Travel"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Mode of Travel"] });
                                    if (EventRequestPanelDetailsdt.Rows[panelcount]["Currency"] == "Others")
                                    {
                                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Other Currency"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Other Currency"] });
                                    }
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Beneficiary Name"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Beneficiary Name"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Pan Number"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Pan Number"] });

                                    if (EventRequestPanelDetailsdt.Rows[panelcount]["HcpRole"] == "Others")
                                    {

                                        newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Other Type"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Other Type"] });
                                    }

                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Tier"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Tier"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["HCP Type"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["HCP Type"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["PresentationDuration"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["PresentationDuration"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["PanelSessionPreparationDuration"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["PanelSessionPreparationDuration"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["PanelDiscussionDuration"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["PanelDiscussionDuration"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["QASessionDuration"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["QASessionDuration"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["BriefingSession"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["BriefingSession"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["TotalSessionHours"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["TotalSessionHours"] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Rationale"], Value = EventRequestPanelDetailsdt.Rows[panelcount]["Rationale "] });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["EventId/EventRequestId"], Value = val });
                                    newRow1.Cells.Add(new Cell { ColumnId = Sheet4columns["Venue name"], Value = EventRequestsWebdt.Rows[loopcount]["Venue Name"] });

                                    IList<Row> panelrow = await Task.Run(() => ApiCalls.PanelDetails(smartsheet, sheet4, newRow1));

                                    string[] panelattachment = EventRequestPanelDetailsdt.Rows[panelcount]["AttachmentPaths"].ToString().Split(',');
                                    for (int panelattachcount = 1; panelattachcount < panelattachment.Length; panelattachcount++)
                                    {
                                        Row addedRow = panelrow[0];
                                        Attachment attachment = await ApiCalls.AddAttachmentsToSheetSync(smartsheet, sheet4, addedRow, panelattachment[panelattachcount]);
                                        if (System.IO.File.Exists(panelattachment[panelattachcount]))
                                        {
                                            SheetHelper.DeleteFile(panelattachment[panelattachcount]);
                                        }
                                    }
                                }

                                //Brand sheet insertion
                                List<Row> newRows2 = new();
                                for (int brandcount = 0; brandcount < EventRequestsBrandsListdt.Rows.Count; brandcount++)
                                {
                                    Row newRow2 = new()
                                    {
                                        Cells = new List<Cell>()
                            {
                            new(){ ColumnId = Sheet2columns[ "% Allocation"], Value = EventRequestsBrandsListdt.Rows[brandcount]["% Allocation"]},
                            new(){ ColumnId = Sheet2columns[ "Brands"], Value = EventRequestsBrandsListdt.Rows[brandcount]["Brands"] },
                            new(){ ColumnId = Sheet2columns[ "Project ID"], Value = EventRequestsBrandsListdt.Rows[brandcount]["Project ID"]},
                            new(){ ColumnId = Sheet2columns[ "EventId/EventRequestId"], Value = val }
                            }
                                    };

                                    newRows2.Add(newRow2);
                                }
                                await Task.Run(() => ApiCalls.BrandsDetails(smartsheet, sheet2, newRows2));

                                //Invitees sheet insertion 
                                List<Row> newRows3 = new();
                                for (int Inviteescount = 0; Inviteescount < EventRequestInviteesdt.Rows.Count; Inviteescount++)
                                {
                                    Row newRow3 = new()
                                    {
                                        Cells = new List<Cell>()
                            {
                            new () { ColumnId = Sheet3columns[ "HCPName"], Value = EventRequestInviteesdt.Rows[Inviteescount]["HCPName"] },
                            new () { ColumnId = Sheet3columns[ "Designation"], Value = EventRequestInviteesdt.Rows[Inviteescount]["Designation"] },
                            new () { ColumnId = Sheet3columns[ "Employee Code"], Value = EventRequestInviteesdt.Rows[Inviteescount]["Employee Code"] },
                            new () { ColumnId = Sheet3columns[ "LocalConveyance"], Value = EventRequestInviteesdt.Rows[Inviteescount]["LocalConveyance"] },
                            new () { ColumnId = Sheet3columns[ "BTC/BTE"], Value = EventRequestInviteesdt.Rows[Inviteescount]["BTC/BTE"] },
                            new () { ColumnId = Sheet3columns[ "LcAmount"], Value = EventRequestInviteesdt.Rows[Inviteescount]["LcAmount"] },
                            new () { ColumnId = Sheet3columns[ "Lc Amount Excluding Tax"], Value = EventRequestInviteesdt.Rows[Inviteescount]["Lc Amount Excluding Tax"] },
                            new () { ColumnId = Sheet3columns[ "EventId/EventRequestId"], Value = val },
                            new () { ColumnId = Sheet3columns[ "Invitee Source"], Value = EventRequestInviteesdt.Rows[Inviteescount]["Invitee Source"] },
                            new () { ColumnId = Sheet3columns[ "HCP Type"], Value = EventRequestInviteesdt.Rows[Inviteescount]["HCP Type"] },
                            new () { ColumnId = Sheet3columns[ "Speciality"], Value = EventRequestInviteesdt.Rows[Inviteescount]["Speciality"] },
                            new () { ColumnId = Sheet3columns[ "MISCode"], Value = EventRequestInviteesdt.Rows[Inviteescount]["MISCode"] },
                            new () { ColumnId = Sheet3columns[ "Event Topic"], Value = EventRequestInviteesdt.Rows[Inviteescount]["Event Topic"] },
                            new () { ColumnId = Sheet3columns[ "Event Type"], Value = EventRequestInviteesdt.Rows[Inviteescount]["Event Type"] },
                            new () { ColumnId = Sheet3columns[ "Event Date Start"], Value = EventRequestInviteesdt.Rows[Inviteescount]["Event Date Start"] },
                            new () { ColumnId = Sheet3columns[ "Event End Date"], Value = EventRequestInviteesdt.Rows[Inviteescount]["Event End Date"] }
                            }
                                    };
                                    newRows3.Add(newRow3);
                                }

                                await Task.Run(() => ApiCalls.InviteesDetails(smartsheet, sheet3, newRows3));


                                // Slidekits insertion
                                for (int SlideKitcount = 0; SlideKitcount < EventRequestHCPSlideKitDetailsdt.Rows.Count; SlideKitcount++)
                                {
                                    Row newRow5 = new()
                                    {
                                        Cells = new List<Cell>()
                                    };

                                    newRow5.Cells.Add(new Cell { ColumnId = Sheet5columns["MIS"], Value = EventRequestHCPSlideKitDetailsdt.Rows[SlideKitcount]["MIS"] });
                                    newRow5.Cells.Add(new Cell { ColumnId = Sheet5columns["Slide Kit Type"], Value = EventRequestHCPSlideKitDetailsdt.Rows[SlideKitcount]["Slide Kit Type"] });
                                    newRow5.Cells.Add(new Cell { ColumnId = Sheet5columns["SlideKit Document"], Value = EventRequestHCPSlideKitDetailsdt.Rows[SlideKitcount]["SlideKit Document"] });
                                    newRow5.Cells.Add(new Cell { ColumnId = Sheet5columns["EventId/EventRequestId"], Value = val });

                                    IList<Row> SlideKitrow = await Task.Run(() => ApiCalls.SlideKitDetails(smartsheet, sheet5, newRow5));

                                    string[] SlideKitattachment = EventRequestHCPSlideKitDetailsdt.Rows[SlideKitcount]["AttachmentPaths"].ToString().Split(',');
                                    for (int SlideKitattachcount = 1; SlideKitattachcount < SlideKitattachment.Length; SlideKitattachcount++)
                                    {
                                        Row addedRow = SlideKitrow[0];
                                        Attachment attachment = await ApiCalls.AddAttachmentsToSheetSync(smartsheet, sheet5, addedRow, SlideKitattachment[SlideKitattachcount]);
                                        if (System.IO.File.Exists(SlideKitattachment[SlideKitattachcount]))
                                        {
                                            SheetHelper.DeleteFile(SlideKitattachment[SlideKitattachcount]);
                                        }
                                    }
                                }

                                //ExpenseSheet Insertion
                                List<Row> newRows6 = new();
                                for (int Expensecount = 0; Expensecount < EventRequestExpensesSheetdt.Rows.Count; Expensecount++)
                                {
                                    Row newRow6 = new()
                                    {
                                        Cells = new List<Cell>()
                            {
                                new () { ColumnId = Sheet6columns[ "Expense"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["Expense"] },
                                new () { ColumnId = Sheet6columns[ "EventId/EventRequestID"], Value = val },
                                new () { ColumnId = Sheet6columns[ "AmountExcludingTax?"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["AmountExcludingTax?"] },
                                new () { ColumnId = Sheet6columns[ "Amount Excluding Tax"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["Amount Excluding Tax"] },
                                new () { ColumnId = Sheet6columns[ "Amount"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["Amount"] },
                                new () { ColumnId = Sheet6columns[ "BTC/BTE"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["BTC/BTE"] },
                                new () { ColumnId = Sheet6columns[ "BudgetAmount"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["BudgetAmount"] },
                                new () { ColumnId = Sheet6columns[ "BTCAmount"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["BTCAmount"] },
                                new () { ColumnId = Sheet6columns[ "BTEAmount"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["BTEAmount"] },
                                new () { ColumnId = Sheet6columns[ "Event Topic"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["Event Topic"] },
                                new () { ColumnId = Sheet6columns[ "Event Type"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["Event Type"] },
                                new () { ColumnId = Sheet6columns[ "Event Date Start"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["Event Date Start"] },
                                new () { ColumnId = Sheet6columns[ "Event End Date"], Value = EventRequestExpensesSheetdt.Rows[Expensecount]["Event End Date"] },
                                new () { ColumnId = Sheet6columns[ "Venue name"], Value = EventRequestsWebdt.Rows[loopcount]["Venue Name"]}
                                }
                                    };
                                    newRows6.Add(newRow6);
                                }
                                await Task.Run(() => ApiCalls.ExpenseDetails(smartsheet, sheet6, newRows6));

                                //Deviation insertion
                                for (int Deviationcount = 0; Deviationcount < Deviation_Processdt.Rows.Count; Deviationcount++)
                                {
                                    Row newRow7 = new()
                                    {
                                        Cells = new List<Cell>()
                                    };
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["EventId/EventRequestId"], Value = val });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Event Topic"], Value = Deviation_Processdt.Rows[Deviationcount]["Event Topic"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Event Type"], Value = Deviation_Processdt.Rows[Deviationcount]["Event Type"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Event Date"], Value = Deviation_Processdt.Rows[Deviationcount]["Event Date"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Start Time"], Value = Deviation_Processdt.Rows[Deviationcount]["Start Time"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["End Time"], Value = Deviation_Processdt.Rows[Deviationcount]["End Time"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["MIS Code"], Value = Deviation_Processdt.Rows[Deviationcount]["MIS Code"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["HCP Name"], Value = Deviation_Processdt.Rows[Deviationcount]["HCP Name"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Honorarium Amount"], Value = Deviation_Processdt.Rows[Deviationcount]["Honorarium Amount"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Travel & Accommodation Amount"], Value = Deviation_Processdt.Rows[Deviationcount]["Travel & Accommodation Amount"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Other Expenses"], Value = Deviation_Processdt.Rows[Deviationcount]["Other Expenses"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Deviation Type"], Value = Deviation_Processdt.Rows[Deviationcount]["Deviation Type"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["EventOpen45days"], Value = Deviation_Processdt.Rows[Deviationcount]["EventOpen45days"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Outstanding Events"], Value = Deviation_Processdt.Rows[Deviationcount]["Outstanding Events"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["EventWithin5days"], Value = Deviation_Processdt.Rows[Deviationcount]["EventWithin5days"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["PRE-F&B Expense Excluding Tax"], Value = Deviation_Processdt.Rows[Deviationcount]["PRE-F&B Expense Excluding Tax"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Travel/Accomodation 3,00,000 Exceeded Trigger"], Value = "Yes" });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Trainer Honorarium 12,00,000 Exceeded Trigger"], Value = "Yes" });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["HCP Honorarium 6,00,000 Exceeded Trigger"], Value = "Yes" });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Sales Head"], Value = Deviation_Processdt.Rows[Deviationcount]["Sales Head"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Finance Head"], Value = Deviation_Processdt.Rows[Deviationcount]["Finance Head"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Initiator Name"], Value = Deviation_Processdt.Rows[Deviationcount]["Initiator Name"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Initiator Email"], Value = Deviation_Processdt.Rows[Deviationcount]["Initiator Email"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Sales Coordinator"], Value = Deviation_Processdt.Rows[Deviationcount]["Sales Coordinator"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["Venue Name"], Value = EventRequestsWebdt.Rows[loopcount]["Venue Name"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["City"], Value = EventRequestsWebdt.Rows[loopcount]["City"] });
                                    newRow7.Cells.Add(new Cell { ColumnId = Sheet7columns["State"], Value = EventRequestsWebdt.Rows[loopcount]["State"] });

                                    IList<Row> addeddeviationrow = ApiCalls.DeviationData(smartsheet, sheet7, newRow7);

                                    string[] Deviationattachment = Deviation_Processdt.Rows[Deviationcount]["AttachmentPaths"].ToString().Split(',');
                                    for (int Deviationattachcount = 1; Deviationattachcount < Deviationattachment.Length; Deviationattachcount++)
                                    {
                                        Row addedRow = addeddeviationrow[0];
                                        Attachment attachment = await ApiCalls.AddAttachmentsToSheetSync(smartsheet, sheet7, addedRow, Deviationattachment[Deviationattachcount]);
                                        Attachment attachmentinmain = await ApiCalls.AddAttachmentsToSheetSync(smartsheet, sheet1, addedRows[0], Deviationattachment[Deviationattachcount]);
                                        if (System.IO.File.Exists(Deviationattachment[Deviationattachcount]))
                                        {
                                            SheetHelper.DeleteFile(Deviationattachment[Deviationattachcount]);
                                        }
                                    }
                                }

                                Row addedrow = addedRows[0];
                                long ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Role");
                                Cell UpdateB = new Cell
                                {
                                    ColumnId = ColumnId,
                                    Value = EventRequestsWebdt.Rows[loopcount]["Role"]
                                };
                                Row updateRows = new Row
                                {
                                    Id = addedrow.Id,
                                    Cells = new Cell[]
                                  { UpdateB }
                                };
                                //Cell? cellsToUpdate = addedrow.Cells.FirstOrDefault(c => c.ColumnId == ColumnId);
                                //if (cellsToUpdate != null) 
                                //{ 
                                //                cellsToUpdate.Value = formDataList.Webinar.Role; 
                                //            }
                                await Task.Run(() => ApiCalls.UpdateRole(smartsheet, sheet1, updateRows));

                                UpdateEventId.Rows.Add(new Object[] { EventRequestsWebdt.Rows[loopcount]["ID"], val });

                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.StackTrace);
                                Console.WriteLine(ex.Message);
                                Log.Error($"Error occured on Class1 Scheduler method {ex.Message} at {DateTime.Now}");
                                Log.Error(ex.StackTrace);
                            }
                        }
                    }
                }

                if (UpdateEventId.Rows.Count > 0)
                {
                    await MyConn.OpenAsync();
                    MySqlCommand Updatecommand = new MySqlCommand("UpdateEventid", MyConn);
                    Updatecommand.CommandType = CommandType.StoredProcedure;
                    for (int syncloop = 0; syncloop < UpdateEventId.Rows.Count; syncloop++)
                    {
                        Updatecommand.Parameters.AddWithValue("@Dbid", UpdateEventId.Rows[syncloop][0]);
                        Updatecommand.Parameters.AddWithValue("@Eventid", UpdateEventId.Rows[syncloop][1]);
                        Updatecommand.ExecuteNonQuery();
                        Updatecommand.Parameters.Clear();
                    }
                    await MyConn.CloseAsync();
                }
                return Ok(new { Message = " Success!" });
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.StackTrace);
                Console.WriteLine(ex.Message);
                Log.Error($"Error occured on webinar method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(new { Message = ex.Message + "------" + ex.StackTrace });
            }
            return Ok();
        }

    }
}