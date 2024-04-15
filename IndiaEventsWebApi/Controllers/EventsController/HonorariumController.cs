using IndiaEvents.Models.Models.EventTypeSheets;
using IndiaEventsWebApi.Helper;
using IndiaEventsWebApi.Models.EventTypeSheets;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Serilog;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Globalization;
using System.Text;

namespace IndiaEventsWebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class HonorariumController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;
        private readonly SmartsheetClient smartsheet;
        //private readonly Sheet sheet1;
        //private readonly Sheet sheet2;
        //private readonly Sheet sheet3;
        //private readonly Sheet sheet4;
        //private readonly Sheet sheet5;
        //private readonly Sheet sheet6;
        //private readonly Sheet sheet7;

        public HonorariumController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
            smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
        }

        [HttpPost("AddHonorariumData")]
        public IActionResult AddHonorariumData(HonorariumPaymentListPh2 formData)
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:HonorariumPayment").Value;
                string sheetId1 = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;
                string sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;

                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId1);
                Sheet sheet7 = SheetHelper.GetSheetById(smartsheet, sheetId7);

                StringBuilder addedHcpData = new();
                int addedHcpDataNo = 1;

                foreach (var i in formData.HcpRoles)
                {
                    string rowData = $"{addedHcpDataNo}. Name:{i.HcpName} | Role:{i.HcpRole} |Code:{i.MisCode} | HCP Type:{i.GOorNGO}| Including GST:{i.IsInclidingGst}| Agreement Amount:{i.AgreementAmount} |Annual Trainer Agreement Valid: {i.IsAnnualTrainerAgreementValid} ";
                    addedHcpData.AppendLine(rowData);
                    addedHcpDataNo++;
                }
                string panalist = addedHcpData.ToString();
                Row newRow = new()
                {
                    Cells = new List<Cell>()
                };
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "SlideKits"), Value = formData.RequestHonorariumList.SlideKits });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventId/EventRequestId"), Value = formData.RequestHonorariumList.EventId });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Event Type"), Value = formData.RequestHonorariumList.EventType });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Event Date"), Value = formData.RequestHonorariumList.EventDate });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Event Topic"), Value = formData.RequestHonorariumList.EventTopic });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "City"), Value = formData.RequestHonorariumList.City });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "State"), Value = formData.RequestHonorariumList.State });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Start Time"), Value = formData.RequestHonorariumList.StartTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "End Time"), Value = formData.RequestHonorariumList.EndTime });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Venue Name"), Value = formData.RequestHonorariumList.VenueName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Travel & Accommodation Amount"), Value = formData.RequestHonorariumList.TotalTravelAndAccomodationSpend });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Honorarium Amount"), Value = formData.RequestHonorariumList.TotalHonorariumSpend });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Budget"), Value = formData.RequestHonorariumList.TotalSpend });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Expenses"), Value = formData.RequestHonorariumList.Expenses });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Travel Amount"), Value = formData.RequestHonorariumList.TotalTravelSpend });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Accommodation Amount"), Value = formData.RequestHonorariumList.TotalAccomodationSpend });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Expense"), Value = formData.RequestHonorariumList.TotalExpenses });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Total Local Conveyance"), Value = formData.RequestHonorariumList.TotalLocalConveyance });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Brands"), Value = formData.RequestHonorariumList.Brands });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Invitees"), Value = formData.RequestHonorariumList.Invitees });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Panelists"), Value = formData.RequestHonorariumList.Panelists });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Initiator Name"), Value = formData.RequestHonorariumList.InitiatorName });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Panelists & Agreements"), Value = panalist });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Honorarium Submitted?"), Value = formData.RequestHonorariumList.HonarariumSubmitted });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Initiator Email"), Value = formData.RequestHonorariumList.InitiatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "RBM/BM"), Value = formData.RequestHonorariumList.RBMorBM });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Sales Head"), Value = formData.RequestHonorariumList.SalesHeadEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Sales Coordinator"), Value = formData.RequestHonorariumList.SalesCoordinatorEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Marketing Head"), Value = formData.RequestHonorariumList.MarketingHeadEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Compliance"), Value = formData.RequestHonorariumList.Compliance });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance Accounts"), Value = formData.RequestHonorariumList.FinanceAccounts });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance Treasury"), Value = formData.RequestHonorariumList.FinanceTreasury });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Reporting Manager"), Value = formData.RequestHonorariumList.ReportingManagerEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "1 Up Manager"), Value = formData.RequestHonorariumList.FirstLevelEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Medical Affairs Head"), Value = formData.RequestHonorariumList.MedicalAffairsEmail });
                newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Role"), Value = formData.RequestHonorariumList.Role });

                IList<Row> addedRows = smartsheet.SheetResources.RowResources.AddRows(sheet.Id.Value, new Row[] { newRow });
                string? eventId = formData.RequestHonorariumList.EventId;
                foreach (string p in formData.RequestHonorariumList.Files)
                {
                    string[] words = p.Split(':');
                    string r = words[0];
                    string q = words[1];
                    string name = r.Split(".")[0];
                    string filePath = SheetHelper.testingFile(q, name);
                    Row addedRow = addedRows[0];
                    Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                           sheet.Id.Value, addedRow.Id.Value, filePath, "application/msword");

                    if (System.IO.File.Exists(filePath))
                    {
                        SheetHelper.DeleteFile(filePath);
                    }
                }

                if (formData.RequestHonorariumList.IsDeviationUpload == "Yes")
                {
                    try
                    {
                        Row newRow7 = new()
                        {
                            Cells = new List<Cell>()
                        };
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventId/EventRequestId"), Value = eventId });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Event Topic"), Value = formData.RequestHonorariumList.EventTopic });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventType"), Value = formData.RequestHonorariumList.EventType });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EventDate"), Value = formData.RequestHonorariumList.EventDate });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "StartTime"), Value = formData.RequestHonorariumList.StartTime });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "EndTime"), Value = formData.RequestHonorariumList.EndTime });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "VenueName"), Value = formData.RequestHonorariumList.VenueName });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "City"), Value = formData.RequestHonorariumList.City });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "State"), Value = formData.RequestHonorariumList.State });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Deviation Type"), Value = configuration.GetSection("DeviationNamesInHonorarium:5WorkingdaysDeviationDateTrigger").Value });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "HON-5Workingdays Deviation Date Trigger"), Value = formData.RequestHonorariumList.IsDeviationUpload });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Sales Head"), Value = formData.RequestHonorariumList.Compliance });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Finance Head"), Value = formData.RequestHonorariumList.FinanceHead });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "InitiatorName"), Value = formData.RequestHonorariumList.InitiatorName });
                        newRow7.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet7, "Initiator Email"), Value = formData.RequestHonorariumList.InitiatorEmail });

                        IList<Row> addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(sheet7.Id.Value, new Row[] { newRow7 });

                        foreach (string p in formData.RequestHonorariumList.DeviationFiles)
                        {
                            string[] words = p.Split(':');
                            string r = words[0];
                            string q = words[1];
                            string name = r.Split("*")[0];
                            string filePath = SheetHelper.testingFile(q, name);
                            Row addedRow = addeddeviationrow[0];
                            Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet7.Id.Value, addedRow.Id.Value, filePath, "application/msword");


                            if (System.IO.File.Exists(filePath))
                            {
                                SheetHelper.DeleteFile(filePath);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.Error($"Error occured on HonorariumController method {ex.Message} at {DateTime.Now}");
                        Log.Error(ex.StackTrace);
                        return BadRequest(ex.Message);
                    }
                }

                Row? targetRow = sheet1.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == eventId));
                if (targetRow != null)
                {
                    long honorariumSubmittedColumnId = SheetHelper.GetColumnIdByName(sheet1, "Honorarium Submitted?");
                    Cell cellToUpdateB = new Cell
                    { ColumnId = honorariumSubmittedColumnId, Value = "Yes" };
                    Row updateRow = new Row { Id = targetRow.Id, Cells = new Cell[] { cellToUpdateB } };
                    Cell? cellToUpdate = targetRow.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                    if (cellToUpdate != null) { cellToUpdate.Value = "Yes"; }
                    smartsheet.SheetResources.RowResources.UpdateRows(sheet1.Id.Value, new Row[] { updateRow });
                }
                return Ok(new
                { Message = "Data added successfully." });
            }
            catch (Exception ex)
            {
                Log.Error($"Error occured on HonorariumController method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(ex.Message);
            }
        }


        [HttpPut("HonorariumUpdate")]
        public IActionResult HonorariumUpdate(List<HonorariumUpdate> formDataArray)
        {
            try
            {
                string sheetId4 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;

                Sheet sheet4 = SheetHelper.GetSheetById(smartsheet, sheetId4);
                foreach (var formdata in formDataArray)
                {
                    var targetRow = sheet4.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == formdata.PanelId));
                    if (targetRow != null)
                    {
                        Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Annual Trainer Agreement Valid?"), Value = formdata.IsAnnualTrainerAgreementValid });
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet4, "Including GST?"), Value = formdata.IsInclidingGst });
                        IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet4.Id.Value, new Row[] { updateRow });

                        foreach (var p in formdata.FilesToUpload)
                        {
                            string[] words = p.Split(':');
                            string r = words[0];
                            string q = words[1];
                            string val = formdata.EventId;
                            string name = r.Split(".")[0];
                            string filePath = SheetHelper.testingFile(q, name);
                            Row addedRow = targetRow;
                            Attachment attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(sheet4.Id.Value, addedRow.Id.Value, filePath, "application/msword");
                            if (System.IO.File.Exists(filePath))
                            {
                                SheetHelper.DeleteFile(filePath);
                            }

                        }
                    }
                }
                return Ok(new { Message = " Updated Successfully" });
            }
            catch (Exception ex)
            {
                Log.Error($"Error occured on HonorariumController method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(ex.Message);
            }

        }
    }
}
