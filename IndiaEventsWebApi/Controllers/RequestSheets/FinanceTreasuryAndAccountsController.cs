using IndiaEventsWebApi.Helper;
using IndiaEventsWebApi.Models;
using IndiaEventsWebApi.Models.EventTypeSheets;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Serilog;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Text;

namespace IndiaEventsWebApi.Controllers.RequestSheets
{
    [Route("api/[controller]")]
    [ApiController]
    public class FinanceTreasuryAndAccountsController : ControllerBase
    {

        private readonly string accessToken;
        private readonly IConfiguration configuration;
        private readonly SemaphoreSlim _externalApiSemaphore;
        public FinanceTreasuryAndAccountsController(IConfiguration configuration, SemaphoreSlim externalApiSemaphore)
        {
            this.configuration = configuration;
            this._externalApiSemaphore = externalApiSemaphore;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
        }

        StringBuilder addedBrandsData = new StringBuilder();
        StringBuilder addedInviteesData = new StringBuilder();
        StringBuilder addedHcpData = new StringBuilder();

        [HttpPut("UpdateFinanceAccountPanelSheet")] //no need to change
        public IActionResult UpdateFinanceAccountPanelSheet(FinanceAccountsUpdate updatedFormData)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
                string sheetId1 = configuration.GetSection("SmartsheetSettings:HonorariumPayment").Value;
                long.TryParse(sheetId, out long parsedSheetId);
                long.TryParse(sheetId1, out long parsedSheetId1);
                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
                Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId1, null, null, null, null, null, null, null);

                StringBuilder FinanceAccountsHonorDetails = new StringBuilder();
                int FTNo = 1;
                if (updatedFormData.Status.ToLower() == "approved")
                {
                    foreach (var formdata in updatedFormData.FinanceAccounts)
                    {
                        string rowData = $"{FTNo}. {formdata.HCPName} | MIS Code: {formdata.MISCode}| JV Number: {formdata.JVNumber} | JV Date: {formdata.JVDate.Value.ToShortDateString()}";
                        FinanceAccountsHonorDetails.AppendLine(rowData);
                        FTNo++;
                    }
                }

                string FinanceAccounts = FinanceAccountsHonorDetails.ToString();
                foreach (var f in updatedFormData.FinanceAccounts)
                {
                    Row existingRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == f.Id));
                    Row updateRow = new Row { Id = existingRow.Id, Cells = new List<Cell>() };
                    if (updatedFormData.Status == "Approved")
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "JV Number"), Value = f.JVNumber });
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "JV Date"), Value = f.JVDate });
                        var updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId, new Row[] { updateRow });
                    }
                }
                var targetRow = sheet1.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == updatedFormData.EventId));
                if (targetRow != null)
                {
                    if (updatedFormData.Status == "Approved")
                    {
                        Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Accounts Given Details"), Value = FinanceAccounts });
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "HON-Finance Accounts Approval"), Value = updatedFormData.Status });
                        smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId1, new Row[] { updateRow });
                    }
                    else
                    {
                        Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Rejection Comments"), Value = updatedFormData.Description });
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "HON-Finance Accounts Approval"), Value = updatedFormData.Status });
                        smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId1, new Row[] { updateRow });
                    }
                }
                else
                {
                    return Ok(new { Error = "Invalid Event ID." });
                }
                return Ok(new { Message = "Data Updated successfully." });
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }


        [HttpPut("UpdateFinanceAccountExpenseSheet")]//not in use
        public IActionResult UpdateFinanceAccountExpenseSheet(FinanceAccountsUpdate updatedFormData)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
                string sheetId1 = configuration.GetSection("SmartsheetSettings:EventSettlement").Value;
                long.TryParse(sheetId, out long parsedSheetId);
                long.TryParse(sheetId1, out long parsedSheetId1);

                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
                Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId1, null, null, null, null, null, null, null);


                StringBuilder FinanceAccountsPostDetails = new StringBuilder();
                int FTNo = 1;


                if (updatedFormData.Status.ToLower() == "approved")
                {
                    foreach (var formdata in updatedFormData.FinanceAccounts)
                    {
                        string rowData = $"{FTNo}. {formdata.HCPName} | MIS Code: {formdata.MISCode}| JV Number: {formdata.JVNumber} | JV Date: {formdata.JVDate.Value.ToShortDateString()}";
                        FinanceAccountsPostDetails.AppendLine(rowData);
                        FTNo++;
                    }
                }

                string FinanceAccounts = FinanceAccountsPostDetails.ToString();
                foreach (var f in updatedFormData.FinanceAccounts)
                {
                    Row existingRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == f.Id));
                    Row updateRow = new Row { Id = existingRow.Id, Cells = new List<Cell>() };
                    if (updatedFormData.Status == "Approved")
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "JV Number"), Value = f.JVNumber });
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "JV Date"), Value = f.JVDate });
                        var updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId, new Row[] { updateRow });
                    }
                }
                var targetRow = sheet1.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == updatedFormData.EventId));
                if (targetRow != null)
                {
                    if (updatedFormData.Status == "Approved")
                    {
                        Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Accounts Given Details"), Value = FinanceAccounts });
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventSettlement-Finance Account Approval"), Value = updatedFormData.Status });
                        smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId1, new Row[] { updateRow });
                    }
                    else
                    {
                        Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventSettlement-Finance Account Comments"), Value = updatedFormData.Description });
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventSettlement-Finance Account Approval"), Value = updatedFormData.Status });
                        smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId1, new Row[] { updateRow });
                    }
                }
                else
                {
                    return Ok(new { Error = "Invalid Event ID." });
                }


                return Ok(new { Message = "Data Updated successfully." });

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }

        }

        [HttpPut("UpdateFinanceAccountsInEventSettlement")]
        public async Task<IActionResult> UpdateFinanceAccountsInEventSettlement(FinanceAccountsUpdateIn3Sheets updatedFormData)
        {
            try
            {
                SmartsheetClient smartsheet = await Task.Run(() => SmartSheetBuilder.AccessClient(accessToken, _externalApiSemaphore));


                string EventSettlement = configuration.GetSection("SmartsheetSettings:EventSettlement").Value;
                string EventRequestsExpensesSheet = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
                string EventRequestsHcpRole = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
                string EventRequestInvitees = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;

                Sheet sheet = SheetHelper.GetSheetById(smartsheet, EventSettlement);
                Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == updatedFormData.EventId));
                if (targetRow != null)
                {
                    string FinanceAccounts = "";
                    StringBuilder FinanceTreasuryHonorDetails = new StringBuilder();
                    int FTNo = 1;

                    if (updatedFormData.Status.ToLower() == "approved")
                    {
                        foreach (var formdata in updatedFormData.InviteesSheet)
                        {
                            string rowData = $"{FTNo}. {formdata.HCPName} | MIS Code: {formdata.MISCode}| JV Number: {formdata.JVNumber} | JV Date: {formdata.JVDate.Value.ToShortDateString()}";
                            FinanceTreasuryHonorDetails.AppendLine(rowData);
                            FTNo++;
                        }
                        foreach (var formdata in updatedFormData.PanelSheet)
                        {
                            string rowData = $"{FTNo}. {formdata.HCPName} | MIS Code: {formdata.MISCode}| JV Number: {formdata.JVNumber} | JV Date: {formdata.JVDate.Value.ToShortDateString()}";
                            FinanceTreasuryHonorDetails.AppendLine(rowData);
                            FTNo++;
                        }
                        foreach (var formdata in updatedFormData.ExpenseSheet)
                        {
                            string rowData = $"{FTNo}. {formdata.HCPName} | MIS Code: {formdata.MISCode}| JV Number: {formdata.JVNumber} | JV Date: {formdata.JVDate.Value.ToShortDateString()}";
                            FinanceTreasuryHonorDetails.AppendLine(rowData);
                            FTNo++;
                        }

                        FinanceAccounts = FinanceTreasuryHonorDetails.ToString();

                        Row updateRow = new() { Id = targetRow.Id, Cells = new List<Cell>() };
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance Accounts Given Details"), Value = FinanceAccounts });
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventSettlement-Finance Account Approval"), Value = updatedFormData.Status });
                        smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, new Row[] { updateRow });

                        if (updatedFormData.ExpenseSheet.Count > 0)
                        {
                            Sheet ExpenseSheet = SheetHelper.GetSheetById(smartsheet, EventRequestsExpensesSheet);
                            foreach (var f in updatedFormData.ExpenseSheet)
                            {
                                Row existingRow = ExpenseSheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == f.Id));
                                if (existingRow != null)
                                {
                                    Row updateRow1 = new Row { Id = existingRow.Id, Cells = new List<Cell>() };
                                    updateRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(ExpenseSheet, "JV Number"), Value = f.JVNumber });
                                    updateRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(ExpenseSheet, "JV Date"), Value = f.JVDate });

                                    IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(ExpenseSheet.Id.Value, new Row[] { updateRow1 });

                                }
                            }

                        }
                        if (updatedFormData.PanelSheet.Count > 0)
                        {
                            Sheet PanelSheet = SheetHelper.GetSheetById(smartsheet, EventRequestsHcpRole);
                            foreach (var f in updatedFormData.PanelSheet)
                            {
                                Row existingRow = PanelSheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == f.Id));
                                if (existingRow != null)
                                {
                                    Row updateRow1 = new Row { Id = existingRow.Id, Cells = new List<Cell>() };
                                    updateRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(PanelSheet, "JV Number"), Value = f.JVNumber });
                                    updateRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(PanelSheet, "JV Date"), Value = f.JVDate });

                                    IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(PanelSheet.Id.Value, new Row[] { updateRow1 });

                                }
                            }
                        }
                        if (updatedFormData.InviteesSheet.Count > 0)
                        {
                            Sheet InviteesSheet = SheetHelper.GetSheetById(smartsheet, EventRequestInvitees);
                            foreach (var f in updatedFormData.InviteesSheet)
                            {
                                Row existingRow = InviteesSheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == f.Id));
                                if (existingRow != null)
                                {
                                    Row updateRow1 = new Row { Id = existingRow.Id, Cells = new List<Cell>() };
                                    updateRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(InviteesSheet, "JV Number"), Value = f.JVNumber });
                                    updateRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(InviteesSheet, "JV Date"), Value = f.JVDate });

                                    IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(InviteesSheet.Id.Value, new Row[] { updateRow1 });

                                }
                            }
                        }

                    }
                    else
                    {
                        Row updateRow = new() { Id = targetRow.Id, Cells = new List<Cell>() };
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventSettlement-Finance Account Comments"), Value = updatedFormData.Description });
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventSettlement-Finance Account Approval"), Value = updatedFormData.Status });

                        smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, new Row[] { updateRow });
                    }
                }
                else
                {
                    return Ok(new { Error = "Invalid Event ID." });
                }
                return Ok(new { Message = "Data Updated successfully." });
            }
            catch (Exception ex)
            {
                Log.Error($"Error occured {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(new
                {
                    Message = ex.Message
                });


            }




        }


        [HttpPut("UpdateFinanceTreasuryPanelSheet")]//no need to change
        public IActionResult UpdateFinanceTreasuryPanelSheet(FinanceTreasuryUpdate updatedFormData)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
                string sheetId1 = configuration.GetSection("SmartsheetSettings:HonorariumPayment").Value;
                long.TryParse(sheetId, out long parsedSheetId);
                long.TryParse(sheetId1, out long parsedSheetId1);
                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
                Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId1, null, null, null, null, null, null, null);
                StringBuilder FinanceTreasuryHonorDetails = new StringBuilder();
                int FTNo = 1;
                if (updatedFormData.Status.ToLower() == "approved")
                {
                    foreach (var formdata in updatedFormData.FinanceTreasury)
                    {
                        string rowData = $"{FTNo}. {formdata.HCPName} | MIS Code: {formdata.MISCode}| PV Number: {formdata.PVNumber} | PV Date: {formdata.PVDate.Value.ToShortDateString()} | Bank Reference Number: {formdata.BankReferenceNumber} | Bank Reference Date: {formdata.BankReferenceDate.Value.ToShortDateString()}";
                        FinanceTreasuryHonorDetails.AppendLine(rowData);
                        FTNo++;
                    }
                }

                string FinanceTreasury = FinanceTreasuryHonorDetails.ToString();

                foreach (var f in updatedFormData.FinanceTreasury)
                {
                    if (updatedFormData.Status == "Approved")
                    {
                        Row existingRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == f.Id));
                        if (existingRow != null)
                        {
                            Row updateRow = new Row { Id = existingRow.Id, Cells = new List<Cell>() };
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "PV Number"), Value = f.PVNumber });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "PV Date"), Value = f.PVDate });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Bank Reference Number"), Value = f.BankReferenceNumber });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Bank Reference Date"), Value = f.BankReferenceDate });
                            var updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId, new Row[] { updateRow });
                        }
                        else
                        {
                            return Ok(new { Error = "Invalid Event ID." });
                        }

                    }


                    var targetRow = sheet1.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == updatedFormData.EventId));

                    if (targetRow != null)
                    {
                        if (updatedFormData.Status == "Approved")
                        {
                            Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Treasury Given Details"), Value = FinanceTreasury });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "HON-Finance Treasury Approval"), Value = updatedFormData.Status });
                            smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId1, new Row[] { updateRow });
                        }
                        else
                        {
                            Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Treasury Rejection Comments"), Value = updatedFormData.Description });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "HON-Finance Treasury Approval"), Value = updatedFormData.Status });
                            smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId1, new Row[] { updateRow });
                        }

                    }
                    else
                    {
                        return Ok(new { Error = "Invalid Event ID." });
                    }
                }

                return Ok(new { Message = "Data Updated successfully." });

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }

        }

        [HttpPut("UpdateFinanceTreasuryExpenseSheet")]//not in use
        public IActionResult UpdateFinanceTreasuryExpenseSheet(FinanceTreasuryUpdate updatedFormData)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
                string sheetId1 = configuration.GetSection("SmartsheetSettings:EventSettlement").Value;
                long.TryParse(sheetId, out long parsedSheetId);
                long.TryParse(sheetId1, out long parsedSheetId1);

                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
                Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId1, null, null, null, null, null, null, null);

                string FinanceTreasury = "";
                StringBuilder FinanceTreasuryHonorDetails = new StringBuilder();
                int FTNo = 1;


                if (updatedFormData.Status.ToLower() == "approved")
                {
                    foreach (var formdata in updatedFormData.FinanceTreasury)
                    {
                        string rowData = $"{FTNo}. {formdata.HCPName} | MIS Code: {formdata.MISCode}| PV Number: {formdata.PVNumber} | PV Date: {formdata.PVDate.Value.ToShortDateString()} | Bank Reference Number: {formdata.BankReferenceNumber} | Bank Reference Date: {formdata.BankReferenceDate.Value.ToShortDateString()}";
                        FinanceTreasuryHonorDetails.AppendLine(rowData);
                        FTNo++;


                    }
                    FinanceTreasury = FinanceTreasuryHonorDetails.ToString();
                }


                foreach (var f in updatedFormData.FinanceTreasury)
                {
                    if (updatedFormData.Status == "Approved")
                    {
                        Row existingRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == f.Id));
                        if (existingRow != null)
                        {
                            if (updatedFormData.Status.ToLower() == "approved")
                            {
                                Row updateRow = new Row { Id = existingRow.Id, Cells = new List<Cell>() };
                                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "PV Number"), Value = f.PVNumber });
                                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "PV Date"), Value = f.PVDate });
                                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Bank Reference Number"), Value = f.BankReferenceNumber });
                                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Bank Reference Date"), Value = f.BankReferenceDate });
                                var updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId, new Row[] { updateRow });

                            }
                        }
                        else
                        {
                            return Ok(new { Error = "Invalid Event ID." });
                        }

                    }


                    var targetRow = sheet1.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == updatedFormData.EventId));

                    if (targetRow != null)
                    {
                        if (updatedFormData.Status == "Approved")
                        {
                            Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "Finance Treasury Given Details"), Value = FinanceTreasury });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventSettlement-Finance Treasury Approval"), Value = updatedFormData.Status });
                            smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId1, new Row[] { updateRow });
                        }
                        else
                        {
                            Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventSettlement-Finance Treasury Comments"), Value = updatedFormData.Description });
                            updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet1, "EventSettlement-Finance Treasury Approval"), Value = updatedFormData.Status });
                            smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId1, new Row[] { updateRow });
                        }

                    }
                    else
                    {
                        return Ok(new { Error = "Invalid Event ID." });
                    }
                }
                return Ok(new { Message = "Data Updated successfully." });

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }

        }

        [HttpPut("UpdateFinanceTreasuryInEventSettlement")]
        public async Task<IActionResult> UpdateFinanceTreasuryInEventSettlement(FinanceTreasuryUpdateIn3Sheets updatedFormData)
        {
            try
            {
                SmartsheetClient smartsheet = await Task.Run(() => SmartSheetBuilder.AccessClient(accessToken, _externalApiSemaphore));


                string EventSettlement = configuration.GetSection("SmartsheetSettings:EventSettlement").Value;
                string EventRequestsExpensesSheet = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
                string EventRequestsHcpRole = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
                string EventRequestInvitees = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;

                Sheet sheet = SheetHelper.GetSheetById(smartsheet, EventSettlement);
                Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == updatedFormData.EventId));
                if (targetRow != null)
                {
                    string FinanceTreasury = "";
                    StringBuilder FinanceTreasuryHonorDetails = new StringBuilder();
                    int FTNo = 1;

                    if (updatedFormData.Status.ToLower() == "approved")
                    {
                        foreach (var formdata in updatedFormData.InviteesSheet)
                        {

                            string rowData = $"{FTNo}. {formdata.HCPName} | MIS Code: {formdata.MISCode}| PV Number: {formdata.PVNumber} | PV Date: {formdata.PVDate.Value.ToShortDateString()} | Bank Reference Number: {formdata.BankReferenceNumber} | Bank Reference Date: {formdata.BankReferenceDate.Value.ToShortDateString()}";
                            FinanceTreasuryHonorDetails.AppendLine(rowData);
                            FTNo++;
                        }
                        foreach (var formdata in updatedFormData.PanelSheet)
                        {

                            string rowData = $"{FTNo}. {formdata.HCPName} | MIS Code: {formdata.MISCode}| PV Number: {formdata.PanelDataInFinance.Travel.PVNumber} | PV Date: {formdata.PanelDataInFinance.Travel.PVDate.Value.ToShortDateString()} | Bank Reference Number: {formdata.PanelDataInFinance.Travel.BankReferenceNumber} | Bank Reference Date: {formdata.PanelDataInFinance.Travel.BankReferenceDate.Value.ToShortDateString()}";
                            FinanceTreasuryHonorDetails.AppendLine(rowData);
                            FTNo++;
                        }
                        foreach (var formdata in updatedFormData.ExpenseSheet)
                        {

                            string rowData = $"{FTNo}. {formdata.HCPName} | MIS Code: {formdata.MISCode}| PV Number: {formdata.PVNumber} | PV Date: {formdata.PVDate.Value.ToShortDateString()} | Bank Reference Number: {formdata.BankReferenceNumber} | Bank Reference Date: {formdata.BankReferenceDate.Value.ToShortDateString()}";
                            FinanceTreasuryHonorDetails.AppendLine(rowData);
                            FTNo++;
                        }

                        FinanceTreasury = FinanceTreasuryHonorDetails.ToString();

                        Row updateRow = new() { Id = targetRow.Id, Cells = new List<Cell>() };
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance Treasury Given Details"), Value = FinanceTreasury });
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventSettlement-Finance Treasury Approval"), Value = updatedFormData.Status });
                        smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, new Row[] { updateRow });

                        if (updatedFormData.ExpenseSheet.Count > 0)
                        {
                            Sheet ExpenseSheet = SheetHelper.GetSheetById(smartsheet, EventRequestsExpensesSheet);
                            foreach (var f in updatedFormData.ExpenseSheet)
                            {
                                Row existingRow = ExpenseSheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == f.Id));
                                if (existingRow != null)
                                {
                                    Row updateRow1 = new Row { Id = existingRow.Id, Cells = new List<Cell>() };
                                    updateRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(ExpenseSheet, "PV Number"), Value = f.PVNumber });
                                    updateRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(ExpenseSheet, "PV Date"), Value = f.PVDate });
                                    updateRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(ExpenseSheet, "Bank Reference Number"), Value = f.BankReferenceNumber });
                                    updateRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(ExpenseSheet, "Bank Reference Date"), Value = f.BankReferenceDate });

                                    IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(ExpenseSheet.Id.Value, new Row[] { updateRow1 });

                                }
                            }

                        }
                        if (updatedFormData.PanelSheet.Count > 0)
                        {
                            Sheet PanelSheet = SheetHelper.GetSheetById(smartsheet, EventRequestsHcpRole);
                            foreach (var f in updatedFormData.PanelSheet)
                            {
                                Row existingRow = PanelSheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == f.Id));
                                if (existingRow != null)
                                {
                                    Row updateRow1 = new Row { Id = existingRow.Id, Cells = new List<Cell>() };
                                    updateRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(PanelSheet, "PV Number"), Value = f.PanelDataInFinance.Travel.PVNumber });
                                    updateRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(PanelSheet, "PV Date"), Value = f.PanelDataInFinance.Travel.PVDate });
                                    updateRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(PanelSheet, "Bank Reference Number"), Value = f.PanelDataInFinance.Travel.BankReferenceNumber });
                                    updateRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(PanelSheet, "Bank Reference Date"), Value = f.PanelDataInFinance.Travel.BankReferenceDate });

                                    IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(PanelSheet.Id.Value, new Row[] { updateRow1 });

                                }
                            }
                        }
                        if (updatedFormData.InviteesSheet.Count > 0)
                        {
                            Sheet InviteesSheet = SheetHelper.GetSheetById(smartsheet, EventRequestInvitees);
                            foreach (var f in updatedFormData.InviteesSheet)
                            {
                                Row existingRow = InviteesSheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == f.Id));
                                if (existingRow != null)
                                {
                                    Row updateRow1 = new Row { Id = existingRow.Id, Cells = new List<Cell>() };
                                    updateRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(InviteesSheet, "PV Number"), Value = f.PVNumber });
                                    updateRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(InviteesSheet, "PV Date"), Value = f.PVDate });
                                    updateRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(InviteesSheet, "Bank Reference Number"), Value = f.BankReferenceNumber });
                                    updateRow1.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(InviteesSheet, "Bank Reference Date"), Value = f.BankReferenceDate });

                                    IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(InviteesSheet.Id.Value, new Row[] { updateRow1 });

                                }
                            }
                        }

                    }
                    else
                    {
                        Row updateRow = new() { Id = targetRow.Id, Cells = new List<Cell>() };
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventSettlement-Finance Treasury Comments"), Value = updatedFormData.Description });
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventSettlement-Finance Treasury Approval"), Value = updatedFormData.Status });

                        smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, new Row[] { updateRow });
                    }
                }
                else
                {
                    return Ok(new { Error = "Invalid Event ID." });
                }
                return Ok(new { Message = "Data Updated successfully." });
            }
            catch (Exception ex)
            {
                Log.Error($"Error occured {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(new
                {
                    Message = ex.Message
                });


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


        private Row GetRowByIdHCP(SmartsheetClient smartsheet, long sheetId, string email)
        {
            Sheet sheet = smartsheet.SheetResources.GetSheet(sheetId, null, null, null, null, null, null, null);



            Column idColumn = sheet.Columns.FirstOrDefault(col => col.Title == "Panelist ID");

            if (idColumn != null)
            {
                foreach (var row in sheet.Rows)
                {
                    var cell = row.Cells.FirstOrDefault(c => c.ColumnId == idColumn.Id && c.Value.ToString() == email);

                    if (cell != null)
                    {
                        return row;
                    }
                }
            }

            return null;
        }


        private Row GetRowByIdEXP(SmartsheetClient smartsheet, long sheetId, string email)
        {
            Sheet sheet = smartsheet.SheetResources.GetSheet(sheetId, null, null, null, null, null, null, null);



            Column idColumn = sheet.Columns.FirstOrDefault(col => col.Title == "Expenses ID");

            if (idColumn != null)
            {
                foreach (var row in sheet.Rows)
                {
                    var cell = row.Cells.FirstOrDefault(c => c.ColumnId == idColumn.Id && c.Value.ToString() == email);

                    if (cell != null)
                    {
                        return row;
                    }
                }
            }

            return null;
        }


    }
}




































