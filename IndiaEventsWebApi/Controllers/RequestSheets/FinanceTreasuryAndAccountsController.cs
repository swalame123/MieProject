using IndiaEventsWebApi.Helper;
using IndiaEventsWebApi.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Text;
using static Org.BouncyCastle.Bcpg.Attr.ImageAttrib;

namespace IndiaEventsWebApi.Controllers.RequestSheets
{
    [Route("api/[controller]")]
    [ApiController]
    public class FinanceTreasuryAndAccountsController : ControllerBase
    {

        private readonly string accessToken;
        private readonly IConfiguration configuration;

        public FinanceTreasuryAndAccountsController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
        }

        StringBuilder addedBrandsData = new StringBuilder();
        StringBuilder addedInviteesData = new StringBuilder();
        StringBuilder addedHcpData = new StringBuilder();

        [HttpPut("UpdateFinanceAccountPanelSheet")]
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


        [HttpPut("UpdateFinanceAccountExpenseSheet")]
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




        [HttpPut("UpdateFinanceTreasuryPanelSheet")]
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

        [HttpPut("UpdateFinanceTreasuryExpenseSheet")]
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




































