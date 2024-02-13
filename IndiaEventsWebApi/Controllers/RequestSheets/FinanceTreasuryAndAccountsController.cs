using IndiaEventsWebApi.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Smartsheet.Api;
using Smartsheet.Api.Models;
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



        //        var targetRow = sheet1.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == eventId));

        //                if (targetRow != null)
        //                {
        //                    long honorariumSubmittedColumnId = GetColumnIdByName(sheet1, "Honorarium Submitted?");
        //        var cellToUpdateB = new Cell
        //        {
        //            ColumnId = honorariumSubmittedColumnId,
        //            Value = "Yes"
        //        };
        //        Row updateRow = new Row { Id = targetRow.Id, Cells = new Cell[] { cellToUpdateB } };
        //        var cellToUpdate = targetRow.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
        //                    if (cellToUpdate != null)
        //                    {
        //                        cellToUpdate.Value = "Yes";
        //                    }

        //    smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId1, new Row[] { updateRow
        //});


        //}

        [HttpPut("UpdateFinanceAccountPanelSheet")]
        public IActionResult UpdateFinanceAccountPanelSheet(FinanceAccounts[] updatedFormData)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
                long.TryParse(sheetId, out long parsedSheetId);

                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);

                foreach (var f in updatedFormData)
                {

                    Row existingRow = GetRowByIdHCP(smartsheet, parsedSheetId, f.Id);
                    Row updateRow = new Row { Id = existingRow.Id, Cells = new List<Cell>() };


                    // Row existingRow = smartsheet.SheetResources.RowResources.GetRow(sheetId, rowId, null, null, null, null, null).Data;


                    updateRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet, "JV Number"), Value = f.JVNumber });
                    updateRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet, "JV Date"), Value = f.JVDate });
                    

                    
                    smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId, new Row[] { updateRow });
                }

                return Ok(new { Message = "Data Updated successfully." });

            }
            catch (Exception ex)
            {
               return  BadRequest(ex.Message);
            }
            
        }


        [HttpPut("UpdateFinanceAccountExpenseSheet")]
        public IActionResult UpdateFinanceAccountExpenseSheet(FinanceAccounts[] updatedFormData)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
                long.TryParse(sheetId, out long parsedSheetId);

                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);

                foreach (var f in updatedFormData)
                {

                    Row existingRow = GetRowByIdEXP(smartsheet, parsedSheetId, f.Id);
                    Row updateRow = new Row { Id = existingRow.Id, Cells = new List<Cell>() };


                    // Row existingRow = smartsheet.SheetResources.RowResources.GetRow(sheetId, rowId, null, null, null, null, null).Data;


                    updateRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet, "JV Number"), Value = f.JVNumber });
                    updateRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet, "JV Date"), Value = f.JVDate });



                    smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId, new Row[] { updateRow });
                }

                return Ok(new { Message = "Data Updated successfully." });

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }

        }




        [HttpPut("UpdateFinanceTreasuryPanelSheet")]
        public IActionResult UpdateFinanceTreasuryPanelSheet(FinanceTreasury[] updatedFormData)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
                long.TryParse(sheetId, out long parsedSheetId);

                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);

                foreach (var f in updatedFormData)
                {

                    Row existingRow = GetRowByIdHCP(smartsheet, parsedSheetId, f.Id);
                    Row updateRow = new Row { Id = existingRow.Id, Cells = new List<Cell>() };


                    // Row existingRow = smartsheet.SheetResources.RowResources.GetRow(sheetId, rowId, null, null, null, null, null).Data;


                    updateRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet, "PV Number"), Value = f.PVNumber });
                    updateRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet, "PV Date"), Value = f.PVDate });
                    updateRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet, "Bank Reference Number"), Value = f.BankReferenceNumber });
                    updateRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet, "Bank Reference Date"), Value = f.BankReferenceDate });



                    smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId, new Row[] { updateRow });
                }

                return Ok(new { Message = "Data Updated successfully." });

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }

        }
        [HttpPut("UpdateFinanceTreasuryExpenseSheet")]
        public IActionResult UpdateFinanceTreasuryExpenseSheet(FinanceTreasury[] updatedFormData)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
                long.TryParse(sheetId, out long parsedSheetId);

                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);

                foreach (var f in updatedFormData)
                {

                    Row existingRow = GetRowByIdEXP(smartsheet, parsedSheetId, f.Id);
                    Row updateRow = new Row { Id = existingRow.Id, Cells = new List<Cell>() };


                    // Row existingRow = smartsheet.SheetResources.RowResources.GetRow(sheetId, rowId, null, null, null, null, null).Data;


                    updateRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet, "PV Number"), Value = f.PVNumber });
                    updateRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet, "PV Date"), Value = f.PVDate });
                    updateRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet, "Bank Reference Number"), Value = f.BankReferenceNumber });
                    updateRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet, "Bank Reference Date"), Value = f.BankReferenceDate });



                    smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId, new Row[] { updateRow });
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




































