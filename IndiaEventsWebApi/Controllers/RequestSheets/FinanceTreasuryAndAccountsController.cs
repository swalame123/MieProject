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

        //[HttpPut("Update")]
        //public void UpdateRow( FinanceAccounts updatedFormData)
        //{
        //    SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
        //    string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
        //    long.TryParse(sheetId, out long parsedSheetId);

        //    Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
        //    Row existingRow = GetRowById(smartsheet, parsedSheetId, updatedFormData.Id);


          
        //   // Row existingRow = smartsheet.SheetResources.RowResources.GetRow(sheetId, rowId, null, null, null, null, null).Data;

           
        //    existingRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet, "pvnumber"), Value = updatedFormData.JVNumber });
        //    existingRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet, "pvdate"), Value = updatedFormData.JVDate });

        //    // Update the existing values if needed
        //    existingRow.Cells.Find(cell => cell.ColumnId == GetColumnIdByName(sheet6, "Expense")).Value = updatedFormData.Expense;
        //    existingRow.Cells.Find(cell => cell.ColumnId == GetColumnIdByName(sheet6, "AmountExcludingTax")).Value = updatedFormData.AmountExcludingTax;
          

        //    // Perform the update
        //    smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId, new Row[] { existingRow });
        //}


        [HttpPut("UpdateData")]
        public IActionResult UpdateData([FromBody] UserRoleMaster updatedData)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:UserRoleMaster").Value;
                long.TryParse(sheetId, out long parsedSheetId);

                // Retrieve the sheet and find the row by id
                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
                Row existingRow = GetRowById(smartsheet, parsedSheetId, updatedData.Email);
                Row updateRow = new Row { Id = existingRow.Id, Cells = new List<Cell>() };

                if (existingRow == null)
                {
                    return NotFound($"Row with id {updatedData.Id} not found.");
                }
                foreach (var cell in existingRow.Cells)
                {
                    if (cell.ColumnId == GetColumnIdByName(sheet, "Email"))
                    {
                        cell.Value = updatedData.Email;
                    }
                    else if (cell.ColumnId == GetColumnIdByName(sheet, "FirstName"))
                    {
                        cell.Value = updatedData.FirstName;
                    }
                    else if (cell.ColumnId == GetColumnIdByName(sheet, "LastName"))
                    {
                        cell.Value = updatedData.LastName;
                    }
                    else if (cell.ColumnId == GetColumnIdByName(sheet, "EmployeeId"))
                    {
                        cell.Value = updatedData.EmployeeId;
                    }
                    else if (cell.ColumnId == GetColumnIdByName(sheet, "RoleId"))
                    {
                        cell.Value = updatedData.RoleId;
                    }
                    else if (cell.ColumnId == GetColumnIdByName(sheet, "RoleName"))
                    {
                        cell.Value = updatedData.RoleName;
                    }
                    else if (cell.ColumnId == GetColumnIdByName(sheet, "CreatedBy"))
                    {
                        cell.Value = updatedData.CreatedBy;
                    }
                    updateRow.Cells.Add(cell);


                }


                smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId, new Row[] { updateRow });
                return Ok(new { Message = "Data updated successfully." });
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


        private Row GetRowById(SmartsheetClient smartsheet, long sheetId, string email)
        {
            Sheet sheet = smartsheet.SheetResources.GetSheet(sheetId, null, null, null, null, null, null, null);

            // Assuming you have a column named "Id"

            Column idColumn = sheet.Columns.FirstOrDefault(col => col.Title == "Email");

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
