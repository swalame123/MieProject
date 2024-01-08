using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;
using MieProject.Models;
using NPOI.OpenXml4Net.OPC;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using Smartsheet.Api.OAuth;
using System.Runtime.InteropServices;

namespace MieProject.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class UserRoleMasterController : ControllerBase
    {

        private readonly string accessToken;
        private readonly IConfiguration configuration;

        public UserRoleMasterController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

        }
        [HttpGet("GetEventData")]
        public IActionResult GetEventData()
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:UserRoleMaster").Value;
                long.TryParse(sheetId, out long parsedSheetId);
                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
                List<Dictionary<string, object>> sheetData = new List<Dictionary<string, object>>();
                List<string> columnNames = new List<string>();
                foreach (Column column in sheet.Columns)
                {
                    columnNames.Add(column.Title);
                }
                foreach (Row row in sheet.Rows)
                {
                    Dictionary<string, object> rowData = new Dictionary<string, object>();
                    for (int i = 0; i < row.Cells.Count && i < columnNames.Count; i++)
                    {
                        rowData[columnNames[i]] = row.Cells[i].Value;

                    }
                    sheetData.Add(rowData);
                }
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        private bool EmailExists(Sheet sheet, string email)
        {
            long emailColumnId = GetColumnIdByName(sheet, "Email");

            return sheet.Rows.Any(row =>
                row.Cells.Any(cell => cell.ColumnId == emailColumnId && cell.Value?.ToString() == email));
        }
        [HttpPost("AddData")]
        public IActionResult AddData([FromBody] UserRoleMaster formData)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                
                string sheetId = configuration.GetSection("SmartsheetSettings:UserRoleMaster").Value;

                long.TryParse(sheetId, out long parsedSheetId);
                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
                if (EmailExists(sheet, formData.Email))
                {
                    return BadRequest("Email already exists in the sheet.");
                }
                var newRow = new Row();
                newRow.Cells = new List<Cell>();
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Email"),
                    Value = formData.Email
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "FirstName"),
                    Value = formData.FirstName
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "LastName"),
                    Value = formData.LastName
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "EmployeeId"),
                    Value = formData.EmployeeId
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "RoleId"),
                    Value = formData.RoleId
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "RoleName"),
                    Value = formData.RoleName
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "CreatedBy"),
                    Value = formData.UserName
                });
                
                smartsheet.SheetResources.RowResources.AddRows(parsedSheetId, new Row[] { newRow });
                //return Ok("Data added successfully.");
                return Ok(new
                { Message = "Data added successfully." });

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

       
        [HttpPut("UpdateData")]
        public IActionResult UpdateData( [FromBody] UserRoleMaster updatedData)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:UserRoleMaster").Value;
                long.TryParse(sheetId, out long parsedSheetId);

                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);

                //Row existingRow = GetRowById(sheet, updatedData.Id);
                Row existingRow = GetRowById(smartsheet, parsedSheetId, updatedData.Id);

                if (existingRow == null)
                {
                    return BadRequest("Row with specified ID not found.");
                }

                // Update the existing row with the new data
                UpdateCellValue(existingRow, sheet, "FirstName", updatedData.FirstName);
                UpdateCellValue(existingRow, sheet, "LastName", updatedData.LastName);
                UpdateCellValue(existingRow, sheet, "EmployeeId", updatedData.EmployeeId);
                UpdateCellValue(existingRow, sheet, "RoleId", updatedData.RoleId);
                UpdateCellValue(existingRow, sheet, "RoleName", updatedData.RoleName);
                UpdateCellValue(existingRow, sheet, "CreatedBy", updatedData.UserName);

                smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId, new Row[] { existingRow });

                return Ok(new { Message = "Data updated successfully." });
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        private void UpdateCellValue(Row row, Sheet sheet, string columnName, object value)
        {
            long columnId = GetColumnIdByName(sheet, columnName);
            var cell = row.Cells.FirstOrDefault(c => c.ColumnId == columnId);
            if (cell != null)
            {
                cell.Value = value;
            }
        }

        //private Row GetRowById(Sheet sheet, int id)
        //{
        //    foreach (var row in sheet.Rows)
        //    {
        //        var idCell = row.Cells.FirstOrDefault(cell => cell.ColumnId == GetColumnIdByName(sheet, "ID"));
        //        if (idCell != null && long.TryParse(idCell.Value.ToString(), out var cellId) && cellId == id)
        //        {
        //            return row;
        //        }
        //    }
        //    return null;
        //}


        //[HttpPut("UpdateData")]
        //public IActionResult UpdateData( [FromBody] UserRoleMaster updatedData)
        //{
        //    try
        //    {
        //        SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
        //        string sheetId = configuration.GetSection("SmartsheetSettings:UserRoleMaster").Value;
        //        long.TryParse(sheetId, out long parsedSheetId);

        //        // Retrieve the sheet and find the row by id
        //        Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
        //        Row existingRow = GetRowById(smartsheet, parsedSheetId, updatedData.Id);

        //        if (existingRow == null)
        //        {
        //            return NotFound($"Row with id {updatedData.Id} not found.");
        //        }
        //        foreach (var cell in rowToEdit.Cells)
        //        {
        //            if (cell.ColumnId == GetColumnIdByName(sheet, "Full Name"))
        //            {
        //                cell.Value = model.FullName;
        //            }
        //            else if (cell.ColumnId == GetColumnIdByName(sheet, "Email"))
        //            {
        //                cell.Value = model.Email;
        //            }
        //            else if (cell.ColumnId == GetColumnIdByName(sheet, "Phone"))
        //            {
        //                cell.Value = model.Phone;
        //            }
        //            updateRow.Cells.Add(cell);


        //        }

        //        // Update the existing row with new data
        //        //existingRow.Cells.First(c => c.ColumnId == GetColumnIdByName(sheet, "Email")).Value = updatedData.Email;
        //        //existingRow.Cells.First(c => c.ColumnId == GetColumnIdByName(sheet, "FirstName")).Value = updatedData.FirstName;
        //        //existingRow.Cells.First(c => c.ColumnId == GetColumnIdByName(sheet, "LastName")).Value = updatedData.LastName;
        //        //existingRow.Cells.First(c => c.ColumnId == GetColumnIdByName(sheet, "EmployeeId")).Value = updatedData.EmployeeId;
        //        //existingRow.Cells.First(c => c.ColumnId == GetColumnIdByName(sheet, "RoleId")).Value = updatedData.RoleId;
        //        //existingRow.Cells.First(c => c.ColumnId == GetColumnIdByName(sheet, "RoleName")).Value = updatedData.RoleName;
        //        //existingRow.Cells.First(c => c.ColumnId == GetColumnIdByName(sheet, "CreatedBy")).Value = updatedData.UserName;

        //        //smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId, new Row[] { new Row { Id = existingRow.Id,Cells= existingRow.Cells } });
        //        smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId, new Row[] { existingRow });

        //        return Ok(new { Message = "Data updated successfully." });
        //    }
        //    catch (Exception ex)
        //    {
        //        return BadRequest(ex.Message);
        //    }
        //}

        ////private Row GetRowById(SmartsheetClient smartsheet, long sheetId, int id)
        ////{
        ////    Sheet sheet = smartsheet.SheetResources.GetSheet(sheetId, null, null, null, null, null, null, null);
        ////    return sheet.Rows.FirstOrDefault(row => row.Id == id);
        ////}

        private Row GetRowById(SmartsheetClient smartsheet, long sheetId, int id)
        {
            Sheet sheet = smartsheet.SheetResources.GetSheet(sheetId, null, null, null, null, null, null, null);

            // Assuming you have a column named "Id"

            Column idColumn = sheet.Columns.FirstOrDefault(col => col.Title == "id");

            if (idColumn != null)
            {
                foreach (var row in sheet.Rows)
                {
                    var cell = row.Cells.FirstOrDefault(c => c.ColumnId == idColumn.Id && c.Value.ToString() == id.ToString());

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


