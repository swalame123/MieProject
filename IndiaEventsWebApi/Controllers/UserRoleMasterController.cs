//using Microsoft.AspNetCore.Http;
//using Microsoft.AspNetCore.Mvc;
//using Microsoft.AspNetCore.Mvc.Filters;
//using IndiaEventsWebApi.Models;
//using NPOI.OpenXml4Net.OPC;
//using Smartsheet.Api;
//using Smartsheet.Api.Models;
//using Smartsheet.Api.OAuth;
//using System.Runtime;
//using System.Runtime.InteropServices;

//namespace IndiaEventsWebApi.Controllers
//{
//    [Route("api/[controller]")]
//    [ApiController]
//    public class UserRoleMasterController : ControllerBase
//    {

//        private readonly string accessToken;
//        private readonly IConfiguration configuration;

//        public UserRoleMasterController(IConfiguration configuration)
//        {
//            this.configuration = configuration;
//            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

//        }
//        [HttpGet("GetEventData")]
//        public IActionResult GetEventData()
//        {
//            try
//            {
//                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
//                string sheetId = configuration.GetSection("SmartsheetSettings:UserRoleMaster").Value;
//                long.TryParse(sheetId, out long parsedSheetId);
//                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
//                List<Dictionary<string, object>> sheetData = new List<Dictionary<string, object>>();
//                List<string> columnNames = new List<string>();
//                foreach (Column column in sheet.Columns)
//                {
//                    columnNames.Add(column.Title);
//                }
//                foreach (Row row in sheet.Rows)
//                {
//                    Dictionary<string, object> rowData = new Dictionary<string, object>();
//                    for (int i = 0; i < row.Cells.Count && i < columnNames.Count; i++)
//                    {
//                        rowData[columnNames[i]] = row.Cells[i].Value;

//                    }
//                    sheetData.Add(rowData);
//                }
//                return Ok(sheetData);
//            }
//            catch (Exception ex)
//            {
//                return BadRequest(ex.Message);
//            }
//        }
//        private bool EmailExists(Sheet sheet, string email)
//        {
//            long emailColumnId = GetColumnIdByName(sheet, "Email");

//            return sheet.Rows.Any(row =>
//                row.Cells.Any(cell => cell.ColumnId == emailColumnId && cell.Value?.ToString() == email));
//        }
//        [HttpPost("AddData")]
//        public IActionResult AddData([FromBody] UserRoleMaster formData)
//        {
//            try
//            {
//                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                
//                string sheetId = configuration.GetSection("SmartsheetSettings:UserRoleMaster").Value;

//                long.TryParse(sheetId, out long parsedSheetId);
//                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
//                if (EmailExists(sheet, formData.Email))
//                {
//                    return BadRequest("Email already exists in the sheet.");
//                }
//                var newRow = new Row();
//                newRow.Cells = new List<Cell>();
//                newRow.Cells.Add(new Cell
//                {
//                    ColumnId = GetColumnIdByName(sheet, "Email"),
//                    Value = formData.Email
//                });
//                newRow.Cells.Add(new Cell
//                {
//                    ColumnId = GetColumnIdByName(sheet, "FirstName"),
//                    Value = formData.FirstName
//                });
//                newRow.Cells.Add(new Cell
//                {
//                    ColumnId = GetColumnIdByName(sheet, "LastName"),
//                    Value = formData.LastName
//                });
//                newRow.Cells.Add(new Cell
//                {
//                    ColumnId = GetColumnIdByName(sheet, "EmployeeId"),
//                    Value = formData.EmployeeId
//                });
//                newRow.Cells.Add(new Cell
//                {
//                    ColumnId = GetColumnIdByName(sheet, "RoleId"),
//                    Value = formData.RoleId
//                });
//                newRow.Cells.Add(new Cell
//                {
//                    ColumnId = GetColumnIdByName(sheet, "RoleName"),
//                    Value = formData.RoleName
//                });
//                newRow.Cells.Add(new Cell
//                {
//                    ColumnId = GetColumnIdByName(sheet, "CreatedBy"),
//                    Value = formData.CreatedBy
//                });
                
//                smartsheet.SheetResources.RowResources.AddRows(parsedSheetId, new Row[] { newRow });
//                //return Ok("Data added successfully.");
//                return Ok(new
//                { Message = "Data added successfully." });

//            }
//            catch (Exception ex)
//            {
//                return BadRequest(ex.Message);
//            }
//        }
//        private long GetColumnIdByName(Sheet sheet, string columnname)
//        {
//            foreach (var column in sheet.Columns)
//            {
//                if (column.Title == columnname)
//                {
//                    return column.Id.Value;
//                }
//            }
//            return 0;
//        }



//        [HttpPut("UpdateData")]
//        public IActionResult UpdateData([FromBody] UserRoleMaster updatedData)
//        {
//            try
//            {
//                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
//                string sheetId = configuration.GetSection("SmartsheetSettings:UserRoleMaster").Value;
//                long.TryParse(sheetId, out long parsedSheetId);

//                // Retrieve the sheet and find the row by id
//                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
//                Row existingRow = GetRowById(smartsheet, parsedSheetId, updatedData.Email);
//                Row updateRow = new Row { Id = existingRow.Id, Cells = new List<Cell>() };

//                if (existingRow == null)
//                {
//                    return NotFound($"Row with id {updatedData.Id} not found.");
//                }
//                foreach (var cell in existingRow.Cells)
//                {
//                    if (cell.ColumnId == GetColumnIdByName(sheet, "Email"))
//                    {
//                        cell.Value = updatedData.Email;
//                    }
//                    else if (cell.ColumnId == GetColumnIdByName(sheet, "FirstName"))
//                    {
//                        cell.Value = updatedData.FirstName;
//                    }
//                    else if (cell.ColumnId == GetColumnIdByName(sheet, "LastName"))
//                    {
//                        cell.Value = updatedData.LastName;
//                    }
//                    else if (cell.ColumnId == GetColumnIdByName(sheet, "EmployeeId"))
//                    {
//                        cell.Value = updatedData.EmployeeId;
//                    }
//                    else if (cell.ColumnId == GetColumnIdByName(sheet, "RoleId"))
//                    {
//                        cell.Value = updatedData.RoleId;
//                    }
//                    else if (cell.ColumnId == GetColumnIdByName(sheet, "RoleName"))
//                    {
//                        cell.Value = updatedData.RoleName;
//                    }
//                    else if (cell.ColumnId == GetColumnIdByName(sheet, "CreatedBy"))
//                    {
//                        cell.Value = updatedData.CreatedBy;
//                    }
//                    updateRow.Cells.Add(cell);


//                }

              
//                smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId, new Row[] { updateRow });
//                return Ok(new { Message = "Data updated successfully." });
//            }
//            catch (Exception ex)
//            {
//                return BadRequest(ex.Message);
//            }
//        }



//        [HttpDelete("DeleteData/{email}")]
//        public IActionResult DeleteData(string email)
//        {
//            try
//            {
//                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
//                string sheetId = configuration.GetSection("SmartsheetSettings:UserRoleMaster").Value;
//                long.TryParse(sheetId, out long parsedSheetId);

//                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
//                Row existingRow = GetRowById(smartsheet, parsedSheetId, email);


//                if (existingRow == null)
//                {
//                    return NotFound($"Row with id {email} not found.");
//                }
//                var Id = (long)existingRow.Id;
//                smartsheet.SheetResources.RowResources.DeleteRows(parsedSheetId, new long[] { Id }, true);

//                return Ok(new { Message = "Data updated successfully." });
//            }
//            catch (Exception ex)
//            {
//                return BadRequest(ex.Message);
//            }
//        }

//        //[HttpDelete("DeleteData/{email}")]
//        //public IActionResult DeleteData([FromBody] UserRoleMaster updatedData)
//        //{
//        //    try
//        //    {
//        //        SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
//        //        string sheetId = configuration.GetSection("SmartsheetSettings:UserRoleMaster").Value;
//        //        long.TryParse(sheetId, out long parsedSheetId);

//        //        Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
//        //        Row existingRow = GetRowById(smartsheet, parsedSheetId, updatedData.Email);


//        //        if (existingRow == null)
//        //        {
//        //            return NotFound($"Row with id {updatedData.Id} not found.");
//        //        }
//        //        var Id = (long)existingRow.Id;
//        //        smartsheet.SheetResources.RowResources.DeleteRows(parsedSheetId, new long[] { Id }, true);

//        //        return Ok(new { Message = "Data updated successfully." });
//        //    }
//        //    catch (Exception ex)
//        //    {
//        //        return BadRequest(ex.Message);
//        //    }
//        //}



//        private Row GetRowById(SmartsheetClient smartsheet, long sheetId, string email)
//        {
//            Sheet sheet = smartsheet.SheetResources.GetSheet(sheetId, null, null, null, null, null, null, null);

//            // Assuming you have a column named "Id"

//            Column idColumn = sheet.Columns.FirstOrDefault(col => col.Title == "Email");

//            if (idColumn != null)
//            {
//                foreach (var row in sheet.Rows)
//                {
//                    var cell = row.Cells.FirstOrDefault(c => c.ColumnId == idColumn.Id && c.Value.ToString() == email);

//                    if (cell != null)
//                    {
//                        return row;
//                    }
//                }
//            }

//            return null;
//        }
        
//        //[HttpDelete("DeleteData/{email}")]
//        //public IActionResult DeleteData(string email)
//        //{
//        //    try
//        //    {
//        //        SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
//        //        string sheetId = configuration.GetSection("SmartsheetSettings:UserRoleMaster").Value;
//        //        long.TryParse(sheetId, out long parsedSheetId);

//        //        // Get the row to be deleted based on email
//        //        Row existingRow = GetRowByEmail(smartsheet, parsedSheetId, email);

//        //        if (existingRow == null)
//        //        {
//        //            return NotFound($"Row with email {email} not found.");
//        //        }

//        //        // Delete the row
//        //        smartsheet.SheetResources.RowResources.DeleteRow(parsedSheetId, existingRow.Id);

//        //        return Ok(new { Message = "Data deleted successfully." });
//        //    }
//        //    catch (Exception ex)
//        //    {
//        //        return BadRequest(ex.Message);
//        //    }
//        //}
//    }
//}




