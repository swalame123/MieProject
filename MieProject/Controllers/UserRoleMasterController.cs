using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using MieProject.Models;
using NPOI.OpenXml4Net.OPC;
using Smartsheet.Api;
using Smartsheet.Api.Models;

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
        [HttpPost("AddData")]
        public IActionResult AddData([FromBody] UserRoleMaster formData)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                // Retrieve sheet ID from configuration
                string sheetId = configuration.GetSection("SmartsheetSettings:UserRoleMaster").Value;

                // Convert sheet ID to long if needed
                long.TryParse(sheetId, out long parsedSheetId);
                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);

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
                    ColumnId = GetColumnIdByName(sheet, "CreatedBy"),
                    Value = formData.UserName
                });
                smartsheet.SheetResources.RowResources.AddRows(parsedSheetId, new Row[] { newRow });
                return Ok("Data added successfully.");

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

    }
}



    //[HttpPost("AddData")]

//public IActionResult AddData([FromBody] UserRoleMaster formData)
//{
//    try
//    {

//        SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

//        string sheet1Id = configuration.GetSection("SmartsheetSettings:sheetId1").Value;
//        string sheet2Id = configuration.GetSection("SmartsheetSettings:sheetId2").Value;
//        string roleId = configuration.GetSection("SmartsheetSettings:RoleMaster").Value;
//        string userRoleMasterSheetId = configuration.GetSection("SmartsheetSettings:UserRoleMasterSheetId").Value;
//        long.TryParse(userRoleMasterSheetId, out long userRoleMasterSheetParsedId);
//        Sheet userRoleMasterSheet = smartsheet.SheetResources.GetSheet(userRoleMasterSheetParsedId, null, null, null, null, null, null, null);
//        string employeeId = LookupEmployeeId(formData.Email, sheet1Id, sheet2Id);                
//        var newRow = new Row
//        {
//            Cells = new List<Cell>
//            {
//                new Cell
//                {
//                    ColumnId = GetColumnIdByName(userRoleMasterSheet, "EmployeeId"),
//                    Value = employeeId
//                },
//                new Cell
//                {
//                    ColumnId = GetColumnIdByName(userRoleMasterSheet, "Username"),
//                    Value = formData.Email
//                },
//                new Cell
//                {
//                    ColumnId = GetColumnIdByName(userRoleMasterSheet, "RoleId"),
//                    Value = roleId
//                },

//            }
//        };

//        smartsheet.SheetResources.RowResources.AddRows(userRoleMasterSheetParsedId, new Row[] { newRow });

//        return Ok("User Role created successfully");
//    }
//    catch (Exception ex)
//    {
//        return BadRequest(ex.Message);
//    }

//}
//private string LookupEmployeeId(string email, string sheet1Id, string sheet2Id)
//{
//    SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

//    string employeeId = LookupInSheet(email, "Email", "EmployeeId", sheet1Id, smartsheet);

//    if (string.IsNullOrEmpty(employeeId))
//    {
//        employeeId = LookupInSheet(email, "Email", "EmployeeId", sheet2Id, smartsheet);
//    }

//    return employeeId;
//}
//private string LookupRoleId(string roleName)
//{
//    SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

//    string roleId = LookupInSheet(roleName, "Role", "RoleId", "RoleSheetId", smartsheet);

//    return roleId;
//}
//private string LookupInSheet(string searchValue, string searchColumn, string resultColumn, string sheetId, SmartsheetClient smartsheet)
//{
//    long.TryParse(sheetId, out long parsedSheetId);
//    Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);

//    // Find the column IDs based on the provided column names
//    long searchColumnId = GetColumnIdByName(sheet, searchColumn);
//    long resultColumnId = GetColumnIdByName(sheet, resultColumn);

//    // Search for the row with the matching value in the search column
//    var matchingRow = sheet.Rows.FirstOrDefault(row =>
//        row.Cells.Any(cell => cell.ColumnId == searchColumnId && cell.Value?.ToString() == searchValue));

//    // If a matching row is found, return the value from the result column
//    if (matchingRow != null)
//    {
//        var resultCell = matchingRow.Cells.FirstOrDefault(cell => cell.ColumnId == resultColumnId);
//        return resultCell?.Value?.ToString();
//    }

//    // If no matching row is found, return null or an empty string based on your preference
//    return null;
//}



//[HttpPost("AddData")]
//public IActionResult AddData([FromBody] FormData formData)
//{
//    try
//    {
//        SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
//        // Retrieve sheet ID from configuration
//        string sheetId = configuration.GetSection("SmartsheetSettings:SheetId1").Value;

//        // Convert sheet ID to long if needed
//        long.TryParse(sheetId, out long parsedSheetId);
//        Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);

//        var newRow = new Row();
//        newRow.Cells = new List<Cell>();
//        newRow.Cells.Add(new Cell
//        {
//            ColumnId = GetColumnIdByName(sheet, "NAME"),
//            Value = formData.NAME
//        });
//        newRow.Cells.Add(new Cell
//        {
//            ColumnId = GetColumnIdByName(sheet, "EMAIL"),
//            Value = formData.EMAIL
//        });
//        newRow.Cells.Add(new Cell
//        {
//            ColumnId = GetColumnIdByName(sheet, "FATHERNAME"),
//            Value = formData.FATHERNAME
//        });
//        newRow.Cells.Add(new Cell
//        {
//            ColumnId = GetColumnIdByName(sheet, "MOTHERNAME"),
//            Value = formData.MOTHERNAME
//        });
//        newRow.Cells.Add(new Cell
//        {
//            ColumnId = GetColumnIdByName(sheet, "GENDER"),
//            Value = formData.GENDER
//        });
//        smartsheet.SheetResources.RowResources.AddRows(parsedSheetId, new Row[] { newRow });




//        using (MemoryStream stream = new MemoryStream())
//        {
//            using (BinaryWriter writer = new BinaryWriter(stream))
//            {
//                smartsheet.SheetResources.GetSheetAsPDF(parsedSheetId, writer, PaperSize.A4);
//            }
//            var fileResult = new FileContentResult(stream.ToArray(), "application/pdf")
//            {
//                FileDownloadName = "SheetData" + formData.NAME + ".pdf"
//            };
//            return fileResult;
//        }


//        //return Ok("Data added successfully.");


//    }
//    catch (Exception ex)
//    {
//        return BadRequest(ex.Message);
//    }
//}