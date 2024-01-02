using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.IdentityModel.Tokens;
using MieProject.Helpers;
using MieProject.Models;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.IdentityModel.Tokens.Jwt;
using System.Runtime.InteropServices;
using System.Security.Claims;
using System.Text;

namespace MieProject.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class LoginAndRegisterController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;
        //private readonly string sheetId1;

        public LoginAndRegisterController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

        }
        //[HttpGet("GetSheetIds")]
        //public IActionResult Get()
        //{
        //    try
        //    {
        //        SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
        //        PaginatedResult<Sheet> sheets = smartsheet.SheetResources.ListSheets(null, new PaginationParameters(true, null, null), null);
        //        List<long> sheetIds = new List<long>();
        //        foreach (Sheet sheet in sheets.Data)
        //        {
        //            sheetIds.Add((long)sheet.Id);
        //        }
        //        return Ok(sheetIds);
        //    }
        //    catch (Exception ex)
        //    {
        //        return BadRequest(ex.Message);
        //    }
        //}
        //[HttpGet("GetSheetData")]
        //public IActionResult GetData()
        //{
        //    try
        //    {
        //        SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
        //        long sheetId = 8716767337598852;
        //        Sheet sheet = smartsheet.SheetResources.GetSheet(sheetId, null, null, null, null, null, null, null);
        //        List<List<object>> sheetData = new List<List<object>>();

        //        foreach (Row row in sheet.Rows)
        //        {
        //            List<object> rowData = new List<object>();
        //            foreach (Cell cell in row.Cells)
        //            {
        //                rowData.Add(cell.Value);
        //            }
        //            sheetData.Add(rowData);
        //        }
        //        return Ok(sheetData);
        //    }
        //    catch (Exception ex)
        //    {
        //        return BadRequest(ex.Message);
        //    }
        //}
        //[HttpGet("GetSheetDataAppSettings")]
        //public IActionResult GetDataex()
        //{
        //    try
        //    {
        //        SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
        //        string sheetId1 = configuration.GetSection("SmartsheetSettings:sheetId1").Value;
        //        long.TryParse(sheetId1, out long parsedSheetId);

        //        Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
        //        List<List<object>> sheetData = new List<List<object>>();

        //        foreach (Row row in sheet.Rows)
        //        {
        //            List<object> rowData = new List<object>();
        //            foreach (Cell cell in row.Cells)
        //            {
        //                rowData.Add(cell.Value);
        //            }
        //            sheetData.Add(rowData);
        //        }
        //        return Ok(sheetData);
        //    }
        //    catch (Exception ex)
        //    {
        //        return BadRequest(ex.Message);
        //    }
        //}

        [HttpGet("GetSheetDataKeyValue")]

        public IActionResult GetSheetData()
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

                // Retrieve sheet ID from configuration
                string sheetId1 = configuration.GetSection("SmartsheetSettings:SheetId1").Value;

                // Convert sheet ID to long if needed
                long.TryParse(sheetId1, out long parsedSheetId);

                // Get sheet by ID
                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);

                // Extract data from the sheet as key-value pairs
                List<Dictionary<string, object>> sheetData = new List<Dictionary<string, object>>();

                // Get column names
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
                        // Use column name as key
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




        [HttpPost("RegisterNew")]
        public async Task<IActionResult> RegisterNew([FromBody] Register formData)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                // Retrieve sheet ID from configuration
                string sheetId = configuration.GetSection("SmartsheetSettings:SheetId1").Value;

                // Convert sheet ID to long if needed
                long.TryParse(sheetId, out long parsedSheetId);
                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);

                var newRow = new Row();
                newRow.Cells = new List<Cell>();
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
                    ColumnId = GetColumnIdByName(sheet, "UserName"),
                    Value = formData.UserName
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Password"),
                    Value = formData.Password
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Email"),
                    Value = formData.Email
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "MobileNumber"),
                    Value = formData.MobileNumber
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Role"),
                    Value = formData.Role
                });

                // Validate Email
                var emailColumnId = GetColumnIdByName(sheet, "Email");
                var existingEmails = sheet.Rows.Select(row => row.Cells.FirstOrDefault(c => c.ColumnId == emailColumnId)?.Value?.ToString());

                if (existingEmails.Contains(formData.Email))
                {
                    return BadRequest("Email already exists.");
                }

                // Validate Username
                var usernameColumnId = GetColumnIdByName(sheet, "UserName");
                var existingUsernames = sheet.Rows.Select(row => row.Cells.FirstOrDefault(c => c.ColumnId == usernameColumnId)?.Value?.ToString());

                if (existingUsernames.Contains(formData.UserName))
                {
                    return BadRequest("Username already exists.");
                }

                // Validate Password
                if (formData.Password.Length < 8 || !HasAlphaNumeric(formData.Password))
                {
                    return BadRequest("Password should be at least 8 characters long and include alphanumeric characters.");
                }

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
        private bool HasAlphaNumeric(string input)
        {
            return input.Any(char.IsLetter) && input.Any(char.IsDigit);
        }


        [HttpPost("Login")]
        public IActionResult Login([FromBody] Register userData)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:SheetId1").Value;

                long.TryParse(sheetId, out long parsedSheetId);
                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);

                var usernameColumnId = GetColumnIdByName(sheet, "UserName");
                var passwordColumnId = GetColumnIdByName(sheet, "Password");
                var roleColumnId = GetColumnIdByName(sheet, "Role");


                if (usernameColumnId == 0 || passwordColumnId == 0)
                {
                    return BadRequest("Column not found");
                }

                var rows = sheet.Rows;

                foreach (var row in rows)
                {
                    var usernameCell = row.Cells.FirstOrDefault(c => c.ColumnId == usernameColumnId);
                    var passwordCell = row.Cells.FirstOrDefault(c => c.ColumnId == passwordColumnId);
                    var roleCell = row.Cells.FirstOrDefault(c => c.ColumnId == roleColumnId);

                    if (usernameCell?.Value?.ToString() == userData.UserName && passwordCell?.Value?.ToString() == userData.Password)
                    {
                        var username=usernameCell.Value?.ToString();
                        var password=passwordCell.Value?.ToString();
                        var role= roleCell.Value?.ToString();
                        // Additional logic if needed
                        var token = CreateJwt(username,role);

                        return Ok(new 
                        {Token=token, Message = "Login Success!" });
                    }
                }

                // If no matching credentials found
                return BadRequest("Username or Password Incorrect");
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

      
        private string CreateJwt(string username,string password)
        {
            var jwtTokenHandler = new JwtSecurityTokenHandler();
            var key = Encoding.ASCII.GetBytes("veryveryveryveryverysecret..................");
            var identity = new ClaimsIdentity(new Claim[]
            {
                new Claim(ClaimTypes.Name,username),
                new Claim(ClaimTypes.Role,password)
            });
            var credentials = new SigningCredentials(new SymmetricSecurityKey(key), SecurityAlgorithms.HmacSha256);

            var tokenDescriptor = new SecurityTokenDescriptor
            {
                Subject = identity,
                Expires = DateTime.Now.AddDays(1),
                SigningCredentials = credentials
            };
            var token = jwtTokenHandler.CreateToken(tokenDescriptor);
            return jwtTokenHandler.WriteToken(token);
        }


    }
}
