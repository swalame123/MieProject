using Google.Apis.Auth;
using Google.Apis.Requests;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.IdentityModel.Tokens;
using MieProject.Helpers;
using MieProject.Models;
using NPOI.SS.Formula;
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
        private readonly string clientId = "460379778026-nhgqueksa9p730jj0lokj8m5dv35jpr5.apps.googleusercontent.com";
        private readonly string clientSecret = "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";

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

                string sheetId1 = configuration.GetSection("SmartsheetSettings:SheetId1").Value;

                long.TryParse(sheetId1, out long parsedSheetId);

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

        [HttpPost("Login")]
        public IActionResult Login([FromBody] EmployeeMaster userData)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:SheetId1").Value;

                long.TryParse(sheetId, out long parsedSheetId);
                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);

                var EmailColumnId = GetColumnIdByName(sheet, "EmailId");
                var passwordColumnId = GetColumnIdByName(sheet, "Password");
                var IsActiveColumnId = GetColumnIdByName(sheet, "IsActive");
                var roleColumnId = GetColumnIdByName(sheet, "RoleName");


                if (EmailColumnId == 0 || passwordColumnId == 0 )
                {
                    return BadRequest("Column not found");
                }

                var rows = sheet.Rows;

                foreach (var row in rows)
                {
                    var EmailIdCell = row.Cells.FirstOrDefault(c => c.ColumnId == EmailColumnId);
                    var passwordCell = row.Cells.FirstOrDefault(c => c.ColumnId == passwordColumnId);
                    

                    var roleCell = row.Cells.FirstOrDefault(c => c.ColumnId == roleColumnId);

                    if (EmailIdCell?.Value?.ToString() == userData.EmailId && passwordCell?.Value?.ToString() == userData.Password)
                    {
                        var isActiveCell = row.Cells.FirstOrDefault(c => c.ColumnId == IsActiveColumnId);
                        if (isActiveCell?.Value?.ToString() == "No")
                        {
                            return BadRequest("Employee is inactive");
                        }
                        var username= EmailIdCell.Value?.ToString();
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
        


        private string CreateJwt(string username,string role)
        {
            var jwtTokenHandler = new JwtSecurityTokenHandler();
            var key = Encoding.ASCII.GetBytes("veryveryveryveryverysecret......................");
            var identity = new ClaimsIdentity(new Claim[]
            {
                new Claim(ClaimTypes.Name,username),
                new Claim(ClaimTypes.Role,role)
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
