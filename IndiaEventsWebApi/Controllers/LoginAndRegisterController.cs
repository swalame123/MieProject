using Google.Apis.Auth;
using Google.Apis.Requests;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.IdentityModel.Tokens;
using IndiaEventsWebApi.Helpers;
using IndiaEventsWebApi.Models;
using NPOI.SS.Formula;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.IdentityModel.Tokens.Jwt;
using System.Runtime.InteropServices;
using System.Security.Claims;
using System.Text;

namespace IndiaEventsWebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class LoginAndRegisterController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;
        private readonly string clientId = "200698853522-5b3nkgrgal38n7eqjqrrt6biinbt46ca.apps.googleusercontent.com";
        private readonly string clientSecret = "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";
        private string GoogleClientId;


        //private readonly string sheetId1;

        public LoginAndRegisterController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

        }
    

        [HttpGet("GetSheetData")]

        public IActionResult GetSheetData()
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

                List<Dictionary<string, object>> sheetData = new List<Dictionary<string, object>>();

                string sheetId1 = configuration.GetSection("SmartsheetSettings:SheetId1").Value;
                string sheetId2 = configuration.GetSection("SmartsheetSettings:SheetId2").Value;

                List<string> Sheets = new List<string>() { sheetId1, sheetId2 };
                foreach (var sheetId in Sheets)
                {
                    long.TryParse(sheetId, out long parsedSheetId);
                    Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
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
                string sheetId1 = configuration.GetSection("SmartsheetSettings:SheetId1").Value;
                string sheetId2 = configuration.GetSection("SmartsheetSettings:SheetId2").Value;


                List<string> Sheets = new List<string>() {sheetId1 ,sheetId2};
                foreach(var sheetId in Sheets)
                {
                    long.TryParse(sheetId, out long parsedSheetId);
                    Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);

                   
                    var EmailColumnId = GetColumnIdByName(sheet, "EmailId");
                    var UsernameColumnId = GetColumnIdByName(sheet, "UserName");
                    var passwordColumnId = GetColumnIdByName(sheet, "Password");
                    var IsActiveColumnId = GetColumnIdByName(sheet, "IsActive");
                    var roleColumnId = GetColumnIdByName(sheet, "Designation");


                    if (EmailColumnId == 0 || passwordColumnId == 0)
                    {
                        return BadRequest("Column not found");
                    }

                    var rows = sheet.Rows;

                    foreach (var row in rows)
                    {
                        //var EmailIdCell = row.Cells.FirstOrDefault(c => c.ColumnId == EmailColumnId);
                        var EmailIdCell = row.Cells.FirstOrDefault(c => c.ColumnId == EmailColumnId);
                        var UsernameIdCell = row.Cells.FirstOrDefault(c => c.ColumnId == UsernameColumnId);
                        var passwordCell = row.Cells.FirstOrDefault(c => c.ColumnId == passwordColumnId);


                        var roleCell = row.Cells.FirstOrDefault(c => c.ColumnId == roleColumnId);

                        if (EmailIdCell?.Value?.ToString() == userData.EmailId && passwordCell?.Value?.ToString() == userData.Password)
                        {
                            var isActiveCell = row.Cells.FirstOrDefault(c => c.ColumnId == IsActiveColumnId);
                            if (isActiveCell?.Value?.ToString() == "No")
                            {
                                return BadRequest("Employee is inactive");
                            }
                            //var username = EmailIdCell.Value?.ToString();
                            var username = UsernameIdCell.Value?.ToString();
                            var email = EmailIdCell.Value?.ToString();
                            var password = passwordCell.Value?.ToString();
                            var role = roleCell.Value?.ToString();
                           
                            var token = CreateJwt(username,email, role);

                            return Ok(new
                            { Token = token, Message = "Login Success!" });
                        }
                    }

                }
                return BadRequest("Username or Password Incorrect");
            }
               

               
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        [HttpPost("LoginWithGoogle")]
        public async Task<IActionResult> LoginWithGoogle([FromBody] string credential)
        {
            try
            {

                string GoogleclientId = "200698853522-5b3nkgrgal38n7eqjqrrt6biinbt46ca.apps.googleusercontent.com";
                ////200698853522 - 5b3nkgrgal38n7eqjqrrt6biinbt46ca.apps.googleusercontent.com
                string sheetId1 = configuration.GetSection("SmartsheetSettings:SheetId1").Value;
                string sheetId2 = configuration.GetSection("SmartsheetSettings:SheetId2").Value;

                var settings = new GoogleJsonWebSignature.ValidationSettings()
                {
                    Audience = new List<string>() { GoogleclientId }
                };
                var payload = await GoogleJsonWebSignature.ValidateAsync(credential, settings);

                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

                List<string> Sheets = new List<string>() { sheetId1, sheetId2 };
                foreach (var sheetId in Sheets)
                {
                    //string sheetIds = configuration.GetSection("SmartsheetSettings:sheetId").Value;

                    long.TryParse(sheetId, out long parsedSheetId);
                    Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);

                    var EmailColumnId = GetColumnIdByName(sheet, "EmailId");
                    var UsernameColumnId = GetColumnIdByName(sheet, "UserName");

                    var IsActiveColumnId = GetColumnIdByName(sheet, "IsActive");
                    var roleColumnId = GetColumnIdByName(sheet, "Designation");


                    if (EmailColumnId == 0)
                    {
                        return BadRequest("Column not found");
                    }

                    var rows = sheet.Rows;

                    foreach (var row in rows)
                    {
                        var EmailIdCell = row.Cells.FirstOrDefault(c => c.ColumnId == EmailColumnId);
                        var UsernameIdCell = row.Cells.FirstOrDefault(c => c.ColumnId == UsernameColumnId);

                        var roleCell = row.Cells.FirstOrDefault(c => c.ColumnId == roleColumnId);

                        if (EmailIdCell?.Value?.ToString() == payload.Email)
                        {
                            var isActiveCell = row.Cells.FirstOrDefault(c => c.ColumnId == IsActiveColumnId);
                            if (isActiveCell?.Value?.ToString() == "No")
                            {
                                return BadRequest("Employee is inactive");
                            }
                            var username = UsernameIdCell.Value?.ToString();
                            var email = EmailIdCell.Value?.ToString();

                            var role = roleCell.Value?.ToString();

                            var token = CreateJwt(username,email, role);

                            return Ok(new
                            { Token = token, Message = "Login Success!" });
                        }
                    }
                }
               
                

                return BadRequest("Username or Password Incorrect");

            }
            catch (Exception ex)
            {
                return BadRequest(BadRequest(ex.Message));
            }
            
        }
        private string CreateJwt(string username, string email,string role)
        {
            var jwtTokenHandler = new JwtSecurityTokenHandler();
            var key = Encoding.ASCII.GetBytes("veryveryveryveryverysecret......................");
            var identity = new ClaimsIdentity(new Claim[]
            {
                new Claim(ClaimTypes.Name,username),
                new Claim(ClaimTypes.Email,email),
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
