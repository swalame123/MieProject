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
using IndiaEventsWebApi.Helper;
using Serilog;

namespace IndiaEventsWebApi.Controllers
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



        [HttpPost("Login")]
        public IActionResult Login([FromBody] EmployeeMaster userData)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId1 = configuration.GetSection("SmartsheetSettings:SheetId1").Value;
                string sheetId2 = configuration.GetSection("SmartsheetSettings:SheetId2").Value;


                List<string> Sheets = new List<string>() { sheetId1, sheetId2 };
                foreach (var sheetId in Sheets)
                {
                    long.TryParse(sheetId, out long parsedSheetId);
                    Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);


                    var EmailColumnId = SheetHelper.GetColumnIdByName(sheet, "EmailId");
                    var UsernameColumnId = SheetHelper.GetColumnIdByName(sheet, "UserName");
                    var passwordColumnId = SheetHelper.GetColumnIdByName(sheet, "Password");
                    var IsActiveColumnId = SheetHelper.GetColumnIdByName(sheet, "IsActive");
                    var roleColumnId = SheetHelper.GetColumnIdByName(sheet, "Designation");
                    var ReportingManagerColumnId = SheetHelper.GetColumnIdByName(sheet, "Reporting Manager");
                    var FirstLevelManagerId = SheetHelper.GetColumnIdByName(sheet, "1stLevelManager");
                    var RBM_BMId = SheetHelper.GetColumnIdByName(sheet, "RBM/BM");
                    var SalesHeadId = SheetHelper.GetColumnIdByName(sheet, "Sales Head");
                    var FinanceHeadId = SheetHelper.GetColumnIdByName(sheet, "Finance Head");
                    var MarketingHeadId = SheetHelper.GetColumnIdByName(sheet, "Marketing Head");
                    var ComplianceId = SheetHelper.GetColumnIdByName(sheet, "Compliance Head");
                    var FinanceCheckerId = SheetHelper.GetColumnIdByName(sheet, "Finance Checker");

                    var MedicalAffairsHeadId = SheetHelper.GetColumnIdByName(sheet, "Medical Affairs Head");
                    var FinanceTreasuryId = SheetHelper.GetColumnIdByName(sheet, "Finance Treasury");
                    var FinanceAccountsId = SheetHelper.GetColumnIdByName(sheet, "Finance Accounts");
                    var SalesCoordinatorId = SheetHelper.GetColumnIdByName(sheet, "Sales Coordinator");
                    var MarketingCoordinatorId = SheetHelper.GetColumnIdByName(sheet, "Marketing Coordinator");



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
                        var ReportingManagerCell = row.Cells.FirstOrDefault(c => c.ColumnId == ReportingManagerColumnId);
                        var FirstLevelManagerCell = row.Cells.FirstOrDefault(c => c.ColumnId == FirstLevelManagerId);
                        var RBM_BMCell = row.Cells.FirstOrDefault(c => c.ColumnId == RBM_BMId);
                        var SalesHeadCell = row.Cells.FirstOrDefault(c => c.ColumnId == SalesHeadId);
                        var FinanceHeadCell = row.Cells.FirstOrDefault(c => c.ColumnId ==FinanceHeadId);
                        var MarketingHeadCell = row.Cells.FirstOrDefault(c => c.ColumnId == MarketingHeadId);
                        var ComplianceCell = row.Cells.FirstOrDefault(c => c.ColumnId == ComplianceId);
                        var MedicalAffairsHeadCell = row.Cells.FirstOrDefault(c => c.ColumnId == MedicalAffairsHeadId);
                        var FinanceTreasuryCell = row.Cells.FirstOrDefault(c => c.ColumnId == FinanceTreasuryId);
                        var FinanceAccountsCell = row.Cells.FirstOrDefault(c => c.ColumnId == FinanceAccountsId);
                        var SalesCoordinatorCell = row.Cells.FirstOrDefault(c => c.ColumnId == SalesCoordinatorId);
                        var MarketingCoordinatorCell = row.Cells.FirstOrDefault(c => c.ColumnId == MarketingCoordinatorId);
                        var FinanceCheckerCell = row.Cells.FirstOrDefault(c => c.ColumnId == FinanceCheckerId);

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
                            var ReportingManager = ReportingManagerCell.Value?.ToString();
                            var FirstLevelManager = FirstLevelManagerCell.Value?.ToString();
                            var RBM_BM = RBM_BMCell.Value?.ToString();
                            var SalesHead = SalesHeadCell.Value?.ToString();
                            var FinanceHead = FinanceHeadCell.Value?.ToString();
                            var MarketingHead = MarketingHeadCell.Value?.ToString();
                            var Compliance = ComplianceCell.Value?.ToString();
                            var MedicalAffairsHead = MedicalAffairsHeadCell.Value?.ToString();
                            var FinanceTreasury = FinanceTreasuryCell.Value?.ToString();
                            var FinanceAccounts = FinanceAccountsCell.Value?.ToString();
                            var SalesCoordinator = SalesCoordinatorCell.Value?.ToString();
                            var MarketingCoordinator = MarketingCoordinatorCell.Value?.ToString();
                            var FinanceChecker = FinanceCheckerCell.Value?.ToString();

                            var token = CreateJwt(username, email, role, ReportingManager, FirstLevelManager, RBM_BM, SalesHead, FinanceHead, MarketingHead, Compliance, MedicalAffairsHead, FinanceTreasury, FinanceChecker, FinanceAccounts, SalesCoordinator, MarketingCoordinator);

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

                string GoogleclientId = configuration.GetSection("GoogleAuthentication:ClientId").Value;

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

                    var EmailColumnId = SheetHelper.GetColumnIdByName(sheet, "EmailId");
                    var UsernameColumnId = SheetHelper.GetColumnIdByName(sheet, "UserName");
                    var ComplianceId = SheetHelper.GetColumnIdByName(sheet, "Compliance Head");
                    var IsActiveColumnId = SheetHelper.GetColumnIdByName(sheet, "IsActive");
                    var roleColumnId = SheetHelper.GetColumnIdByName(sheet, "Designation");
                    var ReportingManagerColumnId = SheetHelper.GetColumnIdByName(sheet, "Reporting Manager");
                    var FirstLevelManagerId = SheetHelper.GetColumnIdByName(sheet, "1stLevelManager");
                    var RBM_BMId = SheetHelper.GetColumnIdByName(sheet, "RBM/BM");
                    var SalesHeadId = SheetHelper.GetColumnIdByName(sheet, "Sales Head");
                    var FinanceHeadId = SheetHelper.GetColumnIdByName(sheet, "Finance Head");
                    var MarketingHeadId = SheetHelper.GetColumnIdByName(sheet, "Marketing Head");
                    var MedicalAffairsHeadId = SheetHelper.GetColumnIdByName(sheet, "Medical Affairs Head");
                    var FinanceTreasuryId = SheetHelper.GetColumnIdByName(sheet, "Finance Treasury");
                    var FinanceAccountsId = SheetHelper.GetColumnIdByName(sheet, "Finance Accounts");
                    var SalesCoordinatorId = SheetHelper.GetColumnIdByName(sheet, "Sales Coordinator");
                    var MarketingCoordinatorId = SheetHelper.GetColumnIdByName(sheet, "Marketing Coordinator");
                    var FinanceCheckerId = SheetHelper.GetColumnIdByName(sheet, "Finance Checker");

                    if (EmailColumnId == 0)
                    {
                        return BadRequest("Column not found");
                    }

                    var rows = sheet.Rows;

                    foreach (var row in rows)
                    {
                        var EmailIdCell = row.Cells.FirstOrDefault(c => c.ColumnId == EmailColumnId);
                        var UsernameIdCell = row.Cells.FirstOrDefault(c => c.ColumnId == UsernameColumnId);
                        var ComplianceCell = row.Cells.FirstOrDefault(c => c.ColumnId == ComplianceId);
                        var roleCell = row.Cells.FirstOrDefault(c => c.ColumnId == roleColumnId);
                        var ReportingManagerCell = row.Cells.FirstOrDefault(c => c.ColumnId == ReportingManagerColumnId);
                        var FirstLevelManagerCell = row.Cells.FirstOrDefault(c => c.ColumnId == FirstLevelManagerId);
                        var RBM_BMCell = row.Cells.FirstOrDefault(c => c.ColumnId == RBM_BMId);
                        var SalesHeadCell = row.Cells.FirstOrDefault(c => c.ColumnId == SalesHeadId);
                        var FinanceHeadCell = row.Cells.FirstOrDefault(c => c.ColumnId == FinanceHeadId);
                        var MarketingHeadCell = row.Cells.FirstOrDefault(c => c.ColumnId == MarketingHeadId);
                        var MedicalAffairsHeadCell = row.Cells.FirstOrDefault(c => c.ColumnId == MedicalAffairsHeadId);
                        var FinanceTreasuryCell = row.Cells.FirstOrDefault(c => c.ColumnId == FinanceTreasuryId);
                        var FinanceAccountsCell = row.Cells.FirstOrDefault(c => c.ColumnId == FinanceAccountsId);
                        var SalesCoordinatorCell = row.Cells.FirstOrDefault(c => c.ColumnId == SalesCoordinatorId);
                        var MarketingCoordinatorCell = row.Cells.FirstOrDefault(c => c.ColumnId == MarketingCoordinatorId);
                        var FinanceCheckerCell = row.Cells.FirstOrDefault(c => c.ColumnId == FinanceCheckerId);
                        if (EmailIdCell?.Value?.ToString() == payload.Email)
                        {
                            var isActiveCell = row.Cells.FirstOrDefault(c => c.ColumnId == IsActiveColumnId);
                            if (isActiveCell?.Value?.ToString() == "No")
                            {
                                return BadRequest("Employee is inactive");
                            }
                            var username = UsernameIdCell.Value?.ToString();
                            var email = EmailIdCell.Value?.ToString();
                            var Compliance = ComplianceCell.Value?.ToString();
                            var role = roleCell.Value?.ToString();
                            var ReportingManager = ReportingManagerCell.Value?.ToString();
                            var FirstLevelManager = FirstLevelManagerCell.Value?.ToString();
                            var RBM_BM = RBM_BMCell.Value?.ToString();
                            var SalesHead = SalesHeadCell.Value?.ToString();
                            var FinanceHead = FinanceHeadCell.Value?.ToString();
                            var MarketingHead = MarketingHeadCell.Value?.ToString();
                            var MedicalAffairsHead = MedicalAffairsHeadCell.Value?.ToString();
                            var FinanceTreasury = FinanceTreasuryCell.Value?.ToString();
                            var FinanceAccounts = FinanceAccountsCell.Value?.ToString();
                            var SalesCoordinator = SalesCoordinatorCell.Value?.ToString();
                            var MarketingCoordinator = MarketingCoordinatorCell.Value?.ToString();
                            var FinanceChecker = FinanceCheckerCell.Value?.ToString();
                            var token = CreateJwt(username, email, role, ReportingManager, FirstLevelManager, RBM_BM, SalesHead, FinanceHead, MarketingHead, Compliance, MedicalAffairsHead, FinanceTreasury, FinanceChecker, FinanceAccounts, SalesCoordinator, MarketingCoordinator);

                            return Ok(new
                            { Token = token, Message = "Login Success!" });
                        }
                    }
                }



                return BadRequest("Username or Password Incorrect");

            }
            catch (Exception ex)
            {
                Log.Error($"Error occured on Webhook apicontroller Attachementfile method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(BadRequest(ex.Message));
            }

        }

        private string CreateJwt(string username, string email, string role, string reportingmanager, string firstLevelManager, string RBM_BM, string SalesHead,string FinanceHead, string MarketingHead, string compliance, string MedicalAffairsHead, string FinanceTreasury,string FinanceChecker, string FinanceAccounts, string SalesCoordinator,string MarketingCoordinator)
        {
            var jwtTokenHandler = new JwtSecurityTokenHandler();
            var key = new SymmetricSecurityKey(Encoding.UTF8.GetBytes("veryveryveryveryverysecret......................"));

            var identity = new Claim[]
            {
        //new Claim(ClaimTypes.Name, username),
        //new Claim(ClaimTypes.Email, email),
        //new Claim(ClaimTypes.Role, role),
        new Claim("unique_name", username),
        new Claim("email", email),
        new Claim("role", role),
        new Claim("reportingmanager", reportingmanager),
        new Claim("firstLevelManager", firstLevelManager),
        new Claim("RBM_BM", RBM_BM),
        new Claim("SalesHead", SalesHead),
        new Claim("FinanceHead", FinanceHead),
        new Claim("MarketingHead", MarketingHead),
        new Claim("ComplianceHead", compliance),
        new Claim("MedicalAffairsHead", MedicalAffairsHead),
        new Claim("FinanceTreasury", FinanceTreasury),
        new Claim("FinanceAccounts", FinanceAccounts),
        new Claim("FinanceChecker", FinanceChecker),
        new Claim("SalesCoordinator", SalesCoordinator),
        new Claim("MarketingCoordinator", MarketingCoordinator),
        new Claim(JwtRegisteredClaimNames.Jti, Guid.NewGuid().ToString())
            };

            // Combine roles into a comma-separated string
            //var roles = string.Join(",", new[] { "AM", "ABM", "BM", "RBM", "MM" });

            var token = new JwtSecurityToken(
                issuer: "http://localhost:5098",
                audience: "ABM", // Use roles as the audience
                expires: DateTime.Now.AddDays(1),
                //expires: DateTime.Now.AddMinutes(5),
                claims: identity,
                signingCredentials: new SigningCredentials(key, SecurityAlgorithms.HmacSha256)
            );

            return jwtTokenHandler.WriteToken(token);
        }

    }
}
