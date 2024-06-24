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
using IndiaEvents.Models.Models.MasterSheets;
using Newtonsoft.Json;

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












        //[HttpPost("Login")]
        //public IActionResult Login([FromBody] EmployeeMaster userData)
        //{
        //    try
        //    {
        //        SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
        //        string sheetId1 = configuration.GetSection("SmartsheetSettings:SheetId2").Value;
        //        Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId1);
        //        var Sheetcolumns = sheet.Columns.ToDictionary(column => column.Title, column => (long)column.Id);

        //        long EmailColumnId = Sheetcolumns["EmailId"];
        //        long UsernameColumnId = Sheetcolumns["UserName"];
        //        long passwordColumnId = Sheetcolumns["Password"];
        //        long IsActiveColumnId = Sheetcolumns["IsActive"];
        //        long roleColumnId = Sheetcolumns["Designation"];
        //        long ReportingManagerColumnId = Sheetcolumns["Reporting Manager"];
        //        long FirstLevelManagerId = Sheetcolumns["1stLevelManager"];
        //        long RBM_BMId = Sheetcolumns["RBM/BM"];
        //        long SalesHeadId = Sheetcolumns["Sales Head"];
        //        long FinanceHeadId = Sheetcolumns["Finance Head"];
        //        long MarketingHeadId = Sheetcolumns["Marketing Head"];
        //        long ComplianceId = Sheetcolumns["Compliance Head"];
        //        long FinanceCheckerId = Sheetcolumns["Finance Checker"];
        //        long MedicalAffairsHeadId = Sheetcolumns["Medical Affairs Head"];
        //        long FinanceTreasuryId = Sheetcolumns["Finance Treasury"];
        //        long FinanceAccountsId = Sheetcolumns["Finance Accounts"];
        //        long SalesCoordinatorId = Sheetcolumns["Sales Coordinator"];
        //        long MarketingCoordinatorId = Sheetcolumns["Marketing Coordinator"];


        //        if (EmailColumnId == 0 || passwordColumnId == 0)
        //        {
        //            return BadRequest("Column not found");
        //        }
        //        IEnumerable<Row> rows = [];
        //        int? EmailId = sheet.Columns.Where(y => y.Title == "EmailId").Select(z => z.Index).FirstOrDefault();
        //        int? Password = sheet.Columns.Where(y => y.Title == "Password").Select(z => z.Index).FirstOrDefault();

        //        rows = sheet.Rows.Where(x =>
        //        {
        //            string EmailIdVAlue = Convert.ToString(x.Cells[(int)EmailId].Value);
        //            string PasswordValue = Convert.ToString(x.Cells[(int)Password].Value);
        //            return EmailIdVAlue == userData.EmailId && PasswordValue == userData.Password;
        //        });
        //        int TotalCount = 0;
        //        foreach (var row in rows)
        //        {
        //            TotalCount = TotalCount + 1;
        //        }
        //        if (TotalCount > 0)
        //        {
        //            foreach (var row in rows)
        //            {
        //                Cell? EmailIdCell = row.Cells.FirstOrDefault(c => c.ColumnId == EmailColumnId);
        //                Cell? UsernameIdCell = row.Cells.FirstOrDefault(c => c.ColumnId == UsernameColumnId);
        //                Cell? passwordCell = row.Cells.FirstOrDefault(c => c.ColumnId == passwordColumnId);
        //                Cell? roleCell = row.Cells.FirstOrDefault(c => c.ColumnId == roleColumnId);
        //                Cell? ReportingManagerCell = row.Cells.FirstOrDefault(c => c.ColumnId == ReportingManagerColumnId);
        //                Cell? FirstLevelManagerCell = row.Cells.FirstOrDefault(c => c.ColumnId == FirstLevelManagerId);
        //                Cell? RBM_BMCell = row.Cells.FirstOrDefault(c => c.ColumnId == RBM_BMId);
        //                Cell? SalesHeadCell = row.Cells.FirstOrDefault(c => c.ColumnId == SalesHeadId);
        //                Cell? FinanceHeadCell = row.Cells.FirstOrDefault(c => c.ColumnId == FinanceHeadId);
        //                Cell? MarketingHeadCell = row.Cells.FirstOrDefault(c => c.ColumnId == MarketingHeadId);
        //                Cell? ComplianceCell = row.Cells.FirstOrDefault(c => c.ColumnId == ComplianceId);
        //                Cell? MedicalAffairsHeadCell = row.Cells.FirstOrDefault(c => c.ColumnId == MedicalAffairsHeadId);
        //                Cell? FinanceTreasuryCell = row.Cells.FirstOrDefault(c => c.ColumnId == FinanceTreasuryId);
        //                Cell? FinanceAccountsCell = row.Cells.FirstOrDefault(c => c.ColumnId == FinanceAccountsId);
        //                Cell? SalesCoordinatorCell = row.Cells.FirstOrDefault(c => c.ColumnId == SalesCoordinatorId);
        //                Cell? MarketingCoordinatorCell = row.Cells.FirstOrDefault(c => c.ColumnId == MarketingCoordinatorId);
        //                Cell? FinanceCheckerCell = row.Cells.FirstOrDefault(c => c.ColumnId == FinanceCheckerId);


        //                var isActiveCell = row.Cells.FirstOrDefault(c => c.ColumnId == IsActiveColumnId);
        //                if (isActiveCell?.Value?.ToString() == "No")
        //                {
        //                    return BadRequest("Employee is inactive");
        //                }
        //                string? username = UsernameIdCell.Value?.ToString();
        //                string? email = EmailIdCell.Value?.ToString();
        //                string? password = passwordCell.Value?.ToString();
        //                string? role = roleCell.Value?.ToString();
        //                string? ReportingManager = ReportingManagerCell.Value?.ToString();
        //                string? FirstLevelManager = FirstLevelManagerCell.Value?.ToString();
        //                string? RBM_BM = RBM_BMCell.Value?.ToString();
        //                string? SalesHead = SalesHeadCell.Value?.ToString();
        //                string? FinanceHead = FinanceHeadCell.Value?.ToString();
        //                string? MarketingHead = MarketingHeadCell.Value?.ToString();
        //                string? Compliance = ComplianceCell.Value?.ToString();
        //                string? MedicalAffairsHead = MedicalAffairsHeadCell.Value?.ToString();
        //                string? FinanceTreasury = FinanceTreasuryCell.Value?.ToString();
        //                string? FinanceAccounts = FinanceAccountsCell.Value?.ToString();
        //                string? SalesCoordinator = SalesCoordinatorCell.Value?.ToString();
        //                string? MarketingCoordinator = MarketingCoordinatorCell.Value?.ToString();
        //                string? FinanceChecker = FinanceCheckerCell.Value?.ToString();

        //                string token = CreateJwt(username, email, role, ReportingManager, FirstLevelManager, RBM_BM, SalesHead, FinanceHead, MarketingHead, Compliance, MedicalAffairsHead, FinanceTreasury, FinanceChecker, FinanceAccounts, SalesCoordinator, MarketingCoordinator);

        //                return Ok(new
        //                { Token = token, Message = "Login Success!" });



        //            }
        //        }


        //        var rows = sheet.Rows;


        //        return BadRequest("Username or Password Incorrect");
        //    }



        //    catch (Exception ex)
        //    {
        //        return BadRequest(ex.Message);
        //    }
        //}


        [HttpPost("Login")]
        public IActionResult Login([FromBody] EmployeeMaster userData)
        {
            try
            {
                string IsActiveVal = "";
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId1 = configuration.GetSection("SmartsheetSettings:SheetId2").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId1);
                var Sheetcolumns = sheet.Columns.ToDictionary(column => column.Title, column => (long)column.Id);

                long EmailColumnId = Sheetcolumns["EmailId"];
                long UsernameColumnId = Sheetcolumns["UserName"];
                long passwordColumnId = Sheetcolumns["Password"];
                long IsActiveColumnId = Sheetcolumns["IsActive"];
                long roleColumnId = Sheetcolumns["Designation"];
                long ReportingManagerColumnId = Sheetcolumns["Reporting Manager"];
                long FirstLevelManagerId = Sheetcolumns["1stLevelManager"];
                long RBM_BMId = Sheetcolumns["RBM/BM"];
                long SalesHeadId = Sheetcolumns["Sales Head"];
                long FinanceHeadId = Sheetcolumns["Finance Head"];
                long MarketingHeadId = Sheetcolumns["Marketing Head"];
                long ComplianceId = Sheetcolumns["Compliance Head"];
                long FinanceCheckerId = Sheetcolumns["Finance Checker"];
                long MedicalAffairsHeadId = Sheetcolumns["Medical Affairs Head"];
                long FinanceTreasuryId = Sheetcolumns["Finance Treasury"];
                long FinanceAccountsId = Sheetcolumns["Finance Accounts"];
                long SalesCoordinatorId = Sheetcolumns["Sales Coordinator"];
                long MarketingCoordinatorId = Sheetcolumns["Marketing Coordinator"];

                if (EmailColumnId == 0 || passwordColumnId == 0)
                {
                    return BadRequest("Column not found");
                }





                IEnumerable<Row> rows = sheet.Rows.Where(x =>
                {
                    string EmailIdValue = Convert.ToString(x.Cells.FirstOrDefault(c => c.ColumnId == EmailColumnId)?.Value);
                    string PasswordValue = Convert.ToString(x.Cells.FirstOrDefault(c => c.ColumnId == passwordColumnId)?.Value);
                    return EmailIdValue == userData.EmailId && PasswordValue == userData.Password;
                });



                Console.WriteLine(rows.Count());

                List<UserDetails> userDetailsList = new List<UserDetails>();

                foreach (var row in rows)
                {
                    var EmailIdCell = row.Cells.FirstOrDefault(c => c.ColumnId == EmailColumnId);
                    var UsernameIdCell = row.Cells.FirstOrDefault(c => c.ColumnId == UsernameColumnId);
                    var passwordCell = row.Cells.FirstOrDefault(c => c.ColumnId == passwordColumnId);
                    var roleCell = row.Cells.FirstOrDefault(c => c.ColumnId == roleColumnId);
                    var ReportingManagerCell = row.Cells.FirstOrDefault(c => c.ColumnId == ReportingManagerColumnId);
                    var FirstLevelManagerCell = row.Cells.FirstOrDefault(c => c.ColumnId == FirstLevelManagerId);
                    var RBM_BMCell = row.Cells.FirstOrDefault(c => c.ColumnId == RBM_BMId);
                    var SalesHeadCell = row.Cells.FirstOrDefault(c => c.ColumnId == SalesHeadId);
                    var FinanceHeadCell = row.Cells.FirstOrDefault(c => c.ColumnId == FinanceHeadId);
                    var MarketingHeadCell = row.Cells.FirstOrDefault(c => c.ColumnId == MarketingHeadId);
                    var ComplianceCell = row.Cells.FirstOrDefault(c => c.ColumnId == ComplianceId);
                    var MedicalAffairsHeadCell = row.Cells.FirstOrDefault(c => c.ColumnId == MedicalAffairsHeadId);
                    var FinanceTreasuryCell = row.Cells.FirstOrDefault(c => c.ColumnId == FinanceTreasuryId);
                    var FinanceAccountsCell = row.Cells.FirstOrDefault(c => c.ColumnId == FinanceAccountsId);
                    var SalesCoordinatorCell = row.Cells.FirstOrDefault(c => c.ColumnId == SalesCoordinatorId);
                    var MarketingCoordinatorCell = row.Cells.FirstOrDefault(c => c.ColumnId == MarketingCoordinatorId);
                    var FinanceCheckerCell = row.Cells.FirstOrDefault(c => c.ColumnId == FinanceCheckerId);
                    var isActiveCell = row.Cells.FirstOrDefault(c => c.ColumnId == IsActiveColumnId);

                    if (isActiveCell?.Value?.ToString() == "No")
                    {
                        IsActiveVal = "No";
                        //return BadRequest("Employee is inactive");
                    }

                    UserDetails userDetails = new UserDetails
                    {
                        unique_name = UsernameIdCell?.Value?.ToString(),
                        email = EmailIdCell?.Value?.ToString(),
                        role = roleCell?.Value?.ToString(),
                        reportingmanager = ReportingManagerCell?.Value?.ToString(),
                        firstLevelManager = FirstLevelManagerCell?.Value?.ToString(),
                        RBM_BM = RBM_BMCell?.Value?.ToString(),
                        SalesHead = SalesHeadCell?.Value?.ToString(),
                        FinanceHead = FinanceHeadCell?.Value?.ToString(),
                        MarketingHead = MarketingHeadCell?.Value?.ToString(),
                        ComplianceHead = ComplianceCell?.Value?.ToString(),
                        MedicalAffairsHead = MedicalAffairsHeadCell?.Value?.ToString(),
                        FinanceTreasury = FinanceTreasuryCell?.Value?.ToString(),
                        FinanceAccounts = FinanceAccountsCell?.Value?.ToString(),
                        SalesCoordinator = SalesCoordinatorCell?.Value?.ToString(),
                        MarketingCoordinator = MarketingCoordinatorCell?.Value?.ToString(),
                        FinanceChecker = FinanceCheckerCell?.Value?.ToString(),

                    };

                    userDetailsList.Add(userDetails);
                }

                if (userDetailsList.Count > 0)
                {
                    if (userDetailsList.Count == 1 && IsActiveVal == "No")
                    {
                        return BadRequest("Employee is inactive");
                    }
                    string token = CreateJwtNew(userDetailsList);
                    return Ok(new { Token = token, Message = "Login Success!" });
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
                Log.Information("google api stated");
                string GoogleclientId = configuration.GetSection("GoogleAuthentication:ClientId").Value;

                string sheetId1 = configuration.GetSection("SmartsheetSettings:SheetId1").Value;
                string sheetId2 = configuration.GetSection("SmartsheetSettings:SheetId2").Value;

                var settings = new GoogleJsonWebSignature.ValidationSettings()
                {
                    Audience = new List<string>() { GoogleclientId }
                };
                Log.Information("Audience successfully loded");
                var payload = await GoogleJsonWebSignature.ValidateAsync(credential, settings);
                Log.Information("payload successfully loded");
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
                Log.Error($"Error occured  {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(BadRequest(ex.Message));
            }

        }



        private string CreateJwt(string username, string email, string role, string reportingmanager, string firstLevelManager, string RBM_BM, string SalesHead, string FinanceHead, string MarketingHead, string compliance, string MedicalAffairsHead, string FinanceTreasury, string FinanceChecker, string FinanceAccounts, string SalesCoordinator, string MarketingCoordinator)
        {
            var jwtTokenHandler = new JwtSecurityTokenHandler();
            var key = new SymmetricSecurityKey(Encoding.UTF8.GetBytes("veryveryveryveryverysecret......................"));

            var identity = new Claim[]
            {
                new("unique_name", username),
                new("email", email),
                new("role", role),
                new("reportingmanager", reportingmanager),
                new("firstLevelManager", firstLevelManager),
                new("RBM_BM", RBM_BM),
                new("SalesHead", SalesHead),
                new("FinanceHead", FinanceHead),
                new("MarketingHead", MarketingHead),
                new("ComplianceHead", compliance),
                new("MedicalAffairsHead", MedicalAffairsHead),
                new("FinanceTreasury", FinanceTreasury),
                new("FinanceAccounts", FinanceAccounts),
                new("FinanceChecker", FinanceChecker),
                new("SalesCoordinator", SalesCoordinator),
                new("MarketingCoordinator", MarketingCoordinator),
                new(JwtRegisteredClaimNames.Jti, Guid.NewGuid().ToString())
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




        //private string CreateJwtNew1(List<UserDetails> userDetailsList)
        //{
        //    var jwtTokenHandler = new JwtSecurityTokenHandler();
        //    var key = new SymmetricSecurityKey(Encoding.UTF8.GetBytes("veryveryveryveryverysecret......................"));

        //    var claims = new List<Claim>();

        //    foreach (var user in userDetailsList)
        //    {
        //        claims.Add(new Claim("username", user.Username));
        //        claims.Add(new Claim("email", user.Email));
        //        claims.Add(new Claim("role", user.Role));
        //        claims.Add(new Claim("reportingmanager", user.ReportingManager));
        //        claims.Add(new Claim("firstLevelManager", user.FirstLevelManager));
        //        claims.Add(new Claim("RBM_BM", user.RBM_BM));
        //        claims.Add(new Claim("SalesHead", user.SalesHead));
        //        claims.Add(new Claim("FinanceHead", user.FinanceHead));
        //        claims.Add(new Claim("MarketingHead", user.MarketingHead));
        //        claims.Add(new Claim("ComplianceHead", user.Compliance));
        //        claims.Add(new Claim("MedicalAffairsHead", user.MedicalAffairsHead));
        //        claims.Add(new Claim("FinanceTreasury", user.FinanceTreasury));
        //        claims.Add(new Claim("FinanceAccounts", user.FinanceAccounts));
        //        claims.Add(new Claim("SalesCoordinator", user.SalesCoordinator));
        //        claims.Add(new Claim("MarketingCoordinator", user.MarketingCoordinator));
        //        claims.Add(new Claim("FinanceChecker", user.FinanceChecker));
        //    }

        //    claims.Add(new Claim(JwtRegisteredClaimNames.Jti, Guid.NewGuid().ToString()));

        //    var token = new JwtSecurityToken(
        //        issuer: "http://localhost:5098",
        //        audience: "ABM",
        //        expires: DateTime.Now.AddDays(1),
        //        claims: claims,
        //        signingCredentials: new SigningCredentials(key, SecurityAlgorithms.HmacSha256)
        //    );

        //    return jwtTokenHandler.WriteToken(token);
        //}

        //private string CreateJwtNew2(List<UserDetails> userDetailsList)
        //{
        //    var jwtTokenHandler = new JwtSecurityTokenHandler();
        //    var key = new SymmetricSecurityKey(Encoding.UTF8.GetBytes("veryveryveryveryverysecret......................"));

        //    var claims = new List<Claim>
        //    {
        //        new Claim(JwtRegisteredClaimNames.Jti, Guid.NewGuid().ToString())
        //    };

        //    // Serialize user details list to JSON and add as a claim
        //    string userDetailsJson = JsonConvert.SerializeObject(userDetailsList);
        //    claims.Add(new Claim("userDetails", userDetailsJson));

        //    var token = new JwtSecurityToken(
        //        issuer: "http://localhost:5098",
        //        audience: "ABM",
        //        expires: DateTime.Now.AddDays(1),
        //        claims: claims,
        //        signingCredentials: new SigningCredentials(key, SecurityAlgorithms.HmacSha256)
        //    );

        //    return jwtTokenHandler.WriteToken(token);
        //}

        private string CreateJwtNew(List<UserDetails> userDetailsList)
        {
            var jwtTokenHandler = new JwtSecurityTokenHandler();
            var key = new SymmetricSecurityKey(Encoding.UTF8.GetBytes("veryveryveryveryverysecret......................"));

            string userDetailsJson = JsonConvert.SerializeObject(userDetailsList);

            var claims = new List<Claim>
                {
                    new Claim(JwtRegisteredClaimNames.Jti, Guid.NewGuid().ToString()),
                    new Claim("userDetails", userDetailsJson, JsonClaimValueTypes.Json)
                };

            var token = new JwtSecurityToken(
                issuer: "http://localhost:5098",
                audience: "ABM",
                expires: DateTime.Now.AddDays(1),
                claims: claims,
                signingCredentials: new SigningCredentials(key, SecurityAlgorithms.HmacSha256)
            );

            return jwtTokenHandler.WriteToken(token);
        }

    }
}
