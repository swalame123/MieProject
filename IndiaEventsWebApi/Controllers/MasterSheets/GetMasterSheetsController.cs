using Aspose.Pdf.Plugins;
using IndiaEvents.Models.Models.GetData;
using IndiaEvents.Models.Models.Webhook;
using IndiaEventsWebApi.Helper;
using IndiaEventsWebApi.Models.MasterSheets;
using IndiaEventsWebApi.Models.RequestSheets;
using LazyCache;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Caching.Memory;
using MySqlConnector;
using Serilog;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Data;

namespace IndiaEventsWebApi.Controllers.MasterSheets
{
    [Route("api/[controller]")]
    [ApiController]
    //[Authorize]
    public class GetMasterSheetsController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;
        private readonly SmartsheetClient smartsheet;
        private ICacheProvider _cacheProvider;

        public GetMasterSheetsController(IConfiguration configuration, ICacheProvider cacheProvider)
        {
            this.configuration = configuration;
            this._cacheProvider = cacheProvider;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
            smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
        }
        [HttpGet("GetApprovedSpeakersData")]
        public IActionResult GetApprovedSpeakersData()
        {
            try
            {

                string sheetId = configuration.GetSection("SmartsheetSettings:ApprovedSpeakers").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = new List<Dictionary<string, object>>();
                List<string> columnNames = new List<string>();
                foreach (Column column in sheet.Columns)
                {
                    columnNames.Add(column.Title);
                }
                foreach (Row row in sheet.Rows)
                {
                    int isActiveColumnIndex = columnNames.IndexOf("Is Active");
                    if (isActiveColumnIndex >= 0 && row.Cells[isActiveColumnIndex].Value?.ToString().Equals("Yes", StringComparison.OrdinalIgnoreCase) == true)
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
        [HttpGet("GetApprovedTrainersData")]
        public IActionResult GetApprovedTrainersData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:ApprovedTrainers").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = new List<Dictionary<string, object>>();
                List<string> columnNames = new List<string>();
                foreach (Column column in sheet.Columns)
                {
                    columnNames.Add(column.Title);
                }
                foreach (Row row in sheet.Rows)
                {
                    int isActiveColumnIndex = columnNames.IndexOf("IsActive");
                    if (isActiveColumnIndex >= 0 && row.Cells[isActiveColumnIndex].Value?.ToString().Equals("Yes", StringComparison.OrdinalIgnoreCase) == true)
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
        [HttpGet("GetApprovedMastersData")]
        public IActionResult GetApprovedMastersData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:ApprovedMaster").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("GetObjectiveCriteriaData")]
        public IActionResult GetObjectiveCriteriaData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:ObjectiveCriteria").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("GetLegitimateNeedData")]
        public IActionResult GetLegitimateNeedData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:LegitimateNeed").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("VenueSelectionChecklistMaster")]
        public IActionResult VenueSelectionChecklistMaster()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:VenueSelectionChecklistMaster").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("GetBrandNameData")]
        public IActionResult GetBrandNameData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:BrandMaster").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("TestVenueSelectionChecklistMaster")]
        public IActionResult TestVenueSelectionChecklistMaster()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:BrandMaster").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);

                List<EventRequestBrandsList> eventRequestBrandsList = sheetData.Select(row =>
                    new EventRequestBrandsList
                    {
                        BrandName = row.ContainsKey("BrandName") ? row["BrandName"].ToString() : null,
                        PercentAllocation = row.ContainsKey("BrandId") ? row["BrandId"].ToString() : null,
                        ProjectId = row.ContainsKey("ProjectId") ? row["ProjectId"].ToString() : null
                    }).ToList();

                return Ok(eventRequestBrandsList);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("GetIndicationsMasterData")]
        public IActionResult GetIndicationsMasterData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:IndicatorsMaster").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("GetCityNameData")]
        public IActionResult GetCityNameData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:City").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
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

                        if ((columnNames[i] == "CityId") || (columnNames[i] == "CityName") || (columnNames[i] == "StateId"))
                        {
                            rowData[columnNames[i]] = row.Cells[i].Value;
                        }


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
        [HttpGet("GetClassIIITypeMasterData")]
        public IActionResult GetClassIIITypeMasterData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:ClassIIITypeMaster").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("GetDHProductNameData")]
        public IActionResult GetDHProductNameData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:DHProductName").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("GetDivisionMasterData")]
        public IActionResult GetDivisionMasterData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:DivisionMaster").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("GetSheetData")]
        public IActionResult GetSheetData()
        {
            try
            {
                List<Dictionary<string, object>> sheetData = new List<Dictionary<string, object>>();
                string sheetId1 = configuration.GetSection("SmartsheetSettings:SheetId1").Value;
                string sheetId2 = configuration.GetSection("SmartsheetSettings:SheetId2").Value;

                List<string> Sheets = new List<string>() { sheetId1, sheetId2 };
                foreach (var sheetId in Sheets)
                {
                    Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
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
        [HttpGet("GetEmployeeRoleMasterData")]
        public IActionResult GetEmployeeRoleMasterData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:EmployeeRoleMaster").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("GetEventData")]
        public IActionResult GetEventData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:EventType").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
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
                        if ((columnNames[i] == "EventTypeId") || (columnNames[i] == "EventType") || (columnNames[i] == "Roles"))
                        {
                            rowData[columnNames[i]] = row.Cells[i].Value;
                        }
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
        [HttpGet("GetExpenseTypeMasterData")]
        public IActionResult GetExpenseTypeMasterData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:ExpenseTypeMaster").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        //[HttpGet("HcpMaster")]
        //public IActionResult HcpMaster(int count = 5)
        //{
        //    int loopCount = 0;
        //    while (loopCount < count)
        //    {
        //        try
        //        {
        //            List<Dictionary<string, object>> sheetData = new List<Dictionary<string, object>>();
        //            string sheetId1 = configuration.GetSection("SmartsheetSettings:HcpMaster1").Value;
        //            string sheetId2 = configuration.GetSection("SmartsheetSettings:HcpMaster2").Value;
        //            string sheetId3 = configuration.GetSection("SmartsheetSettings:HcpMaster3").Value;
        //            string sheetId4 = configuration.GetSection("SmartsheetSettings:HcpMaster4").Value;

        //            List<string> Sheets = new List<string>() { sheetId1, sheetId2, sheetId3, sheetId4 };

        //            foreach (var sheetId in Sheets)
        //            {

        //                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
        //                Log.Information("executed the data in sheet"+sheetId+"---"+DateTime.Now);
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
        //            }
        //            return Ok(sheetData);
        //        }
        //        catch (SmartsheetException ex)
        //        {
        //            Log.Error($"Error occured on SmartsheetException method {ex.Message} at {DateTime.Now}");
        //            Log.Error(ex.StackTrace);
        //            if (loopCount < count)
        //            {
        //                loopCount++;
        //            }
        //            else
        //            {
        //                return BadRequest(new
        //                {
        //                    Message = ex.Message + "------" + ex.StackTrace
        //                });
        //            }

        //        }
        //        catch (Exception ex)
        //        {
        //            Log.Error($"Error occured on GetData method {ex.Message} at {DateTime.Now}");
        //            Log.Error(ex.StackTrace);
        //            if (loopCount < count)
        //            {
        //                loopCount++;
        //                //await Task.Delay(2000);
        //            }
        //            else
        //            {
        //                return BadRequest(new
        //                {
        //                    Message = ex.Message + "------" + ex.StackTrace
        //                });
        //            }

        //        }
        //        //catch (Exception ex)
        //        //{

        //        //    return BadRequest(new
        //        //    {
        //        //        Message = ex.Message + "------" + ex.StackTrace
        //        //    });
        //        //}

        //    }
        //    return Ok();
        //}

        [HttpGet("HcpMaster")]//rock  HcpMasterSmartSheet
        public async Task<IActionResult> HcpMaster()
        {
            try
            {
                Log.Information("Hcp Master loop started at " + DateTime.Now);
                string[] sheetIds = {
                            configuration.GetSection("SmartsheetSettings:HcpMaster1").Value,
                            configuration.GetSection("SmartsheetSettings:HcpMaster2").Value,
                            configuration.GetSection("SmartsheetSettings:HcpMaster3").Value,
                            configuration.GetSection("SmartsheetSettings:HcpMaster4").Value
                        };
                if (!_cacheProvider.TryGetValue(HcpCacheData.HcpData, out List<Dictionary<string, object>> HcpData))
                {
                    HcpData = ApiCalls.HcpData(sheetIds, smartsheet);
                    var cacheEntryOption = new MemoryCacheEntryOptions
                    {
                        AbsoluteExpiration = DateTime.Now.AddHours(5),
                        SlidingExpiration = TimeSpan.FromHours(5),
                        Size = 40000000
                    };
                    _cacheProvider.Set(HcpCacheData.HcpData, HcpData, cacheEntryOption);
                }
                Log.Information("Hcp Master loop Ended at " + DateTime.Now);
                return Ok(HcpData);
            }
            catch (Exception ex)
            {
                Log.Error($"Unexpected error occurred: {ex.Message}");
                return BadRequest(new { Message = ex.Message });
            }


        }

        [HttpGet("HcpMasterUAT")] //UAT  HcpMaster
        public async Task<IActionResult> HcpMasterUAT()
        {
            try
            {
                string MyConnection = configuration.GetSection("ConnectionStrings:mysql").Value;
                MySqlConnection MyConn2 = new MySqlConnection(MyConnection);
                MySqlCommand MyCommand2 = new MySqlCommand("GetHCPMaster", MyConn2);
                MyCommand2.CommandType = CommandType.StoredProcedure;
                await MyConn2.OpenAsync();
                MySqlDataAdapter MyAdapter = new MySqlDataAdapter();
                MyAdapter.SelectCommand = MyCommand2;
                DataTable dTable = new DataTable();
                MyAdapter.Fill(dTable);
                string JsonData = Newtonsoft.Json.JsonConvert.SerializeObject(dTable);
                await MyConn2.CloseAsync();
                return Ok(JsonData);

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        //    return Ok();
        //    //string connStr = "server=10.9.128.92;user=menarini;database=IN-Events;port=3306;password=Men@rini123";
        //    //MySqlConnection conn = new MySqlConnection(connStr);
        //    //try
        //    //{
        //    //    Console.WriteLine("Connecting to MySQL...");
        //    //    conn.Open();

        //    //    string sql = "select * from HCPMaster";
        //    //    MySqlCommand cmd = new MySqlCommand(sql, conn);
        //    //    MySqlDataReader rdr = cmd.ExecuteReader();

        //    //    while (rdr.Read())
        //    //    {
        //    //        Console.WriteLine(rdr[0] + " -- " + rdr[1]);
        //    //    }
        //    //    rdr.Close();
        //    //}
        //    //catch (Exception ex)
        //    //{
        //    //    Console.WriteLine(ex.ToString());
        //    //    return BadRequest(ex.Message);
        //    //}

        //    //conn.Close();
        //    //Console.WriteLine("Done.");
        //    //return Ok();
        //}
        ////[HttpGet("SqlHcpMaster")]
        ////public async Task<IActionResult> SqlHcpMaster(int id)
        ////{
        ////    using var connection = await database.OpenConnectionAsync();
        ////    using var command = connection.CreateCommand();
        ////    command.CommandText = @"SELECT `Id`, `Title`, `Content` FROM `BlogPost` WHERE `Id` = @id";
        ////    command.Parameters.AddWithValue("@id", id);
        ////    var result = await ReadAllAsync(await command.ExecuteReaderAsync());
        ////    return result.FirstOrDefault();
        ////}

        [HttpGet("GetRowsDataInHcpMasterByMisCode")]
        public IActionResult GetRowsDataInHcpMasterByMisCode(string? searchValue)
        {
            try
            {
                List<Dictionary<string, object>> sheetData = new List<Dictionary<string, object>>();
                string sheetId1 = configuration.GetSection("SmartsheetSettings:HcpMaster1").Value;
                string sheetId2 = configuration.GetSection("SmartsheetSettings:HcpMaster2").Value;
                string sheetId3 = configuration.GetSection("SmartsheetSettings:HcpMaster3").Value;
                string sheetId4 = configuration.GetSection("SmartsheetSettings:HcpMaster4").Value;
                List<string> Sheets = new List<string>() { sheetId1, sheetId2, sheetId3, sheetId4 };
                if (searchValue != null)
                {


                    foreach (var sheetId in Sheets)
                    {
                        Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                        List<string> columnNames = new List<string>();

                        foreach (Column column in sheet.Columns)
                        {
                            columnNames.Add(column.Title);
                        }

                        int miscodeColumnIndex = columnNames.IndexOf("MisCode");

                        foreach (Row row in sheet.Rows)
                        {
                            Cell miscodeCell = row.Cells.FirstOrDefault(cell => cell.ColumnId == sheet.Columns[miscodeColumnIndex].Id);
                            if (miscodeCell != null && miscodeCell.Value.ToString().Contains(searchValue))
                            {
                                Dictionary<string, object> rowData = new Dictionary<string, object>();
                                for (int i = 0; i < row.Cells.Count && i < columnNames.Count; i++)
                                {
                                    rowData[columnNames[i]] = row.Cells[i].Value;
                                }
                                sheetData.Add(rowData);
                            }

                        }
                    }


                }
                else
                {
                    foreach (var sheetId in Sheets)
                    {
                        Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
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
                }
                return Ok(sheetData);
            }
            catch (Exception ex)
            {

                return BadRequest(new
                {
                    Message = ex.Message + "------" + ex.StackTrace
                });
            }
        }
        [HttpGet("GetHCPRoleData")]
        public IActionResult GetHCPRoleData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:HCPRole").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
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

                        if ((columnNames[i] == "HCPRoleID") || (columnNames[i] == "HCPRole"))
                        {
                            rowData[columnNames[i]] = row.Cells[i].Value;
                        }
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
        [HttpGet("GetInitiatorMasterData")]
        public IActionResult GetInitiatorMasterData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:InitiatorMaster").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("GetMedicalUtilityData")]
        public IActionResult GetMedicalUtilityData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:MedicalUtility").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("GetRoleData")]
        public IActionResult GetRoleData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:RoleMaster").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
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
                        if ((columnNames[i] == "RoleId") || (columnNames[i] == "RoleName"))
                        {
                            rowData[columnNames[i]] = row.Cells[i].Value;
                        }
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
        [HttpGet("SlideKitMaster")]
        public IActionResult SlideKitMaster()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:SlideKitMaster").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("GetSpeakerCategoriesData")]
        public IActionResult GetSpeakerCategoriesData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:SpeakerCategories").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("GetStateNameData")]
        public IActionResult GetStateNameData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:State").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
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
                        if ((columnNames[i] == "StateId") || (columnNames[i] == "StateName"))
                        {
                            rowData[columnNames[i]] = row.Cells[i].Value;
                        }
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
        [HttpGet("GetTrainerCategoriesData")]
        public IActionResult GetTrainerCategoriesData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:TrainerCategories").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("VendorMasterSheetData")]
        public IActionResult VendorMasterSheetData()
        {
            try
            {

                string sheetId = configuration.GetSection("SmartsheetSettings:VendorMasterSheet").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = new List<Dictionary<string, object>>();
                List<string> columnNames = new List<string>();
                foreach (Column column in sheet.Columns)
                {
                    columnNames.Add(column.Title);
                }
                foreach (Row row in sheet.Rows)
                {
                    int isActiveColumnIndex = columnNames.IndexOf("IsActive?");
                    if (isActiveColumnIndex >= 0 && row.Cells[isActiveColumnIndex].Value?.ToString().Equals("Yes", StringComparison.OrdinalIgnoreCase) == true)
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
        [HttpGet("DeviationMasterSheetData")]
        public IActionResult DeviationMasterSheetData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:DeviationMaster").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("GetApprovedSpeakerSheetData")]
        public IActionResult GetApprovedSpeakerSheetData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:ApprovedSpeakers").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                List<ApprovedSpeakersGetPayload> eventRequestBrandsList = sheetData
                    .Where(row => row.TryGetValue("Is Active", out var isActive) && isActive?.ToString().ToLower().Equals("yes", StringComparison.OrdinalIgnoreCase) == true)
                    .Select(row => new ApprovedSpeakersGetPayload
                    {

                        SpeakerId = row.TryGetValue("SpeakerId", out var eventTopic) ? eventTopic?.ToString() : null,
                        SpeakerName = row.TryGetValue("SpeakerName", out var eventId) ? eventId?.ToString() : null,
                        SpeakerCode = row.TryGetValue("Speaker Code", out var eventType) ? eventType?.ToString() : null,
                        SpeakerCategory = row.TryGetValue("Speaker Category", out var eventDate) ? eventDate?.ToString() : null,
                        Speciality = row.TryGetValue("Speciality", out var startTime) ? startTime?.ToString() : null,
                        SpeakerType = row.TryGetValue("Speaker Type", out var endTime) ? endTime?.ToString() : null,
                        MisCode = row.TryGetValue("MisCode", out var venueName) ? venueName?.ToString() : null,
                        FCPASignOffDate = row.TryGetValue("FCPA Sign Off Date", out var eventEndDate) ? eventEndDate?.ToString() : null,
                        FCPAExpiryDate = row.TryGetValue("FCPA Expiry Date", out var FCPAExpiryDate) ? FCPAExpiryDate?.ToString() : null,
                        FCPAValid = row.TryGetValue("FCPA Valid?", out var state) ? state?.ToString() : null,
                        TRCDate = row.TryGetValue("TRC Date", out var city) ? city?.ToString() : null,
                        InitiatorName = row.TryGetValue("InitiatorName", out var initiatorName) ? initiatorName?.ToString() : null,
                        InitiatorEmail = row.TryGetValue("Initiator Email", out var initiatorEmail) ? initiatorEmail?.ToString() : null,
                        // CreatedDateHelper = row.TryGetValue("Created Date - Helper", out var createdDateHelper) ? createdDateHelper?.ToString() : null,
                        TRCValid = row.TryGetValue("TRC Valid?", out var isAdvanceRequired) ? isAdvanceRequired?.ToString() : null,
                        TotalSpendonSpeaker = row.TryGetValue("Total Spend on Speaker", out var prerBmBmApproval) ? prerBmBmApproval?.ToString() : null,
                        AggregateExhausted = row.TryGetValue("Aggregate Exhausted?", out var prerBmBmApprovalDate) ? prerBmBmApprovalDate?.ToString() : null,
                        Division = row.TryGetValue("Division", out var presalesHeadApproval) ? presalesHeadApproval?.ToString() : null,
                        Qualification = row.TryGetValue("Qualification", out var presalesHeadApprovalDate) ? presalesHeadApprovalDate?.ToString() : null,
                        Address = row.TryGetValue("Address", out var premarketingHeadApproval) ? premarketingHeadApproval?.ToString() : null,
                        State = row.TryGetValue("State", out var premarketingHeadApprovalDate) ? premarketingHeadApprovalDate?.ToString() : null,
                        City = row.TryGetValue("City", out var agreement) ? agreement?.ToString() : null,
                        Country = row.TryGetValue("Country", out var prefFinanceTreasuryApprovalDate) ? prefFinanceTreasuryApprovalDate?.ToString() : null,
                        ContactNumber = row.TryGetValue("Contact Number", out var premedicalAffairsHeadApproval) ? premedicalAffairsHeadApproval?.ToString() : null,
                        ABMNameRequestor = row.TryGetValue("ABM Name (Requestor)", out var premedicalAffairsHeadApprovalDate) ? premedicalAffairsHeadApprovalDate?.ToString() : null,
                        SpeakerCategoryId = row.TryGetValue("SpeakerCategoryId", out var presalesCoordinatorApproval) ? presalesCoordinatorApproval?.ToString() : null,
                        SpeakerCriteria = row.TryGetValue("Speaker Criteria", out var presalesCoordinatorApprovalDate) ? presalesCoordinatorApprovalDate?.ToString() : null,
                        IsActive = row.TryGetValue("Is Active", out var precomplianceApproval) ? precomplianceApproval?.ToString() : null,
                        CVdocument = row.TryGetValue("CVdocument", out var precomplianceApprovalDate) ? precomplianceApprovalDate?.ToString() : null,
                        Created = row.TryGetValue("Created", out var eventRequestStatus) ? eventRequestStatus?.ToString() : null,
                        CreatedDateHelper = row.TryGetValue("Created Date - Helper", out var eventApprovedDate) ? eventApprovedDate?.ToString() : null,
                        SalesAlertTrigger = row.TryGetValue("Sales Alert Trigger", out var honorariumRequestStatus) ? honorariumRequestStatus?.ToString() : null,
                        MedicalAffairsAlertTrigger = row.TryGetValue("Medical Affairs Alert Trigger", out var honorariumApprovedDate) ? honorariumApprovedDate?.ToString() : null,
                        SalesHeadApproval = row.TryGetValue("Sales Head Approval", out var postEventRequeststatus) ? postEventRequeststatus?.ToString() : null,
                        SalesHeadApprovalDate = row.TryGetValue("Sales Head Approval Date", out var helperFinancetreasurytriggerBTE) ? helperFinancetreasurytriggerBTE?.ToString() : null,
                        MedicalAffairsHeadApproval = row.TryGetValue("Medical Affairs Head Approval", out var postEventApprovedDate) ? postEventApprovedDate?.ToString() : null,
                        MedicalAffairsHeadApprovalDate = row.TryGetValue("Medical Affairs Head Approval Date", out var eventOpenSalesHeadApproval) ? eventOpenSalesHeadApproval?.ToString() : null,
                        SpeakerCriteriaDetails = row.TryGetValue("Speaker Criteria Details", out var eventOpenSalesHeadApprovalDate) ? eventOpenSalesHeadApprovalDate?.ToString() : null,
                        SalesHead = row.TryGetValue("Sales Head", out var _7daysSalesHeadApproval) ? _7daysSalesHeadApproval?.ToString() : null,
                        MedicalAffairsHead = row.TryGetValue("Medical Affairs Head", out var _7daysSalesHeadApprovaldate) ? _7daysSalesHeadApprovaldate?.ToString() : null,
                        NAID = row.TryGetValue("NA ID", out var prefBExpenseExcludingTaxApproval) ? prefBExpenseExcludingTaxApproval?.ToString() : null,
                        AggregateHonorariumLimit = row.TryGetValue("Aggregate Honorarium Limit", out var totalAttendees) && double.TryParse(totalAttendees?.ToString(), out var parsedTotalAttendees) ? parsedTotalAttendees : 0,
                        AggregateAccommodataionLimit = row.TryGetValue("Aggregate Accommodataion Limit", out var totalHonorariumAmount) && double.TryParse(totalHonorariumAmount?.ToString(), out var parsedTotalHonorariumAmount) ? parsedTotalHonorariumAmount : 0,
                        AggregateHonorariumSpent = row.TryGetValue("Aggregate Honorarium Spent", out var totalTravelAccommodationAmount) && double.TryParse(totalTravelAccommodationAmount?.ToString(), out var parsedTotalTravelAccommodationAmount) ? parsedTotalTravelAccommodationAmount : 0,
                        AggregateAccommodationSpent = row.TryGetValue("Aggregate Accommodation Spent", out var totalTravelAmount) && double.TryParse(totalTravelAmount?.ToString(), out var parsedTotalTravelAmount) ? parsedTotalTravelAmount : 0,
                        AggregateSpentonMedicalUtility = row.TryGetValue("Aggregate Spent on Medical Utility", out var totalAccommodationAmount) && double.TryParse(totalAccommodationAmount?.ToString(), out var parsedTotalAccommodationAmount) ? parsedTotalAccommodationAmount : 0,
                        AggregateSpentasHCPConsultant = row.TryGetValue("Aggregate Spent as HCP Consultant", out var totalLocalConveyance) && double.TryParse(totalLocalConveyance?.ToString(), out var parsedTotalLocalConveyance) ? parsedTotalLocalConveyance : 0,

                    }).ToList();
                return Ok(eventRequestBrandsList);
            }
            catch (Exception ex)
            {
                Log.Error($"Error occurred on AllPreEventsController AttachmentFile method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("GetAttachmentsFromApprovedSpeakerSheetBasedOnId")]
        public IActionResult GetAttachmentsFromApprovedSpeakerSheetBasedOnId(string Id)
        {
            string sheetId = configuration.GetSection("SmartsheetSettings:ApprovedSpeakers").Value;
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
            Dictionary<string, object> ProductBrandsListrowData = new Dictionary<string, object>();
            Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == Id));
            if (targetRow != null)
            {
                var attachments = SheetHelper.GetAttachmentsForRow(smartsheet, sheet.Id.Value, targetRow.Id.Value);
                ProductBrandsListrowData["Attachments"] = attachments;
            }
            return Ok(ProductBrandsListrowData);
        }
        [HttpGet("GetApprovedTrainersSheetData")]
        public IActionResult GetApprovedTrainersSheetData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:ApprovedTrainers").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                List<ApprovedTrainersGetPayload> eventRequestBrandsList = sheetData
                     .Where(row => row.TryGetValue("IsActive", out var isActive) && isActive?.ToString().ToLower().Equals("yes", StringComparison.OrdinalIgnoreCase) == true)
                    .Select(row => new ApprovedTrainersGetPayload
                    {

                        TrainerId = row.TryGetValue("TrainerId", out var eventTopic) ? eventTopic?.ToString() : null,
                        TrainerName = row.TryGetValue("TrainerName", out var eventId) ? eventId?.ToString() : null,
                        TrainerCode = row.TryGetValue("TrainerCode", out var eventType) ? eventType?.ToString() : null,
                        TrnrCode = row.TryGetValue("TrnrCode", out var eventDate) ? eventDate?.ToString() : null,
                        TierType = row.TryGetValue("TierType", out var startTime) ? startTime?.ToString() : null,
                        Speciality = row.TryGetValue("Speciality", out var endTime) ? endTime?.ToString() : null,
                        Is_NONGO = row.TryGetValue("Is_NONGO", out var venueName) ? venueName?.ToString() : null,
                        MisCode = row.TryGetValue("MisCode", out var city) ? city?.ToString() : null,
                        NAID = row.TryGetValue("NA ID", out var prefBExpenseExcludingTaxApproval) ? prefBExpenseExcludingTaxApproval?.ToString() : null,
                        AggregatespendonHonorariumTrainer = row.TryGetValue("Aggregate spend on Honorarium - Trainer", out var totalAttendees) && double.TryParse(totalAttendees?.ToString(), out var parsedTotalAttendees) ? parsedTotalAttendees : 0,
                        AggregateLimitonAccomodation = row.TryGetValue("Aggregate Limit on Accomodation", out var totalHonorariumAmount) && double.TryParse(totalHonorariumAmount?.ToString(), out var parsedTotalHonorariumAmount) ? parsedTotalHonorariumAmount : 0,
                        AggregateHonorariumSpent = row.TryGetValue("Aggregate Honorarium Spent ", out var totalTravelAccommodationAmount) && double.TryParse(totalTravelAccommodationAmount?.ToString(), out var parsedTotalTravelAccommodationAmount) ? parsedTotalTravelAccommodationAmount : 0,
                        AggregatespendonAccomodation = row.TryGetValue("Aggregate spend on Accomodation", out var totalTravelAmount) && double.TryParse(totalTravelAmount?.ToString(), out var parsedTotalTravelAmount) ? parsedTotalTravelAmount : 0,
                        AggregateSpentonMedicalUtility = row.TryGetValue("Aggregate Spent on Medical Utility", out var totalAccommodationAmount) && double.TryParse(totalAccommodationAmount?.ToString(), out var parsedTotalAccommodationAmount) ? parsedTotalAccommodationAmount : 0,
                        AggregateSpentasHCPConsultant = row.TryGetValue("Aggregate Spent as HCP Consultant", out var totalLocalConveyance) && double.TryParse(totalLocalConveyance?.ToString(), out var parsedTotalLocalConveyance) ? parsedTotalLocalConveyance : 0,

                        // CreatedDateHelper = row.TryGetValue("Created Date - Helper", out var createdDateHelper) ? createdDateHelper?.ToString() : null,
                        Qualification = row.TryGetValue("Qualification", out var isAdvanceRequired) ? isAdvanceRequired?.ToString() : null,
                        TrainerCategoryId = row.TryGetValue("TrainerCategoryId", out var prerBmBmApproval) ? prerBmBmApproval?.ToString() : null,
                        CityId = row.TryGetValue("CityId", out var presalesHeadApprovalDate) ? presalesHeadApprovalDate?.ToString() : null,
                        StateId = row.TryGetValue("StateId)", out var premedicalAffairsHeadApprovalDate) ? premedicalAffairsHeadApprovalDate?.ToString() : null,
                        Country = row.TryGetValue("Country", out var presalesCoordinatorApproval) ? presalesCoordinatorApproval?.ToString() : null,
                        IsActive = row.TryGetValue("IsActive", out var presalesCoordinatorApprovalDate) ? presalesCoordinatorApprovalDate?.ToString() : null,
                        FCPASignOffDate = row.TryGetValue("FCPA Sign Off Date", out var eventEndDate) ? eventEndDate?.ToString() : null,
                        FCPAExpiryDate = row.TryGetValue("FCPA Expiry Date", out var FCPAExpiryDate) ? FCPAExpiryDate?.ToString() : null,
                        FCPAValid = row.TryGetValue("FCPA Valid?", out var state) ? state?.ToString() : null,

                        //FCPASignOffDate = row.TryGetValue("FCPA Sign Off Date", out var precomplianceApproval) ? precomplianceApproval?.ToString() : null,
                        CV_Document = row.TryGetValue("CV_Document", out var precomplianceApprovalDate) ? precomplianceApprovalDate?.ToString() : null,
                        AggregateExhausted = row.TryGetValue("Aggregate Exhausted?", out var prerBmBmApprovalDate) ? prerBmBmApprovalDate?.ToString() : null,

                        Created = row.TryGetValue("Created", out var eventRequestStatus) ? eventRequestStatus?.ToString() : null,
                        CreatedDateHelper = row.TryGetValue("Created Date - Helper", out var eventApprovedDate) ? eventApprovedDate?.ToString() : null,
                        SalesAlertTrigger = row.TryGetValue("Sales Alert Trigger", out var honorariumRequestStatus) ? honorariumRequestStatus?.ToString() : null,
                        //MedicalAffairsAlertTrigger = row.TryGetValue("Medical Affairs Alert Trigger", out var honorariumApprovedDate) ? honorariumApprovedDate?.ToString() : null,
                        SalesHeadApproval = row.TryGetValue("Sales Head Approval", out var postEventRequeststatus) ? postEventRequeststatus?.ToString() : null,
                        SalesHeadApprovalDate = row.TryGetValue("Sales Head Approval Date", out var helperFinancetreasurytriggerBTE) ? helperFinancetreasurytriggerBTE?.ToString() : null,
                        MedicalAffairsHeadApproval = row.TryGetValue("Medical Affairs Head Approval", out var postEventApprovedDate) ? postEventApprovedDate?.ToString() : null,
                        MedicalAffairsHeadApprovalDate = row.TryGetValue("Medical Affairs Head Approval Date", out var eventOpenSalesHeadApproval) ? eventOpenSalesHeadApproval?.ToString() : null,
                        //SpeakerCriteriaDetails = row.TryGetValue("Speaker Criteria Details Date", out var eventOpenSalesHeadApprovalDate) ? eventOpenSalesHeadApprovalDate?.ToString() : null,
                        SalesHead = row.TryGetValue("Sales Head", out var _7daysSalesHeadApproval) ? _7daysSalesHeadApproval?.ToString() : null,
                        MedicalAffairsHead = row.TryGetValue("Medical Affairs Head", out var _7daysSalesHeadApprovaldate) ? _7daysSalesHeadApprovaldate?.ToString() : null,

                        InitiatorName = row.TryGetValue("InitiatorName", out var initiatorName) ? initiatorName?.ToString() : null,
                        InitiatorEmail = row.TryGetValue("Initiator Email", out var initiatorEmail) ? initiatorEmail?.ToString() : null,
                        TrainerTybeShortcode = row.TryGetValue("TrainerTybe - Shortcode", out var TrainerTybeShortcode) ? TrainerTybeShortcode?.ToString() : null,
                        TrainerCodeNew = row.TryGetValue("Trainer Code", out var TrainerCodeNew) ? TrainerCodeNew?.ToString() : null,
                        TrainerBrand = row.TryGetValue("Trainer Brand", out var TrainerBrand) ? TrainerBrand?.ToString() : null,
                        TrainerType = row.TryGetValue("Trainer Type", out var TrainerType) ? TrainerType?.ToString() : null,

                        Division = row.TryGetValue("Division", out var presalesHeadApproval) ? presalesHeadApproval?.ToString() : null,
                        Address = row.TryGetValue("Address", out var premarketingHeadApproval) ? premarketingHeadApproval?.ToString() : null,
                        State = row.TryGetValue("State", out var premarketingHeadApprovalDate) ? premarketingHeadApprovalDate?.ToString() : null,
                        City = row.TryGetValue("City", out var agreement) ? agreement?.ToString() : null,
                        //Country = row.TryGetValue("Country", out var prefFinanceTreasuryApprovalDate) ? prefFinanceTreasuryApprovalDate?.ToString() : null,
                        ContactNumber = row.TryGetValue("Contact Number", out var premedicalAffairsHeadApproval) ? premedicalAffairsHeadApproval?.ToString() : null,
                        Trainedby = row.TryGetValue("Trained by", out var Trainedby) ? Trainedby?.ToString() : null,
                        TrainerCV = row.TryGetValue("Trainer CV", out var TrainerCV) ? TrainerCV?.ToString() : null,
                        Trainercertificate = row.TryGetValue("Trainer certificate", out var Trainercertificate) ? Trainercertificate?.ToString() : null,
                        Trainedon = row.TryGetValue("Trained on", out var Trainedon) ? Trainedon?.ToString() : null,
                        TrainerCategory = row.TryGetValue("Trainer Category", out var TrainerCategory) ? TrainerCategory?.ToString() : null,
                        TrainerCriteria = row.TryGetValue("Trainer Criteria", out var TrainerCriteria) ? TrainerCriteria?.ToString() : null,
                        TrainerCriteriaDetails = row.TryGetValue("Trainer Criteria Details", out var TrainerCriteriaDetails) ? TrainerCriteriaDetails?.ToString() : null,
                        MedicalAffairsAlertTrigger = row.TryGetValue("Medical Affairs Alert Trigger", out var MedicalAffairsAlertTrigger) ? MedicalAffairsAlertTrigger?.ToString() : null,

                    }).ToList();
                return Ok(eventRequestBrandsList);
            }
            catch (Exception ex)
            {
                Log.Error($"Error occurred on AllPreEventsController AttachmentFile method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("GetAttachmentsFromApprovedTrainersSheetBasedOnId")]
        public IActionResult GetAttachmentsFromApprovedTrainersSheetBasedOnId(string Id)
        {
            string sheetId = configuration.GetSection("SmartsheetSettings:ApprovedTrainers").Value;
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
            Dictionary<string, object> ProductBrandsListrowData = new Dictionary<string, object>();
            Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == Id));
            if (targetRow != null)
            {
                var attachments = SheetHelper.GetAttachmentsForRow(smartsheet, sheet.Id.Value, targetRow.Id.Value);
                ProductBrandsListrowData["Attachments"] = attachments;
            }
            return Ok(ProductBrandsListrowData);
        }
        [HttpGet("GetApprovedVendorSheetData")]
        public IActionResult GetApprovedVendorSheetData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:VendorMasterSheet").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                List<ApprovedVendorsGetPayload> eventRequestBrandsList = sheetData
                    .Where(row => row.TryGetValue("IsActive?", out var isActive) && isActive?.ToString().ToLower().Equals("yes", StringComparison.OrdinalIgnoreCase) == true)
                    .Select(row => new ApprovedVendorsGetPayload
                    {
                        VendorId = row.TryGetValue("VendorId", out var eventTopic) ? eventTopic?.ToString() : null,
                        VendorAccount = row.TryGetValue("VendorAccount", out var eventId) ? eventId?.ToString() : null,
                        MisCode = row.TryGetValue("MisCode", out var city) ? city?.ToString() : null,
                        RequestorName = row.TryGetValue("Requestor Name", out var eventType) ? eventType?.ToString() : null,
                        BankName = row.TryGetValue("Bank Name", out var eventDate) ? eventDate?.ToString() : null,
                        BeneficiaryName = row.TryGetValue("BeneficiaryName", out var startTime) ? startTime?.ToString() : null,
                        PanCardName = row.TryGetValue("PanCardName", out var endTime) ? endTime?.ToString() : null,
                        PanNumber = row.TryGetValue("PanNumber", out var venueName) ? venueName?.ToString() : null,
                        BankAccountNumber = row.TryGetValue("BankAccountNumber", out var prefBExpenseExcludingTaxApproval) ? prefBExpenseExcludingTaxApproval?.ToString() : null,
                        PancardDocument = row.TryGetValue("Pancard Document", out var isAdvanceRequired) ? isAdvanceRequired?.ToString() : null,
                        Email = row.TryGetValue("Email ", out var prerBmBmApproval) ? prerBmBmApproval?.ToString() : null,
                        IBNNumber = row.TryGetValue("IBN Number", out var presalesHeadApprovalDate) ? presalesHeadApprovalDate?.ToString() : null,
                        //IfscCode = row.TryGetValue("IfscCode)", out var premedicalAffairsHeadApprovalDate) ? premedicalAffairsHeadApprovalDate?.ToString() : null,
                        IfscCode = row.TryGetValue("IfscCode", out var IfscCode) ? IfscCode?.ToString() : null,
                        SwiftCode = row.TryGetValue("Swift Code", out var presalesCoordinatorApproval) ? presalesCoordinatorApproval?.ToString() : null,
                        IsActive = row.TryGetValue("IsActive?", out var presalesCoordinatorApprovalDate) ? presalesCoordinatorApprovalDate?.ToString() : null,
                        ChequeDocument = row.TryGetValue("Cheque Document", out var eventEndDate) ? eventEndDate?.ToString() : null,
                        TaxResidenceCertificate = row.TryGetValue("Tax Residence Certificate", out var FCPAExpiryDate) ? FCPAExpiryDate?.ToString() : null,
                        TaxResidenceCertificateDate = row.TryGetValue("Tax Residence Certificate Date", out var state) ? state?.ToString() : null,
                        //FinanceChecker1 = row.TryGetValue("Finance Checker-1", out var precomplianceApprovalDate) ? precomplianceApprovalDate?.ToString() : null,
                        FinanceCheckerApproval = row.TryGetValue("Finance Checker Approval", out var prerBmBmApprovalDate) ? prerBmBmApprovalDate?.ToString() : null,
                        FinanceCheckerApprovalDate = row.TryGetValue("Finance Checker Approval Date", out var eventRequestStatus) ? eventRequestStatus?.ToString() : null,
                        Requestor = row.TryGetValue("Requestor", out var Requestor) ? Requestor?.ToString() : null,
                        InitiatorName = row.TryGetValue("Initiator Name", out var initiatorName) ? initiatorName?.ToString() : null,
                        InitiatorEmail = row.TryGetValue("Initiator Email", out var initiatorEmail) ? initiatorEmail?.ToString() : null,
                        FinanceChecker = row.TryGetValue("Finance Checker", out var eventApprovedDate) ? eventApprovedDate?.ToString() : null,

                    }).ToList();
                return Ok(eventRequestBrandsList);
            }
            catch (Exception ex)
            {
                Log.Error($"Error occurred on AllPreEventsController AttachmentFile method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("GetAttachmentsFromApprovedVendorSheetBasedOnId")]
        public IActionResult GetAttachmentsFromApprovedVendorSheetBasedOnId(string Id)
        {
            string sheetId = configuration.GetSection("SmartsheetSettings:VendorMasterSheet").Value;
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
            Dictionary<string, object> ProductBrandsListrowData = new Dictionary<string, object>();
            Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == Id));
            if (targetRow != null)
            {
                var attachments = SheetHelper.GetAttachmentsForRow(smartsheet, sheet.Id.Value, targetRow.Id.Value);
                ProductBrandsListrowData["Attachments"] = attachments;
            }
            return Ok(ProductBrandsListrowData);
        }
        [HttpGet("GetSpeakerCodeCreationSheetData")]
        public IActionResult GetSpeakerCodeCreationSheetData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:SpeakerCodeCreation").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                List<SpeakerCodeCreationGetPayload> eventRequestBrandsList = sheetData
                    .Select(row => new SpeakerCodeCreationGetPayload
                    {
                        SpeakerName = row.TryGetValue("SpeakerName", out var eventTopic) ? eventTopic?.ToString() : null,
                        Speaker_Code = row.TryGetValue("SpeakerCode", out var eventId) ? eventId?.ToString() : null,
                        SpeakerCode = row.TryGetValue("Speaker Code", out var eventType) ? eventType?.ToString() : null,
                        MisCode = row.TryGetValue("MisCode", out var venueName) ? venueName?.ToString() : null,
                        Division = row.TryGetValue("Division", out var presalesHeadApproval) ? presalesHeadApproval?.ToString() : null,
                        Speciality = row.TryGetValue("Speciality", out var startTime) ? startTime?.ToString() : null,
                        Qualification = row.TryGetValue("Qualification", out var presalesHeadApprovalDate) ? presalesHeadApprovalDate?.ToString() : null,
                        Address = row.TryGetValue("Address", out var premarketingHeadApproval) ? premarketingHeadApproval?.ToString() : null,
                        City = row.TryGetValue("City", out var agreement) ? agreement?.ToString() : null,
                        State = row.TryGetValue("State", out var premarketingHeadApprovalDate) ? premarketingHeadApprovalDate?.ToString() : null,
                        Country = row.TryGetValue("Country", out var prefFinanceTreasuryApprovalDate) ? prefFinanceTreasuryApprovalDate?.ToString() : null,
                        ContactNumber = row.TryGetValue("Contact Number", out var premedicalAffairsHeadApproval) ? premedicalAffairsHeadApproval?.ToString() : null,
                        SpeakerCategory = row.TryGetValue("Speaker Category", out var eventDate) ? eventDate?.ToString() : null,
                        SpeakerType = row.TryGetValue("Speaker Type", out var endTime) ? endTime?.ToString() : null,
                        SpeakerCriteria = row.TryGetValue("Speaker Criteria", out var SpeakerCriteria) ? SpeakerCriteria?.ToString() : null,

                        SpeakerCriteriaDetails = row.TryGetValue("Speaker Criteria Details", out var eventOpenSalesHeadApprovalDate) ? eventOpenSalesHeadApprovalDate?.ToString() : null,
                        CreatedDateHelper = row.TryGetValue("Created Date - Helper", out var eventApprovedDate) ? eventApprovedDate?.ToString() : null,
                        SalesAlertTrigger = row.TryGetValue("Sales Alert Trigger", out var eventEndDate) ? eventEndDate?.ToString() : null,
                        SalesHeadApproval = row.TryGetValue("Sales Head Approval", out var FCPAExpiryDate) ? FCPAExpiryDate?.ToString() : null,
                        SalesHeadApprovalDate = row.TryGetValue("Sales Head Approval Date", out var state) ? state?.ToString() : null,
                        MedicalAffairsAlertTrigger = row.TryGetValue("Medical Affairs Alert Trigger", out var city) ? city?.ToString() : null,
                        MedicalAffairsHeadApproval = row.TryGetValue("Medical Affairs Head Approval", out var isAdvanceRequired) ? isAdvanceRequired?.ToString() : null,
                        MedicalAffairsHeadApprovalDate = row.TryGetValue("Medical Affairs Head Approval Date", out var prerBmBmApproval) ? prerBmBmApproval?.ToString() : null,
                        SalesHead = row.TryGetValue("Sales Head", out var _7daysSalesHeadApproval) ? _7daysSalesHeadApproval?.ToString() : null,
                        MedicalAffairsHead = row.TryGetValue("Medical Affairs Head", out var _7daysSalesHeadApprovaldate) ? _7daysSalesHeadApprovaldate?.ToString() : null,
                        InitiatorName = row.TryGetValue("Initiator Name", out var initiatorName) ? initiatorName?.ToString() : null,
                        InitiatorEmail = row.TryGetValue("Initiator Email", out var initiatorEmail) ? initiatorEmail?.ToString() : null,
                        Created = row.TryGetValue("Created", out var eventRequestStatus) ? eventRequestStatus?.ToString() : null,

                    }).ToList();
                return Ok(eventRequestBrandsList);
            }
            catch (Exception ex)
            {
                Log.Error($"Error occurred on AllPreEventsController AttachmentFile method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("GetAttachmentsFromSpeakerCodeCreationCreationSheetBasedOnId")]
        public IActionResult GetAttachmentsFromSpeakerCodeCreationCreationSheetBasedOnId(string Id)
        {
            string sheetId = configuration.GetSection("SmartsheetSettings:SpeakerCodeCreation").Value;
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
            Dictionary<string, object> ProductBrandsListrowData = new Dictionary<string, object>();
            Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == Id));
            if (targetRow != null)
            {
                var attachments = SheetHelper.GetAttachmentsForRow(smartsheet, sheet.Id.Value, targetRow.Id.Value);
                ProductBrandsListrowData["Attachments"] = attachments;
            }
            return Ok(ProductBrandsListrowData);
        }
        [HttpGet("GetTrainersCodeCreationSheetData")]
        public IActionResult GetTrainersCodeCreationSheetData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:TrainerCodeCreation").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                List<TrainerCodeCreationGetPayload> eventRequestBrandsList = sheetData
                    .Select(row => new TrainerCodeCreationGetPayload
                    {
                        TrainerName = row.TryGetValue("TrainerName", out var eventId) ? eventId?.ToString() : null,

                        Number = row.TryGetValue("Number", out var eventTopic) ? eventTopic?.ToString() : null,
                        TrainerTybeShortcode = row.TryGetValue("TrainerTybe - Shortcode", out var TrainerTybe) ? TrainerTybe?.ToString() : null,

                        TrainerCode = row.TryGetValue("Trainer Code", out var eventType) ? eventType?.ToString() : null,

                        TrainerBrand = row.TryGetValue("Trainer Brand", out var eventDate) ? eventDate?.ToString() : null,
                        TrainerType = row.TryGetValue("Trainer Type", out var startTime) ? startTime?.ToString() : null,
                        MisCode = row.TryGetValue("MisCode", out var city) ? city?.ToString() : null,
                        Division = row.TryGetValue("Division", out var presalesHeadApproval) ? presalesHeadApproval?.ToString() : null,
                        Speciality = row.TryGetValue("Speciality", out var endTime) ? endTime?.ToString() : null,
                        Qualification = row.TryGetValue("Qualification", out var isAdvanceRequired) ? isAdvanceRequired?.ToString() : null,
                        Address = row.TryGetValue("Address", out var premarketingHeadApproval) ? premarketingHeadApproval?.ToString() : null,
                        City = row.TryGetValue("City", out var agreement) ? agreement?.ToString() : null,
                        State = row.TryGetValue("State", out var premarketingHeadApprovalDate) ? premarketingHeadApprovalDate?.ToString() : null,
                        Country = row.TryGetValue("Country", out var presalesCoordinatorApproval) ? presalesCoordinatorApproval?.ToString() : null,
                        ContactNumber = row.TryGetValue("Contact Number", out var premedicalAffairsHeadApproval) ? premedicalAffairsHeadApproval?.ToString() : null,
                        Trainedby = row.TryGetValue("Trained by", out var Trainedby) ? Trainedby?.ToString() : null,
                        TrainerCV = row.TryGetValue("Trainer CV", out var TrainerCV) ? TrainerCV?.ToString() : null,
                        Trainercertificate = row.TryGetValue("Trainer certificate", out var Trainercertificate) ? Trainercertificate?.ToString() : null,
                        Trainedon = row.TryGetValue("Trained on", out var Trainedon) ? Trainedon?.ToString() : null,
                        TrainerCategory = row.TryGetValue("Trainer Category", out var TrainerCategory) ? TrainerCategory?.ToString() : null,
                        TrainerCriteria = row.TryGetValue("Trainer Criteria", out var TrainerCriteria) ? TrainerCriteria?.ToString() : null,
                        TrainerCriteriaDetails = row.TryGetValue("Trainer Criteria Details", out var TrainerCriteriaDetails) ? TrainerCriteriaDetails?.ToString() : null,
                        Created = row.TryGetValue("Created", out var eventRequestStatus) ? eventRequestStatus?.ToString() : null,
                        CreatedDateHelper = row.TryGetValue("Created Date - Helper", out var eventApprovedDate) ? eventApprovedDate?.ToString() : null,
                        SalesAlertTrigger = row.TryGetValue("Sales Alert Trigger", out var honorariumRequestStatus) ? honorariumRequestStatus?.ToString() : null,
                        MedicalAffairsAlertTrigger = row.TryGetValue("Medical Affairs Alert Trigger", out var MedicalAffairsAlertTrigger) ? MedicalAffairsAlertTrigger?.ToString() : null,
                        SalesHead = row.TryGetValue("Sales Head", out var _7daysSalesHeadApproval) ? _7daysSalesHeadApproval?.ToString() : null,
                        SalesHeadApproval = row.TryGetValue("Sales Head Approval", out var postEventRequeststatus) ? postEventRequeststatus?.ToString() : null,
                        SalesHeadApprovalDate = row.TryGetValue("Sales Head Approval Date", out var helperFinancetreasurytriggerBTE) ? helperFinancetreasurytriggerBTE?.ToString() : null,
                        MedicalAffairsHead = row.TryGetValue("Medical Affairs Head", out var _7daysSalesHeadApprovaldate) ? _7daysSalesHeadApprovaldate?.ToString() : null,
                        MedicalAffairsHeadApproval = row.TryGetValue("Medical Affairs Head Approval", out var postEventApprovedDate) ? postEventApprovedDate?.ToString() : null,
                        MedicalAffairsHeadApprovalDate = row.TryGetValue("Medical Affairs Head Approval Date", out var eventOpenSalesHeadApproval) ? eventOpenSalesHeadApproval?.ToString() : null,
                        InitiatorName = row.TryGetValue("Initiator Name", out var initiatorName) ? initiatorName?.ToString() : null,
                        InitiatorEmail = row.TryGetValue("Initiator Email", out var initiatorEmail) ? initiatorEmail?.ToString() : null,

                    }).ToList();
                return Ok(eventRequestBrandsList);
            }
            catch (Exception ex)
            {
                Log.Error($"Error occurred on AllPreEventsController AttachmentFile method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("GetAttachmentsFromTrainerCodeCreationSheetBasedOnId")]
        public IActionResult GetAttachmentsFromTrainerCodeCreationSheetBasedOnId(string Id)
        {
            string sheetId = configuration.GetSection("SmartsheetSettings:TrainerCodeCreation").Value;
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
            Dictionary<string, object> ProductBrandsListrowData = new Dictionary<string, object>();
            Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == Id));
            if (targetRow != null)
            {
                var attachments = SheetHelper.GetAttachmentsForRow(smartsheet, sheet.Id.Value, targetRow.Id.Value);
                ProductBrandsListrowData["Attachments"] = attachments;
            }
            return Ok(ProductBrandsListrowData);
        }
        [HttpGet("GetVendorCodeCreationSheetData")]
        public IActionResult GetVendorCodeCreationSheetData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:VendorCodeCreation").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                List<VendorCodeCreationGetPayload> eventRequestBrandsList = sheetData
                    .Select(row => new VendorCodeCreationGetPayload
                    {
                        VendorId = row.TryGetValue("V Id", out var VendorId) ? VendorId?.ToString() : null,
                        VendorAccount = row.TryGetValue("VendorAccount", out var eventId) ? eventId?.ToString() : null,
                        MisCode = row.TryGetValue("MisCode", out var city) ? city?.ToString() : null,
                        BeneficiaryName = row.TryGetValue("BeneficiaryName", out var startTime) ? startTime?.ToString() : null,
                        PanCardName = row.TryGetValue("PanCardName", out var endTime) ? endTime?.ToString() : null,
                        PanNumber = row.TryGetValue("PanNumber", out var venueName) ? venueName?.ToString() : null,
                        BankAccountNumber = row.TryGetValue("BankAccountNumber", out var prefBExpenseExcludingTaxApproval) ? prefBExpenseExcludingTaxApproval?.ToString() : null,
                        BankName = row.TryGetValue("Bank Name", out var eventDate) ? eventDate?.ToString() : null,
                        IfscCode = row.TryGetValue("IfscCode", out var IfscCode) ? IfscCode?.ToString() : null,
                        SwiftCode = row.TryGetValue("Swift Code", out var presalesCoordinatorApproval) ? presalesCoordinatorApproval?.ToString() : null,
                        IBNNumber = row.TryGetValue("IBN Number", out var presalesHeadApprovalDate) ? presalesHeadApprovalDate?.ToString() : null,
                        Email = row.TryGetValue("Email ", out var prerBmBmApproval) ? prerBmBmApproval?.ToString() : null,
                        PancardDocument = row.TryGetValue("Pancard Document", out var isAdvanceRequired) ? isAdvanceRequired?.ToString() : null,
                        ChequeDocument = row.TryGetValue("Cheque Document", out var eventEndDate) ? eventEndDate?.ToString() : null,
                        TaxResidenceCertificate = row.TryGetValue("Tax Residence Certificate", out var FCPAExpiryDate) ? FCPAExpiryDate?.ToString() : null,
                        InitiatorName = row.TryGetValue("Initiator Name", out var initiatorName) ? initiatorName?.ToString() : null,
                        InitiatorEmail = row.TryGetValue("Initiator Email", out var initiatorEmail) ? initiatorEmail?.ToString() : null,
                        TaxResidenceCertificateDate = row.TryGetValue("Tax Residence Certificate Date", out var state) ? state?.ToString() : null,
                        FinanceChecker = row.TryGetValue("Finance Checker", out var eventApprovedDate) ? eventApprovedDate?.ToString() : null,
                        FinanceCheckerapproval = row.TryGetValue("Finance Checker  approval", out var prerBmBmApprovalDate) ? prerBmBmApprovalDate?.ToString() : null,
                        FinanceCheckerApprovalDate = row.TryGetValue("Finance Checker Approval Date", out var eventRequestStatus) ? eventRequestStatus?.ToString() : null,

                    }).ToList();
                return Ok(eventRequestBrandsList);
            }
            catch (Exception ex)
            {
                Log.Error($"Error occurred on AllPreEventsController AttachmentFile method {ex.Message} at {DateTime.Now}");
                Log.Error(ex.StackTrace);
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("GetAttachmentsFromVendorCodeCreationSheetBasedOnId")]
        public IActionResult GetAttachmentsFromVendorCodeCreationSheetBasedOnId(string Id)
        {
            string sheetId = configuration.GetSection("SmartsheetSettings:VendorCodeCreation").Value;
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
            Dictionary<string, object> ProductBrandsListrowData = new Dictionary<string, object>();
            Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == Id));
            if (targetRow != null)
            {
                var attachments = SheetHelper.GetAttachmentsForRow(smartsheet, sheet.Id.Value, targetRow.Id.Value);
                ProductBrandsListrowData["Attachments"] = attachments;
            }
            return Ok(ProductBrandsListrowData);
        }




    }
}





















































