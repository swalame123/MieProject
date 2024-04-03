using IndiaEventsWebApi.Helper;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;

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

        public GetMasterSheetsController(IConfiguration configuration)
        {
            this.configuration = configuration;
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
        [HttpGet("HcpMaster")]

        public IActionResult HcpMaster()
        {
            try
            {
                List<Dictionary<string, object>> sheetData = new List<Dictionary<string, object>>();
                string sheetId1 = configuration.GetSection("SmartsheetSettings:HcpMaster1").Value;
                string sheetId2 = configuration.GetSection("SmartsheetSettings:HcpMaster2").Value;
                string sheetId3 = configuration.GetSection("SmartsheetSettings:HcpMaster3").Value;
                string sheetId4 = configuration.GetSection("SmartsheetSettings:HcpMaster4").Value;

                List<string> Sheets = new List<string>() { sheetId1, sheetId2, sheetId3, sheetId4 };
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



    }
}





















































