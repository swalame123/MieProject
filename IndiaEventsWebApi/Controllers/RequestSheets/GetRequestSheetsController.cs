using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using IndiaEventsWebApi.Junk.Test;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using Microsoft.AspNetCore.Authorization;
using IndiaEventsWebApi.Helper;

namespace IndiaEventsWebApi.Controllers.RequestSheets
{

    [Route("api/[controller]")]
    [ApiController]
    //[Authorize]
    public class GetRequestSheetsController : ControllerBase
    {

        private readonly string accessToken;
        private readonly IConfiguration configuration;
        private readonly SmartsheetClient smartsheet;

        public GetRequestSheetsController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
            smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();


        }

        [HttpGet("GetEventRequestWebData")]
        public IActionResult GetEventRequestWebData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:Class1").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpGet("GetEventRequestProcessData")]
        public IActionResult GetEventRequestProcessData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpGet("GetHonorariumPaymentData")]
        public IActionResult GetHonorariumPaymentData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:HonorariumPayment").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpGet("GetEventSettlementData")]
        public IActionResult GetEventSettlementData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:EventSettlement").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpGet("GetHCPSlideKitDetailsData")]
        public IActionResult GetHCPSlideKitDetailsData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsHcpSlideKit").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpGet("GetExpenseData")]
        public IActionResult GetExpenseData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        [HttpGet("GetHCPRoleDetailsData")]
        public IActionResult GetHCPRoleDetailsData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpGet("GetInviteesData")]
        public IActionResult GetInviteesData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpGet("GetEventRequestBrandsListData")]
        public IActionResult GetEventRequestBrandsListData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestBrandsList").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
                return Ok(sheetData);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }


        [HttpPost("GetEventRequestsHcpRoleByIds")]
        public IActionResult GetEventRequestsHcpRoleByIds([FromBody] getIds eventIdorEventRequestIds)
        {

            string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
            Sheet sheet1 = SheetHelper.GetSheetById(smartsheet, sheetId);
            List<Dictionary<string, object>> hcpRoleData = new List<Dictionary<string, object>>();
            List<string> columnNames = new List<string>();
            foreach (Column column in sheet1.Columns)
            {
                columnNames.Add(column.Title);
            }
            foreach (var val in eventIdorEventRequestIds.EventIds)
            {
                foreach (Row row in sheet1.Rows)
                {
                    var eventIdorEventRequestIdCell = row.Cells.FirstOrDefault(cell => cell.ColumnId == SheetHelper.GetColumnIdByName(sheet1, "EventId/EventRequestId"));
                    var x = eventIdorEventRequestIdCell.Value.ToString();
                    if (eventIdorEventRequestIdCell != null && x == val)
                    {
                        Dictionary<string, object> hcpRoleRowData = new Dictionary<string, object>();

                        for (int i = 0; i < row.Cells.Count && i < columnNames.Count; i++)
                        {
                            hcpRoleRowData[columnNames[i]] = row.Cells[i].Value;
                        }

                        hcpRoleData.Add(hcpRoleRowData);
                    }

                }
            }

            return Ok(hcpRoleData);

        }






        [HttpGet("GetEventRequestsHcpDetailsTotalSpendValue")]
        public IActionResult GetcgtColumnValue(string EventID)
        {
            string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;

            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);

            Column SpecialityColumn = sheet.Columns.FirstOrDefault(column => string.Equals(column.Title, "EventId/EventRequestId", StringComparison.OrdinalIgnoreCase));
            Column targetColumn = sheet.Columns.FirstOrDefault(column => string.Equals(column.Title, "TotalSpend", StringComparison.OrdinalIgnoreCase));

            if (targetColumn != null && SpecialityColumn != null)
            {
                List<Row> targetRows = sheet.Rows
                    .Where(row => row.Cells.Any(cell => cell.ColumnId == SpecialityColumn.Id && cell.Value.ToString() == EventID))
                    .ToList();

                if (targetRows.Any())
                {
                    decimal totalValue = targetRows.Sum(row => Convert.ToDecimal(row.Cells.FirstOrDefault(cell => cell.ColumnId == targetColumn.Id)?.Value ?? 0));
                    return Ok(totalValue);
                }
                else
                {
                    return NotFound("NotFound");
                }
            }
            else
            {
                return NotFound("NotFound");
            }
        }
        [HttpGet("GetEventRequestsInviteesLcAmountValue")]
        public IActionResult GetEventRequestsInviteesLcAmountValue(string EventID)
        {
            string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;

            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);

            Column SpecialityColumn = sheet.Columns.FirstOrDefault(column => string.Equals(column.Title, "EventId/EventRequestId", StringComparison.OrdinalIgnoreCase));
            Column targetColumn = sheet.Columns.FirstOrDefault(column => string.Equals(column.Title, "LcAmount", StringComparison.OrdinalIgnoreCase));

            if (targetColumn != null && SpecialityColumn != null)
            {

                List<Row> targetRows = sheet.Rows
                    .Where(row => row.Cells.Any(cell => cell.ColumnId == SpecialityColumn.Id && cell.Value.ToString() == EventID))
                    .ToList();

                if (targetRows.Any())
                {

                    decimal totalValue = targetRows.Sum(row => Convert.ToDecimal(row.Cells.FirstOrDefault(cell => cell.ColumnId == targetColumn.Id)?.Value ?? 0));

                    return Ok(totalValue);
                }
                else
                {
                    return NotFound("NotFound");
                }
            }
            else
            {
                return NotFound("NotFound");
            }
        }



        [HttpGet("GetEventRequestExpenseSheetAmountValue")]
        public IActionResult GetEventRequestExpenseSheetAmountValue(string EventID)
        {
            string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;

            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);

            Column SpecialityColumn = sheet.Columns.FirstOrDefault(column => string.Equals(column.Title, "EventId/EventRequestID", StringComparison.OrdinalIgnoreCase));
            Column targetColumn = sheet.Columns.FirstOrDefault(column => string.Equals(column.Title, "Amount", StringComparison.OrdinalIgnoreCase));

            if (targetColumn != null && SpecialityColumn != null)
            {
                List<Row> targetRows = sheet.Rows.Where(row => row.Cells?.Any(cell => cell.ColumnId == SpecialityColumn.Id && cell.Value?.ToString() == EventID) == true).ToList();

                if (targetRows.Any() == true)
                {

                    decimal totalValue = targetRows.Sum(row => Convert.ToDecimal(row.Cells?.FirstOrDefault(cell => cell.ColumnId == targetColumn.Id)?.Value ?? 0));

                    return Ok(totalValue);
                }
                else
                {
                    return NotFound("NotFound");
                }
            }
            else
            {
                return NotFound("NotFound");
            }
        }
    }
}
