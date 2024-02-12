using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using IndiaEventsWebApi.Junk.Test;
using Smartsheet.Api;
using Smartsheet.Api.Models;

namespace IndiaEventsWebApi.Controllers.RequestSheets
{
    [Route("api/[controller]")]
    [ApiController]
    public class GetRequestSheetsController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;

        public GetRequestSheetsController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

        }
        [HttpGet("GetEventRequestWebData")]
        public IActionResult GetEventRequestWebData()
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:Class1").Value;
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

        [HttpGet("GetEventRequestProcessData")]
        public IActionResult GetEventRequestProcessData()
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;
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

        [HttpGet("GetHonorariumPaymentData")]
        public IActionResult GetHonorariumPaymentData()
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:HonorariumPayment").Value;
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

        [HttpGet("GetEventSettlementData")]
        public IActionResult GetEventSettlementData()
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:EventSettlement").Value;
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

        [HttpGet("GetHCPSlideKitDetailsData")]
        public IActionResult GetHCPSlideKitDetailsData()
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsHcpSlideKit").Value;
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

        [HttpGet("GetExpenseData")]
        public IActionResult GetExpenseData()
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
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
        [HttpGet("GetHCPRoleDetailsData")]
        public IActionResult GetHCPRoleDetailsData()
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
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

        [HttpGet("GetInviteesData")]
        public IActionResult GetInviteesData()
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
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

        [HttpGet("GetEventRequestBrandsListData")]
        public IActionResult GetEventRequestBrandsListData()
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestBrandsList").Value;
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

        //[HttpGet("GetEventRequestsHcpRole")]
        //public IActionResult GetEventRequestsHcpRole()
        //{
        //    try
        //    {
        //        SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

        //        string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
        //        long.TryParse(sheetId, out long parsedSheetId);
        //        Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
        //        List<Dictionary<string, object>> hcpRoleData = new List<Dictionary<string, object>>();
        //        List<string> columnNames = new List<string>();

        //        foreach (Column column in sheet1.Columns)
        //        {
        //            columnNames.Add(column.Title);
        //        }
        //        foreach (Row row in sheet1.Rows)
        //        {

        //            // Check if the row has the specified EventIdorEventRequestId
        //            var eventIdorEventRequestIdCell = row.Cells.FirstOrDefault(cell => cell.ColumnId == GetColumnIdByName(sheet1, "EventId/EventRequestId"));
        //            var x = eventIdorEventRequestIdCell.Value.ToString();
        //            // Console.WriteLine($"Row {row.RowNumber}: EventId/EventRequestId in cell: {eventIdorEventRequestIdCell.Value}");
        //            if (eventIdorEventRequestIdCell != null)
        //            {
        //                Dictionary<string, object> hcpRoleRowData = new Dictionary<string, object>();

        //                for (int i = 0; i < row.Cells.Count && i < columnNames.Count; i++)
        //                {
        //                    hcpRoleRowData[columnNames[i]] = row.Cells[i].Value;
        //                }

        //                hcpRoleData.Add(hcpRoleRowData);
        //            }

        //        }

        //        return Ok(hcpRoleData);

        //    }
        //    catch (Exception ex)
        //    {
        //        return BadRequest(ex.Message);
        //    }
        //}

        [HttpPost("GetEventRequestsHcpRoleByIds")]
        public IActionResult GetEventRequestsHcpRoleByIds([FromBody] getIds eventIdorEventRequestIds)
        {
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
            long.TryParse(sheetId, out long parsedSheetId);
            Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
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
                    var eventIdorEventRequestIdCell = row.Cells.FirstOrDefault(cell => cell.ColumnId == GetColumnIdByName(sheet1, "EventId/EventRequestId"));
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
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
            long.TryParse(sheetId, out long parsedSheetId);

            Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);

            Column SpecialityColumn = sheet.Columns.FirstOrDefault(column => string.Equals(column.Title, "EventId/EventRequestId", StringComparison.OrdinalIgnoreCase));
            Column targetColumn = sheet.Columns.FirstOrDefault(column => string.Equals(column.Title, "TotalSpend", StringComparison.OrdinalIgnoreCase));
            //Column targetColumn = sheet.Columns.FirstOrDefault(column => string.Equals(column.Title, "PresentationDuration", StringComparison.OrdinalIgnoreCase));

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
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
            long.TryParse(sheetId, out long parsedSheetId);

            Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);

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
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
            long.TryParse(sheetId, out long parsedSheetId);

            Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);

            Column SpecialityColumn = sheet.Columns.FirstOrDefault(column => string.Equals(column.Title, "EventId/EventRequestID", StringComparison.OrdinalIgnoreCase));
            Column targetColumn = sheet.Columns.FirstOrDefault(column => string.Equals(column.Title, "Amount", StringComparison.OrdinalIgnoreCase));

            if (targetColumn != null && SpecialityColumn != null)
            {
               
                List<Row> targetRows = sheet.Rows
                    .Where(row => row.Cells?.Any(cell => cell.ColumnId == SpecialityColumn.Id && cell.Value?.ToString() == EventID)==true)
                    .ToList();

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
