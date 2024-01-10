using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using MieProject.Models;
using MieProject.Models.EventTypeSheets;
using NPOI.OpenXml4Net.OPC;
using Smartsheet.Api;
using Smartsheet.Api.Models;

namespace MieProject.Controllers.EventTypeSheetsControllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class Class1Controller : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;
        public Class1Controller(IConfiguration configuration)
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
        //private bool EmailExists(Sheet sheet, string email)
        //{
        //    long emailColumnId = GetColumnIdByName(sheet, "Email");

        //    return sheet.Rows.Any(row =>
        //        row.Cells.Any(cell => cell.ColumnId == emailColumnId && cell.Value?.ToString() == email));
        //}
        [HttpPost("AddData")]
        public IActionResult AddData([FromBody] Class1 formData)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

                string sheetId = configuration.GetSection("SmartsheetSettings:Class1").Value;

                long.TryParse(sheetId, out long parsedSheetId);
                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
                //if (EmailExists(sheet, formData.Email))
                //{
                //    return BadRequest("Email already exists in the sheet.");
                //}
                var newRow = new Row();
                newRow.Cells = new List<Cell>();
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "EventTopic"),
                    Value = formData.EventTopic
                });
                //newRow.Cells.Add(new Cell
                //{
                //    ColumnId = GetColumnIdByName(sheet, "EventId/EventRequestId"),
                //    Value = formData.EventId
                //});
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "EventType"),
                    Value = formData.EventType
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "EventDate"),
                    Value = formData.EventDate
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "StartTime"),
                    Value = formData.StartTime
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "EndTime"),
                    Value = formData.EndTime
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "VenueName"),
                    Value = formData.VenueName
                });
                //newRow.Cells.Add(new Cell
                //{
                //    ColumnId = GetColumnIdByName(sheet, "InitiatorName"),
                //    Value = formData.InitiatorName
                //});
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "City"),
                    Value = formData.City
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "State"),
                    Value = formData.State
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "BrandName"),
                    Value = formData.BrandName
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "% Allocation"),
                    Value = formData.PercentAllocation
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "ProjectID"),
                    Value = formData.ProjectId
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "HCPRole"),
                    Value = formData.HCPRole
                });
                //newRow.Cells.Add(new Cell
                //{
                //    ColumnId = GetColumnIdByName(sheet, "Objective"),
                //    Value = formData.Objective
                //});
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "MeetingTypeId"),
                    Value = formData.MeetingTypeId
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Honorarium"),
                    Value = formData.Honorarium
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "IsAdvanceRequired"),
                    Value = formData.IsAdvanceRequired
                });



                smartsheet.SheetResources.RowResources.AddRows(parsedSheetId, new Row[] { newRow });
                //return Ok("Data added successfully.");
                return Ok(new
                { Message = "Data added successfully." });

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
