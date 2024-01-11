using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using MieProject.Models;
using MieProject.Models.RequestSheets;
using Smartsheet.Api;
using Smartsheet.Api.Models;

namespace MieProject.Controllers.RequestSheets
{
    [Route("api/[controller]")]
    [ApiController]
    public class EventRequestBrandsListController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;

        public EventRequestBrandsListController(IConfiguration configuration)
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
        [HttpPost("AddDataList")]
        public IActionResult AddDataList([FromBody] List<EventRequestBrandsList> formDataList)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestBrandsList").Value;
                long.TryParse(sheetId, out long parsedSheetId);

                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
                foreach (var formData in formDataList)
                {

                    var newRow = new Row();
                    newRow.Cells = new List<Cell>();
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "% Allocation"),
                        Value = formData.PercentAllocation
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "Brands"),
                        Value = formData.BrandName
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "Project ID"),
                        Value = formData.ProjectId
                    });



                    smartsheet.SheetResources.RowResources.AddRows(parsedSheetId, new Row[] { newRow });
                    //// Create a list to hold the rows to be added
                    //List<Row> rowsToAdd = new List<Row>();

                    //// Iterate through the provided list and create rows

                    //var newRow = new Row();
                    //newRow.Cells = new List<Cell>();
                    //newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet, "% Allocation"), Value = formData.PercentAllocation });
                    //newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet, "Brands"), Value = formData.BrandName });
                    //newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet, "Project ID"), Value = formData.ProjectId });
                    //rowsToAdd.Add(newRow);


                    //// Add the list of rows to the sheet
                    ////smartsheet.SheetResources.RowResources.AddRows(parsedSheetId, rowsToAdd.ToArray());
                    //smartsheet.SheetResources.RowResources.AddRows(parsedSheetId, new Row[] { newRow });
                }

                return Ok(new { Message = "Data added successfully." });
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

















        //[HttpPost("AddData")]
        //public IActionResult AddData([FromBody] EventRequestBrandsList formData)
        //{
        //    try
        //    {
        //        SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

        //        string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestBrandsList").Value;

        //        long.TryParse(sheetId, out long parsedSheetId);
        //        Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);

        //        var newRow = new Row();
        //        newRow.Cells = new List<Cell>();
        //        newRow.Cells.Add(new Cell
        //        {
        //            ColumnId = GetColumnIdByName(sheet, "% Allocation"),
        //            Value = formData.PercentAllocation
        //        });
        //        newRow.Cells.Add(new Cell
        //        {
        //            ColumnId = GetColumnIdByName(sheet, "Brands"),
        //            Value = formData.BrandName
        //        });
        //        newRow.Cells.Add(new Cell
        //        {
        //            ColumnId = GetColumnIdByName(sheet, "Project ID"),
        //            Value = formData.ProjectId
        //        });



        //        smartsheet.SheetResources.RowResources.AddRows(parsedSheetId, new Row[] { newRow });
        //        //return Ok("Data added successfully.");
        //        return Ok(new
        //        { Message = "Data added successfully." });

        //    }
        //    catch (Exception ex)
        //    {
        //        return BadRequest(ex.Message);
        //    }
        //}
        //private long GetColumnIdByName(Sheet sheet, string columnname)
        //{
        //    foreach (var column in sheet.Columns)
        //    {
        //        if (column.Title == columnname)
        //        {
        //            return column.Id.Value;
        //        }
        //    }
        //    return 0;
        //}


    }
}
