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
        [HttpPost("AddData")]
        public IActionResult AddData([FromBody] EventRequestBrandsList formData)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestBrandsList").Value;

                long.TryParse(sheetId, out long parsedSheetId);
                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
                
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
