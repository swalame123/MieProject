using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using MieProject.Models.EventTypeSheets;
using Smartsheet.Api;
using Smartsheet.Api.Models;

namespace MieProject.Controllers.EventTypeSheetsControllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class HonorariumPaymentController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;
        public HonorariumPaymentController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
        }
       
        //[HttpPost("AddData")]
        //public IActionResult AddData(HonorariumPaymentList formData)
        //{
        //    try
        //    {
        //        SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

        //        string sheetId = configuration.GetSection("SmartsheetSettings:HonorariumPayment").Value;

        //        long.TryParse(sheetId, out long parsedSheetId);
        //        Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
        //        foreach(var i in formData.RequestHonorariumList)
        //        {
        //            var newRow = new Row();
        //            newRow.Cells = new List<Cell>();
        //            newRow.Cells.Add(new Cell
        //            {
        //                ColumnId = GetColumnIdByName(sheet, "HCP Name"),
        //                Value = i.HCPName
        //            });

        //            newRow.Cells.Add(new Cell
        //            {
        //                ColumnId = GetColumnIdByName(sheet, "EventId/EventRequestId"),
        //                Value = i.EventId
        //            });
        //            newRow.Cells.Add(new Cell
        //            {
        //                ColumnId = GetColumnIdByName(sheet, "EventType"),
        //                Value = i.EventType
        //            });
        //            newRow.Cells.Add(new Cell
        //            {
        //                ColumnId = GetColumnIdByName(sheet, "HCPRole"),
        //                Value = i.HCPRole
        //            });
        //            newRow.Cells.Add(new Cell
        //            {
        //                ColumnId = GetColumnIdByName(sheet, "MISCODE"),
        //                Value = i.MISCode
        //            });
        //            newRow.Cells.Add(new Cell
        //            {
        //                ColumnId = GetColumnIdByName(sheet, "GO/Non-GO"),
        //                Value = i.GONGO
        //            });
        //            newRow.Cells.Add(new Cell
        //            {
        //                ColumnId = GetColumnIdByName(sheet, "IsItincludingGST?"),
        //                Value = i.IsItincludingGST
        //            });
        //            newRow.Cells.Add(new Cell
        //            {
        //                ColumnId = GetColumnIdByName(sheet, "AgreementAmount"),
        //                Value = i.AgreementAmount
        //            });



        //            smartsheet.SheetResources.RowResources.AddRows(parsedSheetId, new Row[] { newRow });
        //        }


        //        return Ok(new
        //        { Message = "Data added successfully." });

        //    }
        //    catch (Exception ex)
        //    {
        //        return BadRequest(ex.Message);
        //    }
        //}
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
