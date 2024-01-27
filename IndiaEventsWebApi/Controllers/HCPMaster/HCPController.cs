using IndiaEventsWebApi.Models;
using IndiaEventsWebApi.Models.MasterSheets;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;

namespace IndiaEventsWebApi.Controllers.HCPMaster
{
    [Route("api/[controller]")]
    [ApiController]
    public class HCPController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;

        public HCPController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

        }
        [HttpPost("PostHcpData1")]
        public IActionResult PostHcpData1(HCPMaster1 formDataList)
        {
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            string[] sheetIds = {
                //configuration.GetSection("SmartsheetSettings:HcpMaster").Value,
                configuration.GetSection("SmartsheetSettings:HcpMaster1").Value,
                configuration.GetSection("SmartsheetSettings:HcpMaster2").Value,
                configuration.GetSection("SmartsheetSettings:HcpMaster3").Value,
                configuration.GetSection("SmartsheetSettings:HcpMaster4").Value
            };
            foreach (string i in sheetIds)
            {
                long.TryParse(i, out long p);
                Sheet sheeti = smartsheet.SheetResources.GetSheet(p, null, null, null, null, null, null, null);

                // Check if any row contains the same MISCode
                //Row existingRow = sheeti.Rows.FirstOrDefault(row => row.Cells.Any(cell => cell.Value.ToString() == formDataList.MISCode));
                Row existingRow = sheeti.Rows.FirstOrDefault(row =>
                    row.Cells != null &&
                    row.Cells.Any(cell => cell.Value != null && cell.Value.ToString() == formDataList.MISCode));

                if (existingRow != null)
                {
                    // Data with the same MISCode already exists, return a response
                    return BadRequest("Data with the same MISCode already exists.");
                }
            }
            string sheetId = configuration.GetSection("SmartsheetSettings:HcpMaster1").Value;
            long.TryParse(sheetId, out long parsedSheetId);
            Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
            var newRow = new Row();
            newRow.Cells = new List<Cell>();
            newRow.Cells.Add(new Cell
            {
                ColumnId = GetColumnIdByName(sheet, "FirstName"),
                Value = formDataList.FirstName
            });
            newRow.Cells.Add(new Cell
            {
                ColumnId = GetColumnIdByName(sheet, "LastName"),
                Value = formDataList.LastName
            });
            newRow.Cells.Add(new Cell
            {
                ColumnId = GetColumnIdByName(sheet, "HCPName"),
                Value = formDataList.HCPName
            });
            newRow.Cells.Add(new Cell
            {
                ColumnId = GetColumnIdByName(sheet, "GO/Non-GO"),
                Value = formDataList.GOorNGO
            });
            newRow.Cells.Add(new Cell
            {
                ColumnId = GetColumnIdByName(sheet, "MISCode"),
                Value = formDataList.MISCode
            });

            smartsheet.SheetResources.RowResources.AddRows(parsedSheetId, new Row[] { newRow });




            return Ok();

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

        //[HttpPost("PostHcpData")]
        //public IActionResult PostHcpData(AllObjModels formDataList)
        //{

        //    SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
        //    string sheetId1 = configuration.GetSection("SmartsheetSettings:HcpMaster").Value;
        //    string sheetId2 = configuration.GetSection("SmartsheetSettings:HcpMaster1").Value;
        //    string sheetId3 = configuration.GetSection("SmartsheetSettings:HcpMaster2").Value;
        //    string sheetId4 = configuration.GetSection("SmartsheetSettings:HcpMaster3").Value;
        //    string sheetId5 = configuration.GetSection("SmartsheetSettings:HcpMaster4").Value;

        //    long.TryParse(sheetId1, out long parsedSheetId1);
        //    long.TryParse(sheetId2, out long parsedSheetId2);
        //    long.TryParse(sheetId3, out long parsedSheetId3);
        //    long.TryParse(sheetId4, out long parsedSheetId4);
        //    long.TryParse(sheetId5, out long parsedSheetId5);

        //    Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId1, null, null, null, null, null, null, null);
        //    Sheet sheet2 = smartsheet.SheetResources.GetSheet(parsedSheetId2, null, null, null, null, null, null, null);
        //    Sheet sheet3 = smartsheet.SheetResources.GetSheet(parsedSheetId3, null, null, null, null, null, null, null);
        //    Sheet sheet4 = smartsheet.SheetResources.GetSheet(parsedSheetId4, null, null, null, null, null, null, null);
        //    Sheet sheet5 = smartsheet.SheetResources.GetSheet(parsedSheetId5, null, null, null, null, null, null, null);


        //    List<string> Sheets = new List<string>() { sheetId1, sheetId2 , sheetId3 , sheetId4, sheetId5 };



        //    return Ok(sheet1);
        //}


    //    [HttpPost("PostHcpData")]
    //    public IActionResult PostHcpData(AllObjModels formDataList)
    //    {
    //        SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

    //        string[] sheetIds = {
    //    configuration.GetSection("SmartsheetSettings:HcpMaster").Value,
    //    configuration.GetSection("SmartsheetSettings:HcpMaster1").Value,
    //    configuration.GetSection("SmartsheetSettings:HcpMaster2").Value,
    //    configuration.GetSection("SmartsheetSettings:HcpMaster3").Value,
    //    configuration.GetSection("SmartsheetSettings:HcpMaster4").Value
    //};

    //        foreach (string sheetId in sheetIds)
    //        {
    //            long.TryParse(sheetId, out long parsedSheetId);
    //            Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);

    //            // Check if any row contains the same MISCode
    //            Row existingRow = sheet.Rows.FirstOrDefault(row => row.Cells.Any(cell => cell.Value.ToString() == formDataList.MISCode));

    //            if (existingRow != null)
    //            {
    //                // Data with the same MISCode already exists, return a response
    //                return BadRequest("Data with the same MISCode already exists.");
    //            }
    //        }

    //        // If no matching MISCode found in any sheet, proceed to add the data
    //        Row rowToAdd = new Row.AddRowBuilder(true, null, null)
    //            .SetCells(new Cell[] {
    //        new Cell.AddCellBuilder(
    //            sheet.Columns.First(col => col.Title == "LastName").Id,
    //            formDataList.LastName
    //        ).Build(),
    //        new Cell.AddCellBuilder(
    //            sheet.Columns.First(col => col.Title == "FirstName").Id,
    //            formDataList.FirstName
    //        ).Build(),
    //                // Add other cells for HCPName, GOorNGO, MISCode, etc.
    //            })
    //            .Build();

    //        smartsheet.SheetResources.RowResources.AddRows(parsedSheetId, new Row[] { rowToAdd });

    //        return Ok("Data added successfully.");
    //    }


    //}







////}












