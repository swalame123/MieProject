using IndiaEventsWebApi.Models;
using IndiaEventsWebApi.Models.MasterSheets;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using Smartsheet.Api.OAuth;

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
        [HttpGet("GetHCPDataUsingNameAndMISCode")]
        public IActionResult GetHCPDataUsingNameAndMISCode(string Name, string misCode)
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
                Column hcpNameColumn = sheeti.Columns.FirstOrDefault(column => column.Title == "HCPName");
                Column misCodeColumn = sheeti.Columns.FirstOrDefault(column => column.Title == "MisCode");

                if (hcpNameColumn != null && misCodeColumn != null)
                {
                    Row existingRow = sheeti.Rows.FirstOrDefault(row =>
                        row.Cells != null &&
                        row.Cells.Any(cell =>
                            cell.ColumnId == hcpNameColumn.Id && cell.Value != null && cell.Value.ToString() == Name
                        ) &&
                        row.Cells.Any(cell =>
                            cell.ColumnId == misCodeColumn.Id && cell.Value != null && cell.Value.ToString() == misCode
                        )
                    );
                    if (existingRow != null)
                    {
                        // Both Name and MISCode are present in the same row, return success
                        return Ok("True");
                    }

                }
            }
            return Ok("False");
        }




        [HttpPost("PostHcpData")]
        public IActionResult PostHcpData1(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("Please provide a valid Excel file.");
            }
            int totalCount = 0, insertedCount = 0;

            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            string sheetId = configuration.GetSection("SmartsheetSettings:HcpMaster1").Value;
            long.TryParse(sheetId, out long parsedSheetId);

            Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
            using (var memoryStream = new MemoryStream())
            {
                file.CopyTo(memoryStream);
                var workbook = new Aspose.Cells.Workbook(memoryStream);


                var worksheet = workbook.Worksheets[0];

                List<HCPMaster1> formDataList = new List<HCPMaster1>();


                for (int row = 1; row <= worksheet.Cells.MaxDataRow; row++)
                {
                    var formData = new HCPMaster1
                    {
                        FirstName = worksheet.Cells[row, GetColumnIndexByName(worksheet, "FirstName")].StringValue,
                        LastName = worksheet.Cells[row, GetColumnIndexByName(worksheet, "LastName")].StringValue,
                        HCPName = worksheet.Cells[row, GetColumnIndexByName(worksheet, "HCPName")].StringValue,
                        GOorNGO = worksheet.Cells[row, GetColumnIndexByName(worksheet, "HCP Type")].StringValue,
                        MISCode = worksheet.Cells[row, GetColumnIndexByName(worksheet, "MisCode")].StringValue,
                        Speciality = worksheet.Cells[row, GetColumnIndexByName(worksheet, "Speciality")].StringValue
                    };

                    formDataList.Add(formData);
                }
                totalCount = formDataList.Count();

                foreach (var i in formDataList)
                {
                    var MisCodeValue = "";

                    foreach (string sheetId1 in GetSheetIds())
                    {
                        long.TryParse(sheetId, out long parsedSheetId1);
                        Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId1, null, null, null, null, null, null, null);

                        Row existingRow = sheet1.Rows.FirstOrDefault(row =>
                            row.Cells != null &&
                            row.Cells.Any(cell =>
                                cell.Value != null &&
                                cell.Value.ToString() == i.MISCode)
                        );
                        if (existingRow != null)
                        {
                            MisCodeValue = i.MISCode;
                        }
                    }
                    if (MisCodeValue == "")
                    {
                        //foreach (var formData in formDataList)
                        //{

                        var newRow = new Row();
                        newRow.Cells = new List<Cell>();
                        newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet, "FirstName"), Value = i.FirstName });
                        newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet, "LastName"), Value = i.LastName });
                        newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet, "HCPName"), Value = i.HCPName });
                        newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet, "HCP Type"), Value = i.GOorNGO });
                        newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet, "MisCode"), Value = i.MISCode });
                        newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet, "Speciality"), Value = i.Speciality });


                        smartsheet.SheetResources.RowResources.AddRows(parsedSheetId, new Row[] { newRow });
                        insertedCount++;
                        //}
                    }
                }


                return Ok(new { Message = insertedCount.ToString() + " out of " + totalCount.ToString() + " has been added successfully!" });
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
        private int GetColumnIndexByName(Aspose.Cells.Worksheet worksheet, string columnName)
        {
            var headerRow = worksheet.Cells.GetRow(0);

            for (int col = 0; col <= worksheet.Cells.MaxDataColumn; col++)
            {
                if (headerRow.GetCellOrNull(col)?.StringValue == columnName)
                {
                    return col;
                }
            }

            return -1;
        }

        private IEnumerable<string> GetSheetIds()
        {
            return new[]
            {
                //configuration.GetSection("SmartsheetSettings:HcpMaster").Value,
                configuration.GetSection("SmartsheetSettings:HcpMaster1").Value,
                configuration.GetSection("SmartsheetSettings:HcpMaster2").Value,
                configuration.GetSection("SmartsheetSettings:HcpMaster3").Value,
                configuration.GetSection("SmartsheetSettings:HcpMaster4").Value,
                configuration.GetSection("SmartsheetSettings:ApprovedSpeakers").Value,
                configuration.GetSection("SmartsheetSettings:ApprovedTrainers").Value
            };
        }
    }

}


