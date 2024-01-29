using IndiaEventsWebApi.Models.MasterSheets.CodeCreation;
using IndiaEventsWebApi.Models.RequestSheets;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;

namespace IndiaEventsWebApi.Controllers.MasterSheets.CodeCreation
{
    [Route("api/[controller]")]
    [ApiController]
    public class SpeakerCodeCreationController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;

        public SpeakerCodeCreationController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

        }
        [HttpGet("GetSpeakerCodeCreationData")]
        public IActionResult GetSpeakerCodeCreationData()
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:SpeakerCodeCreation").Value;
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



        [HttpPost("AddSpeakersData")]
        public IActionResult AddSpeakersData(SpeakerCodeGeneration formData)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

                string sheetId = configuration.GetSection("SmartsheetSettings:SpeakerCodeCreation").Value;


                long.TryParse(sheetId, out long parsedSheetId);

                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);


                var newRow = new Row();
                newRow.Cells = new List<Cell>();
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Speaker Name"),
                    Value = formData.SpeakerName
                });

                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Speaker Code"),
                    Value = formData.SpeakerCode
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "MIS Code"),
                    Value = formData.MISCode
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Division"),
                    Value = formData.Division
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Speciality"),
                    Value = formData.Speciality
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Qualification"),
                    Value = formData.Qualification
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Address"),
                    Value = formData.Address
                });
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
                    ColumnId = GetColumnIdByName(sheet, "Country"),
                    Value = formData.Country
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Contact Number"),
                    Value = formData.Contact_Number
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "GO/NON-GO"),
                    Value = formData.GOorNGO
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "PAN Card"),
                    Value = formData.PAN_Card
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Speaker Criteria"),
                    Value = formData.Speaker_Criteria
                });



                var addedRows = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId, new Row[] { newRow });
               










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


