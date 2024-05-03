using IndiaEventsWebApi.Helper;
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
        private readonly SmartsheetClient smartsheet;
        public SpeakerCodeCreationController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
            smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
        }
        [HttpGet("GetSpeakerCodeCreationData")]
        public IActionResult GetSpeakerCodeCreationData()
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:SpeakerCodeCreation").Value;
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                List<Dictionary<string, object>> sheetData = SheetHelper.GetSheetData(sheet);
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
                string sheetId = configuration.GetSection("SmartsheetSettings:SpeakerCodeCreation").Value;
                long.TryParse(sheetId, out long parsedSheetId);
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
                string[] sheetIds = {
                configuration.GetSection("SmartsheetSettings:ApprovedSpeakers").Value,
                };
                var mis = "";
                var sheetval = "";
                foreach (string i in sheetIds)
                {
                    long.TryParse(i, out long p);
                    Sheet sheeti = smartsheet.SheetResources.GetSheet(p, null, null, null, null, null, null, null);
                    Column misCodeColumn = sheeti.Columns.FirstOrDefault(column => column.Title == "MisCode");
                    if (misCodeColumn != null)
                    {
                        Row existingRow = sheeti.Rows.FirstOrDefault(row => row.Cells != null && row.Cells.Any(cell => cell.ColumnId == misCodeColumn.Id && cell.Value != null && cell.Value.ToString() == formData.MISCode));
                        if (existingRow != null)
                        {
                            mis = formData.MISCode;
                            sheetval = sheeti.Name;
                        }
                    }
                }
                if (mis != "")
                {
                    return Ok($"MIS Code: {formData.MISCode} already exist in sheetname:{sheetval}");
                }
                else
                {
                    var newRow = new Row();
                    newRow.Cells = new List<Cell>();
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Initiator Name"), Value = formData.InitiatorNameName });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Initiator Email"), Value = formData.InitiatorEmail });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Sales Head"), Value = formData.SalesHead });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Medical Affairs Head"), Value = formData.MedicalAffairsHead });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "SpeakerName"), Value = formData.SpeakerName });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Speaker Code"), Value = formData.SpeakerCode });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "MisCode"), Value = SheetHelper.MisCodeCheck(formData.MISCode )});
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Division"), Value = formData.Division });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Speciality"), Value = formData.Speciality });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Qualification"), Value = formData.Qualification });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Address"), Value = formData.Address });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "City"), Value = formData.City });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "State"), Value = formData.State });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Country"), Value = formData.Country });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Contact Number"), Value = formData.Contact_Number });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Speaker Type"), Value = formData.Speaker_Type });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Speaker Category"), Value = formData.Speaker_Category });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Speaker Criteria"), Value = formData.Speaker_Criteria });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Speaker Criteria Details"), Value = formData.Speaker_Criteria_Details });


                    var addedRows = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId, new Row[] { newRow });
                }
                return Ok(new
                { Message = "Data added successfully." });


            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }


        [HttpPut("UpdateSpeakerDatausingSpeakerId")]
        public IActionResult UpdateSpeakerDatausingSpeakerId(UpdateSpeakerCodeGeneration formData)
        {
            try
            {
                string sheetId = configuration.GetSection("SmartsheetSettings:ApprovedSpeakers").Value;

                long.TryParse(sheetId, out long parsedSheetId);
                Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);

                Row existingRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == formData.SpeakerId));

                Row updateRow = new Row { Id = existingRow.Id, Cells = new List<Cell>() };

                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Initiator Name"), Value = formData.InitiatorNameName });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Initiator Email"), Value = formData.InitiatorEmail });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Sales Head"), Value = formData.SalesHead });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Medical Affairs Head"), Value = formData.MedicalAffairsHead });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "SpeakerName"), Value = formData.SpeakerName });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Speaker Code"), Value = formData.SpeakerCode });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "MisCode"), Value = SheetHelper.MisCodeCheck(formData.MISCode) });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Division"), Value = formData.Division });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Speciality"), Value = formData.Speciality });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Qualification"), Value = formData.Qualification });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Address"), Value = formData.Address });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "City"), Value = formData.City });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "State"), Value = formData.State });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Country"), Value = formData.Country });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Contact Number"), Value = formData.Contact_Number });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Speaker Type"), Value = formData.Speaker_Type });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Speaker Category"), Value = formData.Speaker_Category });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Speaker Criteria"), Value = formData.Speaker_Criteria });
                updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Speaker Criteria Details"), Value = formData.Speaker_Criteria_Details });



                var updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId, new Row[] { updateRow });

                return Ok(new { Message = "Data Updated successfully." });
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

    }

}


