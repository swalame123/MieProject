using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using MieProject.Models;
using Smartsheet.Api;
using Smartsheet.Api.Models;

namespace MieProject.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class TempController : ControllerBase
    {
        [HttpPost("AddFormData")]
        public IActionResult AddFormData( FileUploadodel fileUploadModel)
        {
            try
            {
                var accessToken = "jQ7rAWlaTgbtMPVvlc7RGOqeNqDWwheJRNV83";
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

                var sheetId1 = 6857973674495876;
                Sheet sheet = smartsheet.SheetResources.GetSheet(sheetId1, null, null, null, null, null, null, null);

                var newRow = new Row();
                newRow.Cells = new List<Cell>();
               
                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet, "GENDER"), Value = fileUploadModel.Gender.GENDER });

                var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheetId1, new Row[] { newRow });

                if (fileUploadModel.File != null && fileUploadModel.File.Length > 0)
                {
                    var fileName = fileUploadModel.File.FileName;
                    var folderName = Path.Combine("Resources", "Images");
                    var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                    var fullPath = Path.Combine(pathToSave, fileName);
                 
                    if (!Directory.Exists(pathToSave))
                    {
                        Directory.CreateDirectory(pathToSave);
                    }

                  
                    using (var fileStream = new FileStream(fullPath, FileMode.Create))
                    {
                        fileUploadModel.File.CopyTo(fileStream);
                    }

                    var addedRow = addedRows[0];
                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                        sheetId1, addedRow.Id.Value, fullPath, "application/msword");
                }

                return Ok("Data added successfully.");
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
