using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using IndiaEventsWebApi.Models;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using IndiaEventsWebApi.Models.EventTypeSheets;
using System.Text;
using static System.Net.Mime.MediaTypeNames;

namespace IndiaEventsWebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class TempController : ControllerBase
    {





        [HttpPost("AddFile")]
        public IActionResult AddFormData(Class11 i)
        {
            try
            {


                var accessToken = "jQ7rAWlaTgbtMPVvlc7RGOqeNqDWwheJRNV83";
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

                var sheetId1 = 6857973674495876;
                Sheet sheet = smartsheet.SheetResources.GetSheet(sheetId1, null, null, null, null, null, null, null);

                var newRow = new Row();
                newRow.Cells = new List<Cell>();

                var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheetId1, new Row[] { newRow });
                var x = 1;
                foreach (var p in i.File)
                {



                    byte[] fileBytes = Convert.FromBase64String(p);
                    var folderName = Path.Combine("Resources", "Images");
                    var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                    if (!Directory.Exists(pathToSave))
                    {
                        Directory.CreateDirectory(pathToSave);
                    }

                    string fileType = GetFileType(fileBytes);
                    string fileName = x+" ConvertedFile." + fileType;
                    string filePath = Path.Combine(pathToSave, fileName);

                    //string type = k.ContentType;//fileUploadModel.File.ContentType;
                    var addedRow = addedRows[0];

                    System.IO.File.WriteAllBytes(filePath, fileBytes);
                    string type = GetContentType(fileType);
                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                            sheetId1, addedRow.Id.Value, filePath, "application/msword");
                    x++;
                }


                return Ok(new { Message = "Conversion successful" });


            }
            catch (Exception ex)
            {
                return BadRequest($"Error: {ex.Message}");
            }
        }












        // ////////////////////////////////////////////////SINGLE FILE UPLOAD


        //[HttpPost("AddFile")]
        //public IActionResult AddFormData(Class11 i)
        //{
        //    try
        //    {


        //        var accessToken = "jQ7rAWlaTgbtMPVvlc7RGOqeNqDWwheJRNV83";
        //        SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

        //        var sheetId1 = 6857973674495876;
        //        Sheet sheet = smartsheet.SheetResources.GetSheet(sheetId1, null, null, null, null, null, null, null);

        //        var newRow = new Row();
        //        newRow.Cells = new List<Cell>();

        //        var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheetId1, new Row[] { newRow });


        //        byte[] fileBytes = Convert.FromBase64String(i.File);
        //        var folderName = Path.Combine("Resources", "Images");
        //        var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
        //        if (!Directory.Exists(pathToSave))
        //        {
        //            Directory.CreateDirectory(pathToSave);
        //        }

        //        string fileType = GetFileType(fileBytes);
        //        string fileName = "ConvertedFile." + fileType;
        //        string filePath = Path.Combine(pathToSave, fileName);

        //        //string type = k.ContentType;//fileUploadModel.File.ContentType;
        //        var addedRow = addedRows[0];

        //        System.IO.File.WriteAllBytes(filePath, fileBytes);
        //        string type = GetContentType(fileType);
        //        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
        //                sheetId1, addedRow.Id.Value, filePath, "application/msword");




        //        return Ok(new { Message = "Conversion successful", FilePath = filePath });


        //    }
        //    catch (Exception ex)
        //    {
        //        return BadRequest($"Error: {ex.Message}");
        //    }
        //}













        private string GetContentType(string fileExtension)
        {
            switch (fileExtension.ToLower())
            {
                case "jpg":
                case "jpeg":
                    return "image/jpeg";
                case "pdf":
                    return "application/pdf";
                case "gif":
                    return "image/gif";
                case "png":
                    return "image/png";
                case "webp":
                    return "image/webp";
                case "doc":
                    return "application/msword";
                case "docx":
                    return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                default:
                    return "application/octet-stream";
            }
        }

        private string GetFileType(byte[] bytes)
        {

            if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xD8)
            {
                return "jpg";
            }
            else if (bytes.Length >= 4 && Encoding.UTF8.GetString(bytes, 0, 4) == "%PDF")
            {
                return "pdf";
            }
            else if (bytes.Length >= 3 && Encoding.UTF8.GetString(bytes, 0, 3) == "GIF")
            {
                return "gif";
            }
            else if (bytes.Length >= 8 && Encoding.UTF8.GetString(bytes, 0, 8) == "PNG\r\n\x1A\n")
            {
                return "png";
            }
            else if (bytes.Length >= 4 && Encoding.UTF8.GetString(bytes, 0, 4) == "RIFF" && Encoding.UTF8.GetString(bytes, 8, 4) == "WEBP")
            {
                return "webp";
            }
            else if (bytes.Length >= 4 && (bytes[0] == 0xD0 && bytes[1] == 0xCF && bytes[2] == 0x11 && bytes[3] == 0xE0))
            {
                return "doc"; // .doc format
            }
            else if (bytes.Length >= 4 && (bytes[0] == 0x50 && bytes[1] == 0x4B && bytes[2] == 0x03 && bytes[3] == 0x04))
            {
                return "docx"; // .docx format
            }
            else
            {
                return "unknown";
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



























//[HttpPost("AddFile")]
//public IActionResult AddFormData(Class11 i)
//{
//    var file = i.File;
//    byte[] fileBytes = Convert.FromBase64String(i.File);


//}

//[HttpPost("AddFormData")]
//public IActionResult AddFormData(Class11 i)
//{
//    try
//    {

//        var accessToken = "jQ7rAWlaTgbtMPVvlc7RGOqeNqDWwheJRNV83";
//        SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

//        var sheetId1 = 6857973674495876;
//        Sheet sheet = smartsheet.SheetResources.GetSheet(sheetId1, null, null, null, null, null, null, null);

//        var newRow = new Row();
//        newRow.Cells = new List<Cell>();

//        //newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet, "GENDER"), Value = fileUploadModel.Gender.GENDER });

//        var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheetId1, new Row[] { newRow });

//        byte[] fileBytes = Convert.FromBase64String(i.File);

//        var inputString = Encoding.UTF8.GetString(fileBytes);
//        //using (var ms = new MemoryStream(fileBytes, 0, fileBytes.Length))
//        //{
//        //    Image image = Image.FromStream(ms, true);
//        //    return image;
//        //}


//        if (inputString != null && inputString.Length > 0)
//        {
//            var fileName = "FCPA" + inputString;//fileUploadModel.File.FileName;
//            var folderName = Path.Combine("Resources", "Images");
//            var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
//            var fullPath = Path.Combine(pathToSave, fileName);
//             System.IO.File.WriteAllBytes(fullPath, fileBytes);
//            if (!Directory.Exists(pathToSave))
//            {
//                Directory.CreateDirectory(pathToSave);
//            }


//            using (var fileStream = new FileStream(fullPath, FileMode.Create))
//            {

//                //fileBytes.CopyTo(FileStream);
//            }
//            //string type = fileBytes;//fileUploadModel.File.ContentType;
//            var addedRow = addedRows[0];
//            var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
//                sheetId1, addedRow.Id.Value, fullPath, "application/msword");


//        }




//        return Ok("Data added successfully.");
//    }
//    catch (Exception ex)
//    {
//        return BadRequest(ex.Message);
//    }

//}


//[HttpPost("AddFormData")]
//public IActionResult AddFormData(Class11 i)
//{
//    try
//    {

//        var accessToken = "jQ7rAWlaTgbtMPVvlc7RGOqeNqDWwheJRNV83";
//        SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

//        var sheetId1 = 6857973674495876;
//        Sheet sheet = smartsheet.SheetResources.GetSheet(sheetId1, null, null, null, null, null, null, null);

//        var newRow = new Row();
//        newRow.Cells = new List<Cell>();

//        //newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet, "GENDER"), Value = fileUploadModel.Gender.GENDER });

//        var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheetId1, new Row[] { newRow });

//        byte[] fileBytes = Convert.FromBase64String(i.File);

//        if (fileBytes != null && fileBytes.Length > 0)
//        {
//            var fileName = "FCPA" + fileBytes;//fileUploadModel.File.FileName;
//            var folderName = Path.Combine("Resources", "Images");
//            var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
//            var fullPath = Path.Combine(pathToSave, fileName);

//            if (!Directory.Exists(pathToSave))
//            {
//                Directory.CreateDirectory(pathToSave);
//            }


//            using (var fileStream = new FileStream(fullPath, FileMode.Create))
//            {

//                fileBytes.;
//            }
//            //string type = fileBytes.;//fileUploadModel.File.ContentType;
//            var addedRow = addedRows[0];
//            var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
//                sheetId1, addedRow.Id.Value, fullPath, type);


//        }




//        return Ok("Data added successfully.");
//    }
//    catch (Exception ex)
//    {
//        return BadRequest(ex.Message);
//    }

//}




//NewMultipleUpload
//[HttpPost("AddFormData")]
//public IActionResult AddFormData(IFormFile[] fileUploadModel)
//{
//    try
//    {
//        var accessToken = "jQ7rAWlaTgbtMPVvlc7RGOqeNqDWwheJRNV83";
//        SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

//        var sheetId1 = 6857973674495876;
//        Sheet sheet = smartsheet.SheetResources.GetSheet(sheetId1, null, null, null, null, null, null, null);

//        var newRow = new Row();
//        newRow.Cells = new List<Cell>();

//        //newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet, "GENDER"), Value = fileUploadModel.Gender.GENDER });

//        var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheetId1, new Row[] { newRow });

//        foreach (var k in fileUploadModel)
//        {
//            if (k != null && k.Length > 0)
//            {
//                var fileName = "FCPA" + k;//fileUploadModel.File.FileName;
//                var folderName = Path.Combine("Resources", "Images");
//                var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
//                var fullPath = Path.Combine(pathToSave, fileName);

//                if (!Directory.Exists(pathToSave))
//                {
//                    Directory.CreateDirectory(pathToSave);
//                }


//                using (var fileStream = new FileStream(fullPath, FileMode.Create))
//                {
//                    k.CopyTo(fileStream);
//                }
//                string type = k.ContentType;//fileUploadModel.File.ContentType;
//                var addedRow = addedRows[0];
//                var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
//                    sheetId1, addedRow.Id.Value, fullPath, type);


//            }
//        }



//        return Ok("Data added successfully.");
//    }
//    catch (Exception ex)
//    {
//        return BadRequest(ex.Message);
//    }

//}











//        //Single Upload
//        [HttpPost("AddFormData")]
//        public IActionResult AddFormData([FromForm] FileUploadodel fileUploadModel)
//        {

//            try
//            {
//                var accessToken = "jQ7rAWlaTgbtMPVvlc7RGOqeNqDWwheJRNV83";
//                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

//                var sheetId1 = 6857973674495876;
//                Sheet sheet = smartsheet.SheetResources.GetSheet(sheetId1, null, null, null, null, null, null, null);

//                var newRow = new Row();
//                newRow.Cells = new List<Cell>();

//                newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet, "GENDER"), Value = fileUploadModel.Gender.GENDER });

//                var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheetId1, new Row[] { newRow });


//                if (fileUploadModel.File != null && fileUploadModel.File.Length > 0)
//                {

//                    var fileType = Path.GetExtension(fileUploadModel.File.FileName);
//                    var filePath = System.IO.Path.Combine(Directory.GetCurrentDirectory());
//                    var docName = Path.GetFileName(fileUploadModel.File.FileName);
//                    Guid ID = Guid.NewGuid();
//                    //ExpenseDataRow.Guid = Convert.ToString(ID);
//                    //ExpenseDataRow.DocumentName = docName;
//                    //ExpenseDataRow.DocType = fileType;
//                    var docUrl = Path.Combine(filePath, "Resources\\Images", ID.ToString() + fileType);
//                    //ExpenseDataRow.DocUrl = Path.Combine(filePath, "assets\\Uploads", ID.ToString() + fileType);
//                    //using (var stream = new FileStream(ExpenseDataRow.DocUrl, FileMode.Create))
//                    using (var stream = System.IO.File.Create(docUrl))
//                    {
//                        //Attach_files[k].CopyToAsync(stream);
//                        fileUploadModel.File.CopyTo(stream);

//                    }








//                    //var fileName = "FCPA" + fileUploadModel.File.FileName;//fileUploadModel.File.FileName;
//                    //var folderName = Path.Combine("Resources", "Images");
//                    //var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
//                    //var fullPath = Path.Combine(pathToSave, fileName);

//                    //if (!Directory.Exists(pathToSave))
//                    //{
//                    //    Directory.CreateDirectory(pathToSave);
//                    //}
//                    //using (var stream = System.IO.File.Create(fullPath))
//                    //{
//                    //    //Attach_files[k].CopyToAsync(stream);
//                    //    fileUploadModel.File.CopyTo(stream);

//                    //}


//                    //using (var fileStream = new FileStream(fullPath, FileMode.Create))
//                    //{
//                    //    fileUploadModel.File.CopyTo(fileStream);
//                    //}
//                    string type = fileUploadModel.File.ContentType;//fileUploadModel.File.ContentType;
//                    var addedRow = addedRows[0];
//                    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
//                        sheetId1, addedRow.Id.Value, docUrl, type /*"application/msword"*/);


//                }



//                return Ok("Data added successfully.");
//            }
//            catch (Exception ex)
//            {
//                return BadRequest(ex.Message);
//            }

//        }





//        //formultiple upload
//        //[HttpPost("AddFormData")]
//        //public IActionResult AddFormData( FileUploadodel fileUploadModel)
//        //{
//        //    try
//        //    {
//        //        var accessToken = "jQ7rAWlaTgbtMPVvlc7RGOqeNqDWwheJRNV83";
//        //        SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

//        //        var sheetId1 = 6857973674495876;
//        //        Sheet sheet = smartsheet.SheetResources.GetSheet(sheetId1, null, null, null, null, null, null, null);

//        //        var newRow = new Row();
//        //        newRow.Cells = new List<Cell>();

//        //        newRow.Cells.Add(new Cell { ColumnId = GetColumnIdByName(sheet, "GENDER"), Value = fileUploadModel.Gender.GENDER });

//        //        var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheetId1, new Row[] { newRow });
//        //        foreach (var p in fileUploadModel.File)
//        //        {
//        //            if (p != null && p.Length > 0)
//        //            {
//        //                var fileName = "FCPA"+p.FileName;//fileUploadModel.File.FileName;
//        //                var folderName = Path.Combine("Resources", "Images");
//        //                var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
//        //                var fullPath = Path.Combine(pathToSave, fileName);

//        //                if (!Directory.Exists(pathToSave))
//        //                {
//        //                    Directory.CreateDirectory(pathToSave);
//        //                }


//        //                using (var fileStream = new FileStream(fullPath, FileMode.Create))
//        //                {
//        //                    p.CopyTo(fileStream);
//        //                }
//        //                string type = p.ContentType;//fileUploadModel.File.ContentType;
//        //                var addedRow = addedRows[0];
//        //                var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
//        //                    sheetId1, addedRow.Id.Value, fullPath, type);


//        //            }
//        //        }


//        //        return Ok("Data added successfully.");
//        //    }
//        //    catch (Exception ex)
//        //    {
//        //        return BadRequest(ex.Message);
//        //    }

//        //}


//        //[HttpGet("GetFormDataByGender")]
//        //public IActionResult GetFormDataByGender(string gender)
//        //{
//        //    try
//        //    {
//        //        var accessToken = "jQ7rAWlaTgbtMPVvlc7RGOqeNqDWwheJRNV83";
//        //        SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
//        //        var sheetId1 = 6857973674495876;

//        //        
//        //        var genderColumnId = GetColumnIdByName(smartsheet.SheetResources.GetSheet(sheetId1, null, null, null, null, null, null, null), "GENDER");

//        //        var filter = new SearchFilter
//        //        {
//        //            ColumnIds = new long[] { genderColumnId },
//        //            Text = gender
//        //        };

//        //        var filteredRows = smartsheet.SearchResources.Filter(sheetId1, filter);



//        //        return Ok(filteredRows);
//        //    }
//        //    catch (Exception ex)
//        //    {
//        //        return BadRequest(ex.Message);
//        //    }
//        //}





//        private long GetColumnIdByName(Sheet sheet, string columnname)
//        {
//            foreach (var column in sheet.Columns)
//            {
//                if (column.Title == columnname)
//                {
//                    return column.Id.Value;
//                }
//            }
//            return 0;
//        }
//    }
//}

////}
////}
