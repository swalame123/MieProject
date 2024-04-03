using IndiaEventsWebApi.Helper;
using IndiaEventsWebApi.Models;
using IndiaEventsWebApi.Models.MasterSheets;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using Smartsheet.Api.OAuth;
using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory.Database;

namespace IndiaEventsWebApi.Controllers.HCPMaster
{
    [Route("api/[controller]")]
    [ApiController]
    public class HCPController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;
        private readonly SmartsheetClient smartsheet;
        public HCPController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
            smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
        }



        [HttpGet("GetHCPDataUsingNameAndMISCode")]
        public IActionResult GetHCPDataUsingNameAndMISCode(string Name, string misCode)
        {
            string[] sheetIds = {
                configuration.GetSection("SmartsheetSettings:HcpMaster1").Value,
                configuration.GetSection("SmartsheetSettings:HcpMaster2").Value,
                configuration.GetSection("SmartsheetSettings:HcpMaster3").Value,
                configuration.GetSection("SmartsheetSettings:HcpMaster4").Value
            };
            foreach (string i in sheetIds)
            {
                Sheet sheeti = SheetHelper.GetSheetById(smartsheet, i);
                Column hcpNameColumn = sheeti.Columns.FirstOrDefault(column => column.Title == "HCPName");
                Column misCodeColumn = sheeti.Columns.FirstOrDefault(column => column.Title == "MisCode");
                if (hcpNameColumn != null && misCodeColumn != null)
                {
                    Row existingRow = sheeti.Rows.FirstOrDefault(row => row.Cells != null && row.Cells.Any(cell => cell.ColumnId == hcpNameColumn.Id && cell.Value != null && cell.Value.ToString() == Name) &&
                        row.Cells.Any(cell => cell.ColumnId == misCodeColumn.Id && cell.Value != null && cell.Value.ToString() == misCode));
                    if (existingRow != null)
                    {
                        return Ok("True");
                    }
                }
            }
            return Ok("False");
        }



        [HttpGet("GetRowDataUsingMISCode")]
        public IActionResult GetRowDataUsingMISCode(string misCode)
        {
            var Rowdata = GetExistingRowData(smartsheet, misCode);
            if (Rowdata.Count > 0)
            {
                return Ok(Rowdata);
            }
            return Ok("No Data Found");


        }



        [HttpPost("GetRowDataUsingMISCodeandType")]
        public IActionResult GetRowDataUsingMISCodeandType(MIsandType formdata)
        {
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            string ApprovedSpeakers = configuration.GetSection("SmartsheetSettings:ApprovedSpeakers").Value;
            string ApprovedTrainers = configuration.GetSection("SmartsheetSettings:ApprovedTrainers").Value;
            string VendorMasterSheet = configuration.GetSection("SmartsheetSettings:VendorMasterSheet").Value;
            if (formdata.Type == "Speaker")
            {
                long.TryParse(ApprovedSpeakers, out long p);
                Sheet sheeti = SheetHelper.GetSheetById(smartsheet, ApprovedSpeakers);
                Column misCodeColumn = sheeti.Columns.FirstOrDefault(column => column.Title == "MisCode");
                if (misCodeColumn != null)
                {
                    Row existingRow = sheeti.Rows.FirstOrDefault(row => row.Cells != null && row.Cells.Any(cell => cell.ColumnId == misCodeColumn.Id && cell.Value != null && cell.Value.ToString() == formdata.MISCode));
                    if (existingRow != null)
                    {
                        return Ok(new
                        { Message = "MIS code Already Exist in sheet" });
                    }
                }
            }
            else if (formdata.Type == "Trainer")
            {

                Sheet sheeti = SheetHelper.GetSheetById(smartsheet, ApprovedTrainers);
                Column misCodeColumn = sheeti.Columns.FirstOrDefault(column => column.Title == "MisCode");
                if (misCodeColumn != null)
                {
                    Row existingRow = sheeti.Rows.FirstOrDefault(row => row.Cells != null && row.Cells.Any(cell => cell.ColumnId == misCodeColumn.Id && cell.Value != null && cell.Value.ToString() == formdata.MISCode));
                    if (existingRow != null)
                    {
                        return Ok(new
                        { Message = "MIS code Already Exist in sheet" });
                    }
                }
            }
            else if (formdata.Type == "Vendor")
            {

                Sheet sheeti = SheetHelper.GetSheetById(smartsheet, VendorMasterSheet);
                Column misCodeColumn = sheeti.Columns.FirstOrDefault(column => column.Title == "MisCode");
                if (misCodeColumn != null)
                {
                    Row existingRow = sheeti.Rows.FirstOrDefault(row => row.Cells != null && row.Cells.Any(cell => cell.ColumnId == misCodeColumn.Id && cell.Value != null && cell.Value.ToString() == formdata.MISCode));
                    if (existingRow != null)
                    {
                        return Ok(new
                        { Message = "MIS code Already Exist in sheet" });
                    }
                }
            }

            ConcurrentBag<Dictionary<string, string>> Rowdata = GetExistingRowData(smartsheet, formdata.MISCode);

            return Rowdata.Count > 0 ? Ok(Rowdata) : (IActionResult)Ok(new { Message = "No Data Found" });
        }








        [HttpPost("PostHcpData")]
        public IActionResult PostHcpData1(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("Please provide a valid Excel file.");
            }
            int totalCount = 0, insertedCount = 0;


            string sheetId = configuration.GetSection("SmartsheetSettings:HcpMaster1").Value;
            long.TryParse(sheetId, out long parsedSheetId);

            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId);
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
                        Speciality = worksheet.Cells[row, GetColumnIndexByName(worksheet, "Speciality")].StringValue,
                        CompanyName = worksheet.Cells[row, GetColumnIndexByName(worksheet, "Company Name")].StringValue
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
                            //MisCodeValue = i.MISCode;
                            //Row existingRowId = GetRowById(smartsheet, parsedSheetId, MisCodeValue);

                            var Id = (long)existingRow.Id;
                            smartsheet.SheetResources.RowResources.DeleteRows(parsedSheetId1, new long[] { Id }, true);
                        }
                    }

                    //foreach (var formData in formDataList)
                    //{

                    Row newRow = new Row();
                    newRow.Cells = new List<Cell>();
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "FirstName"), Value = i.FirstName });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "LastName"), Value = i.LastName });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "HCPName"), Value = i.HCPName });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "HCP Type"), Value = i.GOorNGO });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "MisCode"), Value = i.MISCode });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Speciality"), Value = i.Speciality });
                    newRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Company Name"), Value = i.CompanyName });


                    smartsheet.SheetResources.RowResources.AddRows(parsedSheetId, new Row[] { newRow });
                    insertedCount++;
                    //}

                }


                return Ok(new { Message = insertedCount.ToString() + " out of " + totalCount.ToString() + " has been added successfully!" });
            }
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
                configuration.GetSection("SmartsheetSettings:HcpMaster1").Value,
                configuration.GetSection("SmartsheetSettings:HcpMaster2").Value,
                configuration.GetSection("SmartsheetSettings:HcpMaster3").Value,
                configuration.GetSection("SmartsheetSettings:HcpMaster4").Value
            };
        }

        private ConcurrentBag<Dictionary<string, string>> GetExistingRowData(SmartsheetClient smartsheet, string MisCode)
        {


            string[] sheetIds = {

                configuration.GetSection("SmartsheetSettings:HcpMaster1").Value,
                configuration.GetSection("SmartsheetSettings:HcpMaster2").Value,
                configuration.GetSection("SmartsheetSettings:HcpMaster3").Value,
                configuration.GetSection("SmartsheetSettings:HcpMaster4").Value
            };
            ConcurrentBag<Dictionary<string, string>> resultData = new ConcurrentBag<Dictionary<string, string>>();
            bool foundMatch = false;

            Parallel.ForEach(sheetIds, sheetId =>
            {
                if (!foundMatch)
                {
                    long.TryParse(sheetId, out long p);
                    Sheet sheeti = smartsheet.SheetResources.GetSheet(p, null, null, null, null, null, null, null);
                    Column hcpNameColumn = sheeti.Columns.FirstOrDefault(column => column.Title == "HCPName");
                    Column misCodeColumn = sheeti.Columns.FirstOrDefault(column => column.Title == "MisCode");
                    Column Firstname = sheeti.Columns.FirstOrDefault(column => column.Title == "FirstName");
                    Column LastName = sheeti.Columns.FirstOrDefault(column => column.Title == "LastName");
                    Column gongoColumn = sheeti.Columns.FirstOrDefault(column => column.Title == "HCP Type");
                    Column Company = sheeti.Columns.FirstOrDefault(column => column.Title == "Company Name");


                    if (misCodeColumn != null)
                    {
                        Row existingRow = sheeti.Rows.FirstOrDefault(row => row.Cells != null && row.Cells.Any(cell => cell.ColumnId == misCodeColumn.Id && cell.Value != null && cell.Value.ToString() == MisCode));
                        if (existingRow != null)
                        {
                            var hcpNameValue = existingRow.Cells.FirstOrDefault(cell => cell.ColumnId == hcpNameColumn.Id)?.Value.ToString();
                            var gongoValue = existingRow.Cells.FirstOrDefault(cell => cell.ColumnId == gongoColumn.Id)?.Value.ToString();
                            var Mis = existingRow.Cells.FirstOrDefault(cell => cell.ColumnId == misCodeColumn.Id)?.Value.ToString();
                            var firstName = existingRow.Cells.FirstOrDefault(cell => cell.ColumnId == Firstname.Id)?.Value.ToString();
                            var lastName = existingRow.Cells.FirstOrDefault(cell => cell.ColumnId == LastName.Id)?.Value.ToString();
                            var CompanyName = existingRow.Cells.FirstOrDefault(cell => cell.ColumnId == Company.Id)?.Value.ToString();
                            var cellData = new Dictionary<string, string>
                        {
                            { "MIS Code", Mis },
                            { "FirstName", firstName },
                            { "LastName", lastName },
                            { "HCPName", hcpNameValue },
                            { "HcpType", gongoValue },
                            {"Company Name",CompanyName }
                        };
                            resultData.Add(cellData);

                            foundMatch = true;

                        }
                    }
                }


            });
            return resultData;

            //return new ConcurrentBag<Dictionary<string, string>>;
            //return /*"No Data Found"*/;




        }

    }

}














