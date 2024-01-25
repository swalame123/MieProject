using Microsoft.AspNetCore.Components.Forms;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using MieProject.Models;
using MieProject.Models.EventTypeSheets;
using MieProject.Models.RequestSheets;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using Smartsheet.Api.OAuth;
using System.Text;

namespace MieProject.Controllers.RequestSheets
{
    [Route("api/[controller]")]
    [ApiController]
    public class PostReqestSheetsController : ControllerBase
    {

        private readonly string accessToken;
        private readonly IConfiguration configuration;

        public PostReqestSheetsController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

        }


        [HttpPost("AllObjModelsData")]
        public IActionResult AllObjModelsData(AllObjModels formDataList)
        {

            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            string sheetId1 = configuration.GetSection("SmartsheetSettings:Class1").Value;
            string sheetId2 = configuration.GetSection("SmartsheetSettings:EventRequestBrandsList").Value;
            string sheetId3 = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
            string sheetId4 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
            string sheetId5 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpSlideKit").Value;
            string sheetId6 = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
            long.TryParse(sheetId1, out long parsedSheetId1);
            long.TryParse(sheetId2, out long parsedSheetId2);
            long.TryParse(sheetId3, out long parsedSheetId3);
            long.TryParse(sheetId4, out long parsedSheetId4);
            long.TryParse(sheetId5, out long parsedSheetId5);
            long.TryParse(sheetId6, out long parsedSheetId6);

            Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId1, null, null, null, null, null, null, null);
            Sheet sheet2 = smartsheet.SheetResources.GetSheet(parsedSheetId2, null, null, null, null, null, null, null);
            Sheet sheet3 = smartsheet.SheetResources.GetSheet(parsedSheetId3, null, null, null, null, null, null, null);
            Sheet sheet4 = smartsheet.SheetResources.GetSheet(parsedSheetId4, null, null, null, null, null, null, null);
            Sheet sheet5 = smartsheet.SheetResources.GetSheet(parsedSheetId5, null, null, null, null, null, null, null);
            Sheet sheet6 = smartsheet.SheetResources.GetSheet(parsedSheetId6, null, null, null, null, null, null, null);

            StringBuilder addedBrandsData = new StringBuilder();
            StringBuilder addedInviteesData = new StringBuilder();
            StringBuilder addedHcpData = new StringBuilder();
            StringBuilder addedSlideKitData = new StringBuilder();
            StringBuilder addedExpences = new StringBuilder();

            int addedSlideKitDataNo = 1;
            int addedHcpDataNo = 1;
            int addedInviteesDataNo = 1;
            int addedBrandsDataNo = 1;
            int addedExpencesNo = 1;

            foreach (var formdata in formDataList.EventRequestExpenseSheet)
            {
                var newRow = new Row();
                newRow.Cells = new List<Cell>();

                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet6, "Expense"),
                    Value = formdata.Expense
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet6, "AmountExcludingTax?"),
                    Value = formdata.AmountExcludingTax
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet6, "Amount"),
                    Value = formdata.Amount
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet6, "BTC/BTE"),
                    Value = formdata.BtcorBte
                });
                string rowData = $"{addedExpencesNo}. {formdata.Expense} | AmountExcludingTax: {formdata.AmountExcludingTax}| Amount: {formdata.Amount} | {formdata.BtcorBte}";
                addedExpences.AppendLine(rowData);
                addedExpencesNo++;
            }
            string Expense = addedExpences.ToString();

            foreach (var formdata in formDataList.EventRequestHCPSlideKits)
            {
                var newRow = new Row();
                newRow.Cells = new List<Cell>();

                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet5, "MIS"),
                    Value = formdata.MIS
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet5, "Slide Kit Type"),
                    Value = formdata.SlideKitType
                });
                string rowData = $"{addedSlideKitDataNo}. {formdata.MIS} | {formdata.SlideKitType}";
                addedSlideKitData.AppendLine(rowData);
                addedSlideKitDataNo++;
            }
            string slideKit = addedSlideKitData.ToString();

            foreach (var formdata in formDataList.RequestBrandsList)
            {
                var newRow = new Row();
                newRow.Cells = new List<Cell>();
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet2, "% Allocation"),
                    Value = formdata.PercentAllocation
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet2, "Brands"),
                    Value = formdata.BrandName
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet2, "Project ID"),
                    Value = formdata.ProjectId
                });
                string rowData = $"{addedBrandsDataNo}. {formdata.BrandName} | {formdata.ProjectId} | {formdata.PercentAllocation}";
                addedBrandsData.AppendLine(rowData);
                addedBrandsDataNo++;
            }
            string brand = addedBrandsData.ToString();

            foreach (var formdata in formDataList.EventRequestInvitees)
            {
                var newRow = new Row();
                newRow.Cells = new List<Cell>();
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet3, "InviteeName"),
                    Value = formdata.InviteeName
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet3, "MISCode"),
                    Value = formdata.MISCode
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet3, "LocalConveyance"),
                    Value = formdata.LocalConveyance
                });
               
              
               
                string rowData = $"{addedInviteesDataNo}. Name: {formdata.InviteeName} | MIS Code {formdata.MISCode} | LocalConveyance: {formdata.LocalConveyance} ";
                addedInviteesData.AppendLine(rowData);
                addedInviteesDataNo++;
            }
            string Invitees = addedInviteesData.ToString();


            foreach (var formdata in formDataList.EventRequestHcpRole)
            {
                var newRow = new Row();
                newRow.Cells = new List<Cell>();
                newRow.Cells.Add(new Cell
                {
                    
                    Value = formdata.HcpRole
                });
                newRow.Cells.Add(new Cell
                {
                    
                    Value = formdata.HcpName
                });
                newRow.Cells.Add(new Cell
                {
                   
                    Value = formdata.Speciality
                });
                newRow.Cells.Add(new Cell
                {
                    
                    Value = formdata.Tier
                });
                newRow.Cells.Add(new Cell
                {

                    Value = formdata.HonarariumAmount
                });
                newRow.Cells.Add(new Cell
                {

                    Value = formdata.Travel
                });
                newRow.Cells.Add(new Cell
                {

                    Value = formdata.Accomdation
                });
                newRow.Cells.Add(new Cell
                {

                    Value = formdata.GOorNGO
                });
               

                string rowData = $"{addedHcpDataNo}.Role: {formdata.HcpRole} |Name: {formdata.HcpName} |Speciality: {formdata.Speciality} |Tier: {formdata.Tier} | Honorarium Amount: {formdata.HonarariumAmount} |Travel Amount: {formdata.Travel} |Accomodation Amount: {formdata.Accomdation} |HCP Type: {formdata.GOorNGO}";
                addedHcpData.AppendLine(rowData);
                addedHcpDataNo++;
            }
            string HCP = addedHcpData.ToString();



            try
            {

                var newRow = new Row();
                newRow.Cells = new List<Cell>();
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "Event Topic"),
                    Value = formDataList.class1.EventTopic
                });

                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "EventType"),
                    Value = formDataList.class1.EventType
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "EventDate"),
                    Value = formDataList.class1.EventDate
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "StartTime"),
                    Value = formDataList.class1.StartTime
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "EndTime"),
                    Value = formDataList.class1.EndTime
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "VenueName"),
                    Value = formDataList.class1.VenueName
                });

                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "City"),
                    Value = formDataList.class1.City
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "State"),
                    Value = formDataList.class1.State
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "SelectedBrands"),
                    Value = brand
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "Expenses"),
                    Value = Expense
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "SelectedPanelists&TheirDetails"),
                    Value = HCP
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "SelectedInvitees"),
                    Value = Invitees
                });

                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "SelectedSlideKit"),
                    Value = slideKit
                });

                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "IsAdvanceRequired"),
                    Value = formDataList.class1.IsAdvanceRequired
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "EventOpen30days"),
                    Value = formDataList.class1.EventOpen30days
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "EventWithin7days"),
                    Value = formDataList.class1.EventWithin7days
                });



                var addedRows = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId1, new Row[] { newRow });
               
                var eventIdColumnId = GetColumnIdByName(sheet1, "EventId/EventRequestId");
                var eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                var val = eventIdCell.DisplayValue;

                //if (formDataList.formFile != null && formDataList.formFile.Length > 0)
                //{
                //    var fileName = formDataList.formFile.FileName;
                //    var folderName = Path.Combine("Resources", "Images");
                //    var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                //    var fullPath = Path.Combine(pathToSave, fileName);
                   
                //    if (!Directory.Exists(pathToSave))
                //    {
                //        Directory.CreateDirectory(pathToSave);
                //    }

                //    using (var fileStream = new FileStream(fullPath, FileMode.Create))
                //    {
                //        formDataList.formFile.CopyTo(fileStream);
                //    }

                //    var addedRow = addedRows[0];
                //    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                //        parsedSheetId1, addedRow.Id.Value, fullPath, "application/msword");
                //}




                foreach (var formData in formDataList.EventRequestHcpRole)
                {
                    var newRow1 = new Row();
                    newRow1.Cells = new List<Cell>();
                    newRow1.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet4, "HcpRole"),
                        Value = formData.HcpRole
                    });

                    newRow1.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet4, "MISCode"),
                        Value = formData.MisCode
                    });
                    newRow1.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet4, "Travel"),
                        Value = formData.Travel
                    });
                    newRow1.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet4, "TotalSpend"),
                        Value = formData.FinalAmount
                    });
                    newRow1.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet4, "Accomodation"),
                        Value = formData.Accomdation
                    });
                    newRow1.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet4, "LocalConveyance"),
                        Value = formData.LocalConveyance
                    });
                    newRow1.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet4, "SpeakerCode"),
                        Value = formData.SpeakerCode
                    });
                    newRow1.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet4, "TrainerCode"),
                        Value = formData.TrainerCode
                    });
                    newRow1.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet4, "HonorariumRequired"),
                        Value = formData.HonorariumRequired
                    });
                    newRow1.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet4, "Speciality"),
                        Value = formData.Speciality
                    });
                    newRow1.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet4, "Tier"),
                        Value = formData.Tier
                    });
                    newRow1.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet4, "GO/NGO"),
                        Value = formData.GOorNGO
                    });
                    newRow1.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet4, "PresentationDuration"),
                        Value = formData.PresentationDuration
                    });
                    newRow1.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet4, "PanelSessionPreparationDuration"),
                        Value = formData.PanelSessionPreperationDuration
                    });
                    newRow1.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet4, "PanelDiscussionDuration"),
                        Value = formData.PanelDisscussionDuration
                    });
                    newRow1.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet4, "QASessionDuration"),
                        Value = formData.QASessionDuration
                    });
                    newRow1.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet4, "BriefingSession"),
                        Value = formData.BriefingSession
                    });
                    newRow1.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet4, "TotalSessionHours"),
                        Value = formData.TotalSessionHours
                    });
                    newRow1.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet4, "Rationale"),
                        Value = formData.Rationale
                    });
                    newRow1.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet4, "EventId/EventRequestId"),
                        Value = val
                    });
                    newRow1.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet4, "HCPName"),
                        Value = formData.HcpName
                    });

                    smartsheet.SheetResources.RowResources.AddRows(parsedSheetId4, new Row[] { newRow1 });







                }

                foreach (var formdata in formDataList.RequestBrandsList)
                {
                    var newRow2 = new Row();
                    newRow2.Cells = new List<Cell>();
                    newRow2.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet2, "% Allocation"),
                        Value = formdata.PercentAllocation
                    });
                    newRow2.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet2, "Brands"),
                        Value = formdata.BrandName
                    });
                    newRow2.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet2, "Project ID"),
                        Value = formdata.ProjectId
                    });
                    newRow2.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet2, "EventId/EventRequestId"),
                        Value = val
                    });

                    smartsheet.SheetResources.RowResources.AddRows(parsedSheetId2, new Row[] { newRow2 });

                }
                foreach (var formdata in formDataList.EventRequestInvitees)
                {
                    var newRow3 = new Row();
                    newRow3.Cells = new List<Cell>();
                    newRow3.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet3, "InviteeName"),
                        Value = formdata.InviteeName
                    });
                    newRow3.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet3, "MISCode"),
                        Value = formdata.MISCode
                    });
                    newRow3.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet3, "LocalConveyance"),
                        Value = formdata.LocalConveyance
                    });
                    newRow3.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet3, "BTC/BTE"),
                        Value = formdata.BtcorBte
                    });
                    newRow3.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet3, "LcAmount"),
                        Value = formdata.LcAmount
                    });
                    newRow3.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet3, "EventId/EventRequestId"),
                        Value = val
                    });

                    smartsheet.SheetResources.RowResources.AddRows(parsedSheetId3, new Row[] { newRow3 });
                }


                foreach (var formdata in formDataList.EventRequestHCPSlideKits)
                {
                    var newRow5 = new Row();
                    newRow5.Cells = new List<Cell>();

                    newRow5.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet5, "MIS"),
                        Value = formdata.MIS
                    });
                    newRow5.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet5, "Slide Kit Type"),
                        Value = formdata.SlideKitType
                    });
                    newRow5.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet5, "SlideKit Document"),
                        Value = formdata.SlideKitDocument
                    });
                    newRow5.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet5, "EventId/EventRequestId"),
                        Value = val
                    });


                    smartsheet.SheetResources.RowResources.AddRows(parsedSheetId5, new Row[] { newRow5 });
                }

                foreach (var formdata in formDataList.EventRequestExpenseSheet)
                {
                    var newRow6 = new Row();
                    newRow6.Cells = new List<Cell>();

                    newRow6.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet6, "Expense"),
                        Value = formdata.Expense
                    });
                    newRow6.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet6, "EventId/EventRequestID"),
                        Value = val
                    });
                    newRow6.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet6, "AmountExcludingTax?"),
                        Value = formdata.AmountExcludingTax
                    });
                    newRow6.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet6, "Amount"),
                        Value = formdata.Amount
                    });
                    newRow6.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet6, "BTC/BTE"),
                        Value = formdata.BtcorBte
                    });
                    newRow6.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet6, "BudgetAmount"),
                        Value = formdata.BudgetAmount
                    });
                    newRow6.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet6, "BTCAmount"),
                        Value = formdata.BtcAmount
                    });
                    newRow6.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet6, "BTEAmount"),
                        Value = formdata.BteAmount
                    });
                    smartsheet.SheetResources.RowResources.AddRows(parsedSheetId6, new Row[] { newRow6 });
                }

                return Ok(new
                {  Message = " Success!" });


               
            }



            catch (Exception ex)
            {
                return BadRequest($"Could not find {ex.Message}");
            }












        }


        [HttpPost("AddHonorariumData")]
        public IActionResult AddHonorariumData(HonorariumPaymentList formData)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

                string sheetId = configuration.GetSection("SmartsheetSettings:HonorariumPayment").Value;
                string sheetId1 = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;

                long.TryParse(sheetId, out long parsedSheetId);
                long.TryParse(sheetId1, out long parsedSheetId1);
                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
                Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId1, null, null, null, null, null, null, null);
                foreach (var i in formData.RequestHonorariumList)
                {
                    var newRow = new Row();
                    newRow.Cells = new List<Cell>();
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "HCP Name"),
                        Value = i.HCPName
                    });

                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "EventId/EventRequestId"),
                        Value = i.EventId
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "EventType"),
                        Value = i.EventType
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "HCPRole"),
                        Value = i.HCPRole
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "MISCODE"),
                        Value = i.MISCode
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "GO/Non-GO"),
                        Value = i.GONGO
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "IsItincludingGST?"),
                        Value = i.IsItincludingGST
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "AgreementAmount"),
                        Value = i.AgreementAmount
                    });



                    var addedRows = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId, new Row[] { newRow });
                    var eventId = i.EventId;
                    var targetRow = sheet1.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == eventId));
                    
                    if (targetRow != null)
                    {
                        long honorariumSubmittedColumnId = GetColumnIdByName(sheet1, "Honorarium Submitted?");
                        var cellToUpdateB = new Cell
                        {
                            ColumnId = honorariumSubmittedColumnId,
                            Value = "Yes"
                        };
                        Row updateRow = new Row { Id = targetRow.Id, Cells = new Cell[] { cellToUpdateB } };
                        var cellToUpdate = targetRow.Cells.FirstOrDefault(c => c.ColumnId == honorariumSubmittedColumnId);
                        if (cellToUpdate != null)
                        {
                            cellToUpdate.Value = "Yes";
                        }

                        smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId1, new Row[] { updateRow });

                    }
                    
                }



                return Ok(new
                { Message = "Data added successfully." });

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }



        [HttpPost("AddEventRequestExpensesData")]
        public IActionResult AddEventRequestExpensesData(EventRequestExpenseSheet[] formData)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
                

                long.TryParse(sheetId, out long parsedSheetId);
                
                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
                
                foreach (var i in formData)
                {
                    var newRow = new Row();
                    newRow.Cells = new List<Cell>();
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "HCP Name"),
                        Value = i.EventId
                    });

                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "EventId/EventRequestId"),
                        Value = i.Expense
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "EventType"),
                        Value = i.Amount
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "HCPRole"),
                        Value = i.AmountExcludingTax
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "MISCODE"),
                        Value = i.BtcorBte
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "GO/Non-GO"),
                        Value = i.BtcAmount
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "IsItincludingGST?"),
                        Value = i.BteAmount
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "AgreementAmount"),
                        Value = i.BudgetAmount
                    });



                    var addedRows = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId, new Row[] { newRow });
                    var eventId = i.EventId;
                   

                    

                    

                }



                return Ok(new
                { Message = "Data added successfully." });

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }



        [HttpPost("AddEventSettlementData")]
        public IActionResult AddEventSettlementData(EventSettlementData formData)
        {

            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

                string sheetId = configuration.GetSection("SmartsheetSettings:EventSettlement").Value;


                long.TryParse(sheetId, out long parsedSheetId);

                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
                StringBuilder addedExpenseData = new StringBuilder();
                StringBuilder addedInviteesData = new StringBuilder();
                int addedInviteesDataNo = 1;
                int addedExpencesNo = 1;


                foreach (var formdata in formData.ExpenseSheet)
                {
                    var newRow1 = new Row();
                    newRow1.Cells = new List<Cell>();

                    newRow1.Cells.Add(new Cell
                    {

                        Value = formdata.Expense
                    });
                    newRow1.Cells.Add(new Cell
                    {

                        Value = formdata.AmountExcludingTax
                    });
                    newRow1.Cells.Add(new Cell
                    {

                        Value = formdata.Amount
                    });
                    newRow1.Cells.Add(new Cell
                    {

                        Value = formdata.BtcorBte
                    });
                    string rowData = $"{addedExpencesNo}. {formdata.Expense} | AmountExcludingTax: {formdata.AmountExcludingTax}| Amount: {formdata.Amount} | {formdata.BtcorBte}";
                    addedExpenseData.AppendLine(rowData);
                    addedExpencesNo++;
                }
                string Expense = addedExpenseData.ToString();


                foreach (var formdata in formData.RequestInvitees)
                {
                    var newRow2 = new Row();
                    newRow2.Cells = new List<Cell>();
                    newRow2.Cells.Add(new Cell
                    {
                        //ColumnId = GetColumnIdByName(sheet3, "InviteeName"),
                        Value = formdata.InviteeName
                    });
                    newRow2.Cells.Add(new Cell
                    {
                        //ColumnId = GetColumnIdByName(sheet3, "MISCode"),
                        Value = formdata.MISCode
                    });
                    newRow2.Cells.Add(new Cell
                    {
                        //ColumnId = GetColumnIdByName(sheet3, "LocalConveyance"),
                        Value = formdata.LocalConveyance
                    });
                    newRow2.Cells.Add(new Cell
                    {
                        // ColumnId = GetColumnIdByName(sheet3, "BTC/BTE"),
                        Value = formdata.BtcorBte
                    });
                    newRow2.Cells.Add(new Cell
                    {
                        // ColumnId = GetColumnIdByName(sheet3, "LcAmount"),
                        Value = formdata.LcAmount
                    });
                    string rowData = $"{addedInviteesDataNo}. {formdata.InviteeName} | {formdata.MISCode} | {formdata.LocalConveyance} | {formdata.BtcorBte} | {formdata.LcAmount}";
                    addedInviteesData.AppendLine(rowData);
                    addedInviteesDataNo++;
                }
                string Invitees = addedInviteesData.ToString();




                var newRow = new Row();
                newRow.Cells = new List<Cell>();
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "EventId/EventRequestId"),
                    Value = formData.EventSettlement.EventId
                });

                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "EventTopic"),
                    Value = formData.EventSettlement.EventTopic
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "EventType"),
                    Value = formData.EventSettlement.EventType
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "EventDate"),
                    Value = formData.EventSettlement.EventDate
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "StartTime"),
                    Value = formData.EventSettlement.StartTime
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "EndTime"),
                    Value = formData.EventSettlement.EndTime
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "VenueName"),
                    Value = formData.EventSettlement.VenueName
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "City"),
                    Value = formData.EventSettlement.City
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "State"),
                    Value = formData.EventSettlement.State
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Attended"),
                    Value = formData.EventSettlement.Attended
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "InviteesParticipated"),
                    Value = Invitees
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "ExpenseDetails"),
                    Value = Expense
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "TotalExpenseDetails"),
                    Value = formData.EventSettlement.TotalExpense
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "AdvanceDetails"),
                    Value = formData.EventSettlement.Advance
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "InitiatorName"),
                    Value = formData.EventSettlement.InitiatorName
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
        private Row GetRowById(SmartsheetClient smartsheet, long sheetId, string val)
        {
            Sheet sheet = smartsheet.SheetResources.GetSheet(sheetId, null, null, null, null, null, null, null);

            // Assuming you have a column named "Id"

            Column idColumn = sheet.Columns.FirstOrDefault(col => col.Title == "Honorarium Submitted?");

            if (idColumn != null)
            {
                foreach (var row in sheet.Rows)
                {
                    var cell = row.Cells.FirstOrDefault(c => c.ColumnId == idColumn.Id && c.Value.ToString() == val);

                    if (cell != null)
                    {
                        return row;
                    }
                }
            }

            return null;
        }
    }
}







  //var eventIdColumnId = GetColumnIdByName(sheet, "EventId/EventRequestId");
  //                  var eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
  //                  var val = eventIdCell.DisplayValue;
  //                  var IsHonorarium = "Yes";
  //                  Row existingRow = GetRowById(smartsheet, parsedSheetId1, val);
  //                  Row updateRow = new Row { Id = existingRow.Id, Cells = new List<Cell>() };


  //                  if (existingRow == null)
  //                  {
  //                      return NotFound($"Row with id {val} not found.");
  //                  }

  //                  foreach (var cell in existingRow.Cells)
  //                  {
  //                      if (cell.ColumnId == GetColumnIdByName(sheet, "Honorarium Submitted?"))
  //                      {
  //                          cell.Value = IsHonorarium;
  //                      }
  //                      updateRow.Cells.Add(cell);
  //                      //else
  //                      //{
  //                      //    c
  //                      //}

  //                  }
  //                  smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId1, new Row[] { updateRow });