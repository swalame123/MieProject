using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using MieProject.Models;
using MieProject.Models.RequestSheets;
using MieProject.Models.Test;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Text;

namespace MieProject.Controllers.Testing
{
    [Route("api/[controller]")]
    [ApiController]
    public class TestController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;

        public TestController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

        }
        [HttpGet("GetApprovedSpeakersData")]
        public IActionResult GetApprovedSpeakersData()
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
                string sheetId = configuration.GetSection("SmartsheetSettings:TestApprovedSpeakers").Value;
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
        [HttpGet("GetEventRequestsHcpRoleById/{eventIdorEventRequestId}")]
        public IActionResult GetEventRequestsHcpRoleById(string eventIdorEventRequestId)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
                long.TryParse(sheetId, out long parsedSheetId);
                Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
                List<Dictionary<string, object>> hcpRoleData = new List<Dictionary<string, object>>();
                List<string> columnNames = new List<string>();

                foreach (Column column in sheet1.Columns)
                {
                    columnNames.Add(column.Title);
                }
                foreach (Row row in sheet1.Rows)
                {
                   
                    // Check if the row has the specified EventIdorEventRequestId
                    var eventIdorEventRequestIdCell = row.Cells.FirstOrDefault(cell => cell.ColumnId == GetColumnIdByName(sheet1, "EventId/EventRequestId"));
                    var x= eventIdorEventRequestIdCell.Value.ToString();
                   // Console.WriteLine($"Row {row.RowNumber}: EventId/EventRequestId in cell: {eventIdorEventRequestIdCell.Value}");
                    if (eventIdorEventRequestIdCell != null && x == eventIdorEventRequestId)
                    {
                        Dictionary<string, object> hcpRoleRowData = new Dictionary<string, object>();

                        for (int i = 0; i < row.Cells.Count && i < columnNames.Count; i++)
                        {
                            hcpRoleRowData[columnNames[i]] = row.Cells[i].Value;
                        }

                        hcpRoleData.Add(hcpRoleRowData);
                    }
                    
                }

                return Ok(hcpRoleData);

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpGet("GetEventRequestsHcpRole")]
        public IActionResult GetEventRequestsHcpRole()
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
                long.TryParse(sheetId, out long parsedSheetId);
                Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
                List<Dictionary<string, object>> hcpRoleData = new List<Dictionary<string, object>>();
                List<string> columnNames = new List<string>();

                foreach (Column column in sheet1.Columns)
                {
                    columnNames.Add(column.Title);
                }
                foreach (Row row in sheet1.Rows)
                {

                    // Check if the row has the specified EventIdorEventRequestId
                    var eventIdorEventRequestIdCell = row.Cells.FirstOrDefault(cell => cell.ColumnId == GetColumnIdByName(sheet1, "EventId/EventRequestId"));
                    var x = eventIdorEventRequestIdCell.Value.ToString();
                    // Console.WriteLine($"Row {row.RowNumber}: EventId/EventRequestId in cell: {eventIdorEventRequestIdCell.Value}");
                    if (eventIdorEventRequestIdCell != null)
                    {
                        Dictionary<string, object> hcpRoleRowData = new Dictionary<string, object>();

                        for (int i = 0; i < row.Cells.Count && i < columnNames.Count; i++)
                        {
                            hcpRoleRowData[columnNames[i]] = row.Cells[i].Value;
                        }

                        hcpRoleData.Add(hcpRoleRowData);
                    }

                }

                return Ok(hcpRoleData);

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpPost("GetEventRequestsHcpRoleByIds")]
        public IActionResult GetEventRequestsHcpRoleByIds([FromBody] getIds eventIdorEventRequestIds)
        {
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
            long.TryParse(sheetId, out long parsedSheetId);
            Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
            List<Dictionary<string, object>> hcpRoleData = new List<Dictionary<string, object>>();
            List<string> columnNames = new List<string>();
            foreach (Column column in sheet1.Columns)
            {
                columnNames.Add(column.Title);
            }
            foreach (var val in eventIdorEventRequestIds.EventIds)
            {
                foreach (Row row in sheet1.Rows)
                {
                    var eventIdorEventRequestIdCell = row.Cells.FirstOrDefault(cell => cell.ColumnId == GetColumnIdByName(sheet1, "EventId/EventRequestId"));
                    var x = eventIdorEventRequestIdCell.Value.ToString();
                    if (eventIdorEventRequestIdCell != null && x == val)
                    {
                        Dictionary<string, object> hcpRoleRowData = new Dictionary<string, object>();

                        for (int i = 0; i < row.Cells.Count && i < columnNames.Count; i++)
                        {
                            hcpRoleRowData[columnNames[i]] = row.Cells[i].Value;
                        }

                        hcpRoleData.Add(hcpRoleRowData);
                    }
                   
                }
            }

            return Ok(hcpRoleData);
          
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
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet3, "BTC/BTE"),
                    Value = formdata.BtcorBte
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet3, "LcAmount"),
                    Value = formdata.LcAmount
                });
                string rowData = $"{addedInviteesDataNo}. {formdata.InviteeName} | {formdata.MISCode} | {formdata.LocalConveyance} | {formdata.BtcorBte} | {formdata.LcAmount}";
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
                    ColumnId = GetColumnIdByName(sheet4, "HcpRole"),
                    Value = formdata.HcpRole
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet4, "GO/NGO"),
                    Value = formdata.GOorNGO
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet4, "MISCode"),
                    Value = formdata.MisCode
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet4, "HCPName"),
                    Value = formdata.HcpName
                });

                string rowData = $"{addedHcpDataNo}. {formdata.HcpRole} | {formdata.HcpName} | {formdata.MisCode} | {formdata.GOorNGO}";
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
                    ColumnId = GetColumnIdByName(sheet1, "SelectedHcp's&TheirDetails"),
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



                var addedRows = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId1, new Row[] { newRow });
                //var eventId = addedRows[0].Id.Value;
                //long eventIdColumnId = GetColumnIdByName(sheet, "Event ID");
                //var eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                var eventIdColumnId = GetColumnIdByName(sheet1, "EventId/EventRequestId");
                var eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                var val = eventIdCell.DisplayValue;
                //return Ok($"Data saved successfully. BrandsList: {brand}, Invitees: {Invitees}, HcpRole: {HCP}, ");


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
                    return Ok("done");
            }



            catch (Exception ex)
            {
                return BadRequest($"Could not find {ex.Message}");
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


        [HttpGet("GetfmvColumnValue")]
        public IActionResult GetfmvColumnValue(string specialty, string columnTitle)
        {
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            string sheetId = configuration.GetSection("SmartsheetSettings:fmv").Value;
            long.TryParse(sheetId, out long parsedSheetId);
            Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);

            Column SpecialityColumn = sheet.Columns.FirstOrDefault(column =>
           string.Equals(column.Title, "Speciality", StringComparison.OrdinalIgnoreCase));
            Column targetColumn = sheet.Columns.FirstOrDefault(column =>
           string.Equals(column.Title, columnTitle, StringComparison.OrdinalIgnoreCase));
            if (targetColumn != null && SpecialityColumn!= null)
            {
                // Find the row with the specified speciality
                Row targetRow = sheet.Rows.FirstOrDefault(row =>
                    row.Cells.Any(cell => cell.ColumnId == SpecialityColumn.Id && cell.Value.ToString() == specialty));

                if (targetRow != null)
                {
                    // Retrieve the value of the specified column for the given speciality
                    var columnValue = targetRow.Cells.FirstOrDefault(cell => cell.ColumnId == targetColumn.Id)?.Value;
                    if (columnValue != null)
                    {
                        return Ok(columnValue);
                    }
                    else
                    {
                        return NotFound($"Value not found for {specialty} in {columnTitle} column.");
                    }
                }
                else
                {
                    return NotFound($"Speciality '{specialty}' not found.");
                }
            }
            else
            {
                return NotFound($"Column '{columnTitle}' not found.");
            }
        }
       
    }
}
