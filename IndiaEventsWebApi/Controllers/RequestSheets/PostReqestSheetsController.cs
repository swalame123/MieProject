using Microsoft.AspNetCore.Components.Forms;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using IndiaEventsWebApi.Models;
using IndiaEventsWebApi.Models.EventTypeSheets;
using IndiaEventsWebApi.Models.RequestSheets;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using Smartsheet.Api.OAuth;
using System.Text;
using System.Globalization;

namespace IndiaEventsWebApi.Controllers.RequestSheets
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

            var TotalHonorariumAmount = 0;
            var TotalTravelAmount = 0;
            var TotalAccomodateAmount = 0;
            var TotalHCPLcAmount = 0;
            var TotalInviteesLcAmount = 0;
            var TotalExpenseAmount = 0;

            CultureInfo hindi = new CultureInfo("hi-IN");





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
                var amount = int.Parse(formdata.Amount);
                TotalExpenseAmount = TotalExpenseAmount + amount;

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



                // string rowData = $"{addedInviteesDataNo}. Name: {formdata.InviteeName} | MIS Code: {formdata.MISCode} | LocalConveyance: {formdata.LocalConveyance} ";
                string rowData = $"{addedInviteesDataNo}. {formdata.InviteeName}";
                addedInviteesData.AppendLine(rowData);
                addedInviteesDataNo++;
                TotalInviteesLcAmount = TotalInviteesLcAmount + int.Parse(formdata.LcAmount);
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
                newRow.Cells.Add(new Cell
                {

                    Value = formdata.LocalConveyance
                });

                var HM = int.Parse(formdata.HonarariumAmount);
                var x = string.Format(hindi, "{0:#,#}", HM);
                var t = int.Parse(formdata.Travel) + int.Parse(formdata.Accomdation);
                var y = string.Format(hindi, "{0:#,#}", t);
                //string rowData = $"{addedHcpDataNo}. {formdata.HcpRole} |Name: {formdata.HcpName} | Honr.Amt: {formdata.HonarariumAmount} |Trav.Amt: {formdata.Travel} |Acco.Amt: {formdata.Accomdation} ";
                string rowData = $"{addedHcpDataNo}. {formdata.HcpRole} |{formdata.HcpName} | Honr.Amt: {x} |Trav.&Acc.Amt: {y} ";

                addedHcpData.AppendLine(rowData);
                addedHcpDataNo++;
                TotalHonorariumAmount = TotalHonorariumAmount +int.Parse( formdata.HonarariumAmount);
                TotalTravelAmount = TotalTravelAmount + int.Parse(formdata.Travel);
                TotalAccomodateAmount=TotalAccomodateAmount+ int.Parse(formdata.Accomdation);
                TotalHCPLcAmount = TotalHCPLcAmount + int.Parse(formdata.LocalConveyance);
            }
            string HCP = addedHcpData.ToString();



            var FormattedTotalHonorariumAmount = string.Format(hindi, "{0:#,#}", TotalHonorariumAmount);
            var FormattedTotalTravelAmount = string.Format(hindi, "{0:#,#}", TotalTravelAmount);
            var FormattedTotalAccomodateAmount = string.Format(hindi, "{0:#,#}", TotalAccomodateAmount);
            var FormattedTotalHCPLcAmount = string.Format(hindi, "{0:#,#}", TotalHCPLcAmount);
            var FornattedTotalInviteesLcAmount = string.Format(hindi, "{0:#,#}", TotalInviteesLcAmount);
            var FormattedTotalExpenseAmount = string.Format(hindi, "{0:#,#}", TotalExpenseAmount);
            var c = TotalHCPLcAmount + TotalInviteesLcAmount;
            var FormattedTotalLC = string.Format(hindi, "{0:#,#}", c);
            var total = TotalHonorariumAmount + TotalTravelAmount + TotalAccomodateAmount + TotalHCPLcAmount + TotalInviteesLcAmount + TotalExpenseAmount;

            var FormattedTotal = string.Format(hindi, "{0:#,#}", total);
            var s = (TotalTravelAmount + TotalAccomodateAmount);
            var FormattedTotalTAAmount = string.Format(hindi, "{0:#,#}",s );




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
                    ColumnId = GetColumnIdByName(sheet1, "Brands"),
                    Value = brand
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "Expenses"),
                    Value = Expense
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "Panelists"),
                    Value = HCP
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "Invitees"),
                    Value = Invitees
                });

                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "SlideKits"),
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
                // //////////////////////////////////////////////////////////////
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "RBM/BM"),
                    Value = formDataList.class1.RBMorBM
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "Sales Head"),
                    Value = formDataList.class1.Sales_Head
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "Marketing Head"),
                    Value = formDataList.class1.Marketing_Head
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "Finance Accounts"),
                    Value = formDataList.class1.Finance
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "InitiatorName"),
                    Value = formDataList.class1.InitiatorName
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "Initiator Email"),
                    Value = formDataList.class1.Initiator_Email
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "Total Honorarium Spend"),
                    Value = TotalHonorariumAmount
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "Total Travel Spend"),
                    Value = TotalTravelAmount
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "Total Travel & Accomodation Spend"),
                    Value = s
                });

                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "Total Accomodation Spend"),
                    Value = TotalAccomodateAmount
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "Total Local Conveyance"),
                    Value = c
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "Total Expense"),
                    Value = TotalExpenseAmount
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "Total Spend"),
                    Value = total
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
                        ColumnId = GetColumnIdByName(sheet4, "HonorariumAmount"),
                        Value = formData.HonarariumAmount
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
                    // ///////////////////////////////////////////////////
                  

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
        public IActionResult AddHonorariumData(HonorariumPaymentList formData )
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

                //StringBuilder addedBrandsData = new StringBuilder();
                //StringBuilder addedInviteesData = new StringBuilder();
                StringBuilder addedHcpData = new StringBuilder();
                //StringBuilder addedSlideKitData = new StringBuilder();
                //StringBuilder addedExpences = new StringBuilder();
                //StringBuilder HCP = new StringBuilder();

                //int addedSlideKitDataNo = 1;
                int addedHcpDataNo = 1;
                //int addedInviteesDataNo = 1;
                //int addedBrandsDataNo = 1;
                //int addedExpencesNo = 1;
                //int hcpNo = 1;
                
                
                CultureInfo hindi = new CultureInfo("hi-IN");
                foreach (var i in formData.HcpRoles)
                {
                    
                    string rowData = $"{addedHcpDataNo}. Name:{i.HcpName} | Role:{i.HcpRole} |Code:{i.MisCode} | HCP Type:{i.GOorNGO}| Including GST:{i.IsInclidingGst}| Agreement Amount:{i.AgreementAmount} ";
                   
                    addedHcpData.AppendLine(rowData);
                    addedHcpDataNo++;
                }
                string panalist = addedHcpData.ToString();


                //foreach (var formdata in formData.BrandDetails)
                //{                 
                    
                //    string rowData = $"{addedBrandsDataNo}. {formdata.BrandName} | {formdata.ProjectId} | {formdata.PercentAllocation}";
                //    addedBrandsData.AppendLine(rowData);
                //    addedBrandsDataNo++;
                //}
                //string brand = addedBrandsData.ToString();

                //foreach (var formdata in formData.Invitees)
                //{             
                //    string rowData = $"{addedInviteesDataNo}. Name: {formdata.InviteeName} | MIS Code: {formdata.MISCode} | LocalConveyance: {formdata.LocalConveyance} ";
                    
                //    addedInviteesData.AppendLine(rowData);
                //    addedInviteesDataNo++;
                   
                //}
                //string Invitees = addedInviteesData.ToString();


                //foreach (var formdata in formData.panalist)
                //{
                   

                //    var HM = int.Parse(formdata.HonarariumAmount);
                //    var x = string.Format(hindi, "{0:#,#}", HM);
                //    var t = int.Parse(formdata.TravelAmount) + int.Parse(formdata.AccomdationAmount);
                //    var y = string.Format(hindi, "{0:#,#}", t);
                   
                //    string rowData = $"{hcpNo}. {formdata.HcpRole} |{formdata.HcpName} | Honr.Amt: {HM} |Trav.&Acc.Amt: {t} ";

                //    HCP.AppendLine(rowData);
                //    hcpNo++;
                  
                //}
                //string HCPd = HCP.ToString();

                //foreach (var formdata in formData.hCPSlideKits)
                //{
                    
                //    string rowData = $"{addedSlideKitDataNo}. {formdata.MIS} | {formdata.SlideKitType}";
                //    addedSlideKitData.AppendLine(rowData);
                //    addedSlideKitDataNo++;
                //}
                //string slideKits = addedSlideKitData.ToString();



              
                
                    var newRow = new Row();
                    newRow.Cells = new List<Cell>();

                    //newRow.Cells.Add(new Cell
                    //{
                    //    ColumnId = GetColumnIdByName(sheet, ""),
                    //    Value = Expense
                    //});
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "SlideKits"),
                        Value = formData.RequestHonorariumList.slideKits
                    });

                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "EventId/EventRequestId"),
                        Value = formData.RequestHonorariumList.EventId
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "Event Type"),
                        Value = formData.RequestHonorariumList.EventType
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "Event Date"),
                        Value = formData.RequestHonorariumList.EventDate
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "Event Topic"),
                        Value = formData.RequestHonorariumList.EventTopic
                    });
                   
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "City"),
                        Value = formData.RequestHonorariumList.City
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "State"),
                        Value = formData.RequestHonorariumList.State
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "Start Time"),
                        Value = formData.RequestHonorariumList.StartTime
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "End Time"),
                        Value = formData.RequestHonorariumList.EndTime
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "Venue Name"),
                        Value = formData.RequestHonorariumList.VenueName
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "Total Travel & Accomodation Spend"),
                        Value = formData.RequestHonorariumList.TotalTravelAndAccomodationSpend
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "Total Honorarium Spend"),
                        Value = formData.RequestHonorariumList.TotalHonorariumSpend
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "Total Spend"),
                        Value = formData.RequestHonorariumList.TotalSpend
                    });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Expenses"),
                    Value = formData.RequestHonorariumList.Expenses
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Total Travel Spend"),
                    Value = formData.RequestHonorariumList.TotalTravelSpend
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Total Accomodation Spend"),
                    Value = formData.RequestHonorariumList.TotalAccomodationSpend
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Total Expense"),
                    Value = formData.RequestHonorariumList.TotalExpenses
                });
                newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "Total Local Conveyance"),
                        Value = formData.RequestHonorariumList.TotalLocalConveyance
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "Brands"),
                        Value = formData.RequestHonorariumList.Brands
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "Invitees"),
                        Value = formData.RequestHonorariumList.Invitees
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "Panelists"),
                        Value = panalist
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "Initiator Name"),
                        Value = formData.RequestHonorariumList.InitiatorName
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "Initiator Email"),
                        Value = formData.RequestHonorariumList.InitiatorEmail
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "RBM/BM"),
                        Value = formData.RequestHonorariumList.RBMorBM
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "Compliance"),
                        Value = formData.RequestHonorariumList.Compliance
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "Finance Accounts"),
                        Value = formData.RequestHonorariumList.FinanceAccounts
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "Finance Treasury"),
                        Value = formData.RequestHonorariumList.FinanceTreasury
                    });

                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "Panelists & Agreements"),
                        Value = formData.RequestHonorariumList.slideKits
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "Honorarium Submitted?"),
                        Value = formData.RequestHonorariumList.HonarariumSubmitted
                    });


                    var addedRows = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId, new Row[] { newRow });
                    var eventId = formData.RequestHonorariumList.EventId;
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
                    
                



                return Ok(new
                { Message = "Data added successfully." });

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }



        [HttpPost("AddEventRequestExpensesData")]
        public IActionResult AddEventRequestExpensesData(EventRequestExpenseSheet formData)
        {
            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

                string sheetId = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
                

                long.TryParse(sheetId, out long parsedSheetId);
                
                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
                
                
                
                    var newRow = new Row();
                    newRow.Cells = new List<Cell>();
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "HCP Name"),
                        Value = formData.EventId
                    });

                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "EventId/EventRequestId"),
                        Value = formData.Expense
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "EventType"),
                        Value = formData.Amount
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "HCPRole"),
                        Value = formData.AmountExcludingTax
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "MISCODE"),
                        Value = formData.BtcorBte
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "GO/Non-GO"),
                        Value = formData.BtcAmount
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "IsItincludingGST?"),
                        Value = formData.BteAmount
                    });
                    newRow.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet, "AgreementAmount"),
                        Value = formData.BudgetAmount
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



        [HttpPost("AddEventSettlementData")]
        public IActionResult AddEventSettlementData(EventSettlement formData)
        {

            try
            {
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

                string sheetId = configuration.GetSection("SmartsheetSettings:EventSettlement").Value;
                string sheetId1 = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;

                long.TryParse(sheetId, out long parsedSheetId);
                long.TryParse(sheetId1, out long parsedSheetId1);

                Sheet sheet = smartsheet.SheetResources.GetSheet(parsedSheetId, null, null, null, null, null, null, null);
                Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId1, null, null, null, null, null, null, null);

                //StringBuilder addedBrandsData = new StringBuilder();
                StringBuilder addedInviteesData = new StringBuilder();
                //StringBuilder addedHcpData = new StringBuilder();
                //StringBuilder addedSlideKitData = new StringBuilder();
                StringBuilder addedExpences = new StringBuilder();
                //StringBuilder HCP = new StringBuilder();

                //int addedSlideKitDataNo = 1;
                //int addedHcpDataNo = 1;
                int addedInviteesDataNo = 1;
                //int addedBrandsDataNo = 1;
                int addedExpencesNo = 1;
                //int hcpNo = 1;


                //CultureInfo hindi = new CultureInfo("hi-IN");
                //foreach (var i in formData.HcpRoles)
                //{

                //    string rowData = $"{addedHcpDataNo}. Name:{i.HcpName} | Role:{i.HcpRole} |Code:{i.MisCode} | HCP Type:{i.GOorNGO}| Including GST:{i.IsInclidingGst}| Agreement Amount:{i.AgreementAmount} ";

                //    addedHcpData.AppendLine(rowData);
                //    addedHcpDataNo++;
                //}
                //string slideKit = addedHcpData.ToString();


                //foreach (var formdata in formData.branddetails)
                //{

                //    string rowData = $"{addedBrandsDataNo}. {formdata.BrandName} | {formdata.ProjectId} | {formdata.PercentAllocation}";
                //    addedBrandsData.AppendLine(rowData);
                //    addedBrandsDataNo++;
                //}
                //string brand = addedBrandsData.ToString();

                //foreach (var formdata in formData.Invitee)
                //{
                //    string rowData = $"{addedInviteesDataNo}. Name: {formdata.InviteeName} | MIS Code: {formdata.MISCode} | LocalConveyance: {formdata.LocalConveyance} ";

                //    addedInviteesData.AppendLine(rowData);
                //    addedInviteesDataNo++;

                //}
                //string Invitees = addedInviteesData.ToString();


                //foreach (var formdata in formData.panalists)
                //{


                //    //var HM = int.Parse(formdata.HonarariumAmount);
                //    //var x = string.Format(hindi, "{0:#,#}", HM);
                //    //var t = int.Parse(formdata.TravelAmount) + int.Parse(formdata.AccomdationAmount);
                //    //var y = string.Format(hindi, "{0:#,#}", t);

                //    string rowData = $"{hcpNo}. {formdata.HcpRole} |{formdata.HcpName} | Honr.Amt: {formdata.HonarariumAmount} |Trav.&Acc.Amt: {formdata.TravelAmount} ";

                //    HCP.AppendLine(rowData);
                //    hcpNo++;

                //}
                //string HCPd = HCP.ToString();

                //foreach (var formdata in formData.hCPSlideKits)
                //{

                //    string rowData = $"{addedSlideKitDataNo}. {formdata.MIS} | {formdata.SlideKitType}";
                //    addedSlideKitData.AppendLine(rowData);
                //    addedSlideKitDataNo++;
                //}
                //string slideKits = addedSlideKitData.ToString();



                foreach (var formdata in formData.expenseSheets)
                {
                   
                    string rowData = $"{addedExpencesNo}. {formdata.Expense} | AmountExcludingTax: {formdata.AmountExcludingTax}| Amount: {formdata.Amount} | {formdata.BtcorBte}";
                    addedExpences.AppendLine(rowData);
                    addedExpencesNo++;
                }
                string Expense = addedExpences.ToString();


                foreach (var formdata in formData.Invitee)
                {

                    string rowData = $"{addedInviteesDataNo}. {formdata.InviteeName} | {formdata.MISCode} | {formdata.LocalConveyance}";
                    addedInviteesData.AppendLine(rowData);
                    addedInviteesDataNo++;
                }
                string Invitee = addedInviteesData.ToString();


                // /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                var newRow = new Row();
                newRow.Cells = new List<Cell>();
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "EventId/EventRequestId"),
                    Value = formData.EventId
                });

                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "EventTopic"),
                    Value = formData.EventTopic
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "EventType"),
                    Value = formData.EventType
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "EventDate"),
                    Value = formData.EventDate
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "StartTime"),
                    Value = formData.StartTime
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "EndTime"),
                    Value = formData.EndTime
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "VenueName"),
                    Value = formData.VenueName
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
                    ColumnId = GetColumnIdByName(sheet, "Attended"),
                    Value = formData.Attended
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "InviteesParticipated"),
                    Value = Invitee
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "ExpenseDetails"),
                    Value = Expense
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "TotalExpenseDetails"),
                    Value = formData.TotalExpense
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "AdvanceDetails"),
                    Value = formData.Advance
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "InitiatorName"),
                    Value = formData.InitiatorName
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Brands"),
                    Value = formData.Brands
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Invitees"),
                    Value = Invitee
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Panelists"),
                    Value = formData.Panalists
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "SlideKits"),
                    Value = formData.SlideKits
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Expenses"),
                    Value = Expense
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Total Invitees"),
                    Value = formData.totalInvitees
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Total Attendees"),
                    Value = formData.TotalAttendees
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Initiator Email"),
                    Value = formData.InitiatorEmail
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "IsAdvanceRequired"),
                    Value = formData.IsAdvanceRequired
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "PostEventSubmitted?"),
                    Value = formData.PostEventSubmitted
                });

                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "RBM/BM"),
                    Value = formData.RBMorBM
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Compliance"),
                    Value = formData.Compliance
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Finance Accounts"),
                    Value = formData.FinanceAccounts
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Finance Treasury"),
                    Value = formData.FinanceTreasury
                });

                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Sales Head"),
                    Value = formData.SalesHead
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Marketing Head"),
                    Value = formData.MarkeringHead
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Medical Affairs Head "),
                    Value = formData.MedicalAffairsHead
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Finance Head"),
                    Value = formData.FinanceHead
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Finance"),
                    Value = formData.FinanceHead
                });

                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Total Travel Spend"),
                    Value = formData.TotalTravelSpend
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Total Accomodation Spend"),
                    Value = formData.TotalAccomodationSpend
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Total Expenses"),
                    Value = formData.TotalExpenses
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Total Travel & Accomodation Spend"),
                    Value = formData.TotalTravelAndAccomodationSpend
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Total Honorarium Spend"),
                    Value = formData.TotalHonorariumSpend
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Total Spend"),
                    Value = formData.TotalSpend
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet, "Total Local Conveyance"),
                        Value = formData.TotalLocalConveyance
                    });


                var addedRows = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId, new Row[] { newRow });

                //var eventIdColumnId = GetColumnIdByName(sheet, "EventId/EventRequestId");
                //var eventIdCell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == eventIdColumnId);
                //var val = eventIdCell.DisplayValue;
                //var targetRow = sheet1.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == val));

                //if (targetRow != null)
                //{
                //    long EventSettlementSubmittedColumnId = GetColumnIdByName(sheet1, "PostEventSubmitted?");
                //    var cellToUpdateB = new Cell
                //    {
                //        ColumnId = EventSettlementSubmittedColumnId,
                //        Value = "Yes"
                //    };
                //    Row updateRow = new Row { Id = targetRow.Id, Cells = new Cell[] { cellToUpdateB } };
                //    var cellToUpdate = targetRow.Cells.FirstOrDefault(c => c.ColumnId == EventSettlementSubmittedColumnId);
                //    if (cellToUpdate != null)
                //    {
                //        cellToUpdate.Value = "Yes";
                //    }

                //    smartsheet.SheetResources.RowResources.UpdateRows(parsedSheetId1, new Row[] { updateRow });

                //}



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