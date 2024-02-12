using IndiaEventsWebApi.Models.EventTypeSheets;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Globalization;
using System.Text;
//using static IndiaEventsWebApi.Models.EventTypeSheets.MedicalUtility;

namespace IndiaEventsWebApi.Controllers.RequestSheets
{
    [Route("api/[controller]")]
    [ApiController]
    public class medicalPreUtilityController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;

        public medicalPreUtilityController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;

        }
        //[HttpPost("PreEventData")]
        //public IActionResult PreEventData(MedicalPreEventPayload formDataList)
        //{
        //    return Ok(formDataList);
        //}
        [HttpPost("PreEventData")]
        public IActionResult PreEventData(MedicalUtilityPreEventPayload formDataList)
        {

            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();
            string sheetId1 = configuration.GetSection("SmartsheetSettings:Class1").Value;
            string sheetId2 = configuration.GetSection("SmartsheetSettings:EventRequestBrandsList").Value;
            //string sheetId3 = configuration.GetSection("SmartsheetSettings:EventRequestInvitees").Value;
            string sheetId4 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpRole").Value;
            //string sheetId5 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpSlideKit").Value;
            string sheetId6 = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
            string sheetId7 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;



            long.TryParse(sheetId1, out long parsedSheetId1);
            long.TryParse(sheetId2, out long parsedSheetId2);
            //long.TryParse(sheetId3, out long parsedSheetId3);
            long.TryParse(sheetId4, out long parsedSheetId4);
            //long.TryParse(sheetId5, out long parsedSheetId5);
            long.TryParse(sheetId6, out long parsedSheetId6);
            long.TryParse(sheetId7, out long parsedSheetId7);

            Sheet sheet1 = smartsheet.SheetResources.GetSheet(parsedSheetId1, null, null, null, null, null, null, null);
            Sheet sheet2 = smartsheet.SheetResources.GetSheet(parsedSheetId2, null, null, null, null, null, null, null);
            //Sheet sheet3 = smartsheet.SheetResources.GetSheet(parsedSheetId3, null, null, null, null, null, null, null);
            Sheet sheet4 = smartsheet.SheetResources.GetSheet(parsedSheetId4, null, null, null, null, null, null, null);
            //Sheet sheet5 = smartsheet.SheetResources.GetSheet(parsedSheetId5, null, null, null, null, null, null, null);
            Sheet sheet6 = smartsheet.SheetResources.GetSheet(parsedSheetId6, null, null, null, null, null, null, null);
            Sheet sheet7 = smartsheet.SheetResources.GetSheet(parsedSheetId7, null, null, null, null, null, null, null);

            StringBuilder addedBrandsData = new StringBuilder();
            //StringBuilder addedInviteesData = new StringBuilder();
            StringBuilder addedHcpData = new StringBuilder();
            //StringBuilder addedSlideKitData = new StringBuilder();
            StringBuilder addedExpences = new StringBuilder();

            // int addedSlideKitDataNo = 1;
            int addedHcpDataNo = 1;
            //  int addedInviteesDataNo = 1;
            int addedBrandsDataNo = 1;
            int addedExpencesNo = 1;

            //var TotalHonorariumAmount = 0;
            //var TotalTravelAmount = 0;
            //var TotalAccomodateAmount = 0;
            //var TotalHCPLcAmount = 0;
            //var TotalInviteesLcAmount = 0;
            var TotalExpenseAmount = 0;

            CultureInfo hindi = new CultureInfo("hi-IN");


            var EventOpen30Days = "";
            var EventWithin7Days = "";
            var UploadDeviationFile = "";
            var FCPA = "";
            //var InvoiceUpload = "";
            if (formDataList.MedicalUtilityData.EventWithin7daysFile != "")
            {
                EventWithin7Days = "Yes";
            }
            else
            {
                EventWithin7Days = "No";
            }
            if (formDataList.MedicalUtilityData.EventOpen30daysFile != "")
            {
                EventOpen30Days = "Yes";
            }
            else
            {
                EventOpen30Days = "No";
            }
            if (formDataList.MedicalUtilityData.UploadDeviationFile != "")
            {
                UploadDeviationFile = "Yes";
            }
            else
            {
                UploadDeviationFile = "No";
            }




            foreach (var formdata in formDataList.ExpenseSheet)
            {
                string rowData = $"{addedExpencesNo}. {formdata.Expense} | TotalAmount: {formdata.TotalExpenseAmount}| {formdata.BTC_BTE}";
                addedExpences.AppendLine(rowData);
                addedExpencesNo++;
                var amount = int.Parse(formdata.TotalExpenseAmount);
                TotalExpenseAmount = TotalExpenseAmount + amount;

            }
            string Expense = addedExpences.ToString();



            foreach (var formdata in formDataList.BrandsList)
            {

                string rowData = $"{addedBrandsDataNo}. {formdata.BrandName} | {formdata.ProjectId} | {formdata.PercentAllocation}";
                addedBrandsData.AppendLine(rowData);
                addedBrandsDataNo++;
            }
            string brand = addedBrandsData.ToString();




            foreach (var formdata in formDataList.HcpList)
            {
                string rowData = $"{addedHcpDataNo}. {formdata.MisCode} |{formdata.HcpName} |Speciality: {formdata.Speciality} |Tier: {formdata.Tier} ";

                addedHcpData.AppendLine(rowData);
                addedHcpDataNo++;

            }
            string HCP = addedHcpData.ToString();



            var total = TotalExpenseAmount;






            try
            {

                var newRow = new Row();
                newRow.Cells = new List<Cell>();
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "Event Topic"),
                    Value = formDataList.MedicalUtilityData.EventTopic
                });

                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "EventType"),
                    Value = formDataList.MedicalUtilityData.EventType
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "EventDate"),
                    Value = formDataList.MedicalUtilityData.EventDate
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "Valid From"),
                    Value = formDataList.MedicalUtilityData.ValidFrom
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "Valid To"),
                    Value = formDataList.MedicalUtilityData.ValidTill
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "Medical Utility Type"),
                    Value = formDataList.MedicalUtilityData.MedicalUtilityType
                });

                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "Medical Utility Description"),
                    Value = formDataList.MedicalUtilityData.MedicalUtilityDescription
                });

                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "IsAdvanceRequired"),
                    Value = formDataList.MedicalUtilityData.IsAdvanceRequired
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
                    ColumnId = GetColumnIdByName(sheet1, "Role"),
                    Value = formDataList.MedicalUtilityData.Role
                });


                //newRow.Cells.Add(new Cell
                //{
                //    ColumnId = GetColumnIdByName(sheet1, "EventOpen30days"),
                //    Value = formDataList.HcpConsultant.EventOpen30days
                //});
                //newRow.Cells.Add(new Cell
                //{
                //    ColumnId = GetColumnIdByName(sheet1, "EventWithin7days"),
                //    Value = formDataList.HcpConsultant.EventWithin7days
                //});


                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "RBM/BM"),
                    Value = formDataList.MedicalUtilityData.RBMorBM
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "Sales Head"),
                    Value = formDataList.MedicalUtilityData.Sales_Head
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "Marketing Head"),
                    Value = formDataList.MedicalUtilityData.Marketing_Head
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "Finance Treasury"),
                    Value = formDataList.MedicalUtilityData.Finance
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "InitiatorName"),
                    Value = formDataList.MedicalUtilityData.InitiatorName
                });
                newRow.Cells.Add(new Cell
                {
                    ColumnId = GetColumnIdByName(sheet1, "Initiator Email"),
                    Value = formDataList.MedicalUtilityData.Initiator_Email
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






                if (formDataList.MedicalUtilityData.EventOpen30daysFile == "Yes" || formDataList.MedicalUtilityData.EventWithin7daysFile == "Yes" || formDataList.MedicalUtilityData.IsDeviationUpload == "Yes")
                {
                    var eventId = val;
                    try
                    {

                        var newRow7 = new Row();
                        newRow7.Cells = new List<Cell>();

                        newRow7.Cells.Add(new Cell
                        {
                            ColumnId = GetColumnIdByName(sheet7, "EventId/EventRequestId"),
                            Value = eventId
                        });
                        newRow7.Cells.Add(new Cell
                        {
                            ColumnId = GetColumnIdByName(sheet7, "Event Topic"),
                            Value = formDataList.MedicalUtilityData.EventTopic
                        });

                        newRow7.Cells.Add(new Cell
                        {
                            ColumnId = GetColumnIdByName(sheet7, "EventType"),
                            Value = formDataList.MedicalUtilityData.EventType
                        });
                        newRow7.Cells.Add(new Cell
                        {
                            ColumnId = GetColumnIdByName(sheet7, "EventDate"),
                            Value = formDataList.MedicalUtilityData.EventDate
                        });

                        newRow7.Cells.Add(new Cell
                        {
                            ColumnId = GetColumnIdByName(sheet7, "EventOpen30days"),
                            Value = formDataList.MedicalUtilityData.EventOpen30daysFile
                        });
                        newRow7.Cells.Add(new Cell
                        {
                            ColumnId = GetColumnIdByName(sheet7, "EventWithin7days"),
                            Value = formDataList.MedicalUtilityData.EventWithin7daysFile
                        });
                        newRow7.Cells.Add(new Cell
                        {
                            ColumnId = GetColumnIdByName(sheet7, "HCP exceeds 1,00,000 Trigger"),
                            Value = formDataList.MedicalUtilityData.UploadDeviationFile
                        });

                        newRow7.Cells.Add(new Cell
                        {
                            ColumnId = GetColumnIdByName(sheet7, "Sales Head"),
                            Value = formDataList.MedicalUtilityData.Sales_Head
                        });
                        newRow7.Cells.Add(new Cell
                        {
                            ColumnId = GetColumnIdByName(sheet7, "Finance Head"),
                            Value = formDataList.MedicalUtilityData.Sales_Head
                        });

                        newRow7.Cells.Add(new Cell
                        {
                            ColumnId = GetColumnIdByName(sheet7, "InitiatorName"),
                            Value = formDataList.MedicalUtilityData.InitiatorName
                        });
                        newRow7.Cells.Add(new Cell
                        {
                            ColumnId = GetColumnIdByName(sheet7, "Initiator Email"),
                            Value = formDataList.MedicalUtilityData.Initiator_Email
                        });


                        var addeddeviationrow = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId7, new Row[] { newRow7 });

                        if (EventWithin7Days == "Yes")
                        {
                            byte[] fileBytes = Convert.FromBase64String(formDataList.MedicalUtilityData.EventWithin7daysFile);
                            var folderName = Path.Combine("Resources", "Images");
                            var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                            if (!Directory.Exists(pathToSave))
                            {
                                Directory.CreateDirectory(pathToSave);
                            }

                            string fileType = GetFileType(fileBytes);
                            string fileName = val + "-" + " 7DaysDeviation." + fileType;
                            // string fileName = val+x + ": AttachedFile." + fileType;
                            string filePath = Path.Combine(pathToSave, fileName);


                            var addedRow = addeddeviationrow[0];

                            System.IO.File.WriteAllBytes(filePath, fileBytes);
                            string type = GetContentType(fileType);
                            var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                    parsedSheetId7, addedRow.Id.Value, filePath, "application/msword");

                            if (System.IO.File.Exists(filePath))
                            {
                                System.IO.File.Delete(filePath);
                            }
                        }
                        if (EventOpen30Days == "Yes")
                        {
                            byte[] fileBytes = Convert.FromBase64String(formDataList.MedicalUtilityData.EventOpen30daysFile);
                            var folderName = Path.Combine("Resources", "Images");
                            var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                            if (!Directory.Exists(pathToSave))
                            {
                                Directory.CreateDirectory(pathToSave);
                            }

                            string fileType = GetFileType(fileBytes);
                            string fileName = val + "-" + " 30DaysDeviation." + fileType;
                            // string fileName = val+x + ": AttachedFile." + fileType;
                            string filePath = Path.Combine(pathToSave, fileName);


                            var addedRow = addeddeviationrow[0];

                            System.IO.File.WriteAllBytes(filePath, fileBytes);
                            string type = GetContentType(fileType);
                            var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                    parsedSheetId7, addedRow.Id.Value, filePath, "application/msword");

                            if (System.IO.File.Exists(filePath))
                            {
                                System.IO.File.Delete(filePath);
                            }
                        }
                        if (UploadDeviationFile == "Yes")
                        {
                            byte[] fileBytes = Convert.FromBase64String(formDataList.MedicalUtilityData.UploadDeviationFile);
                            var folderName = Path.Combine("Resources", "Images");
                            var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                            if (!Directory.Exists(pathToSave))
                            {
                                Directory.CreateDirectory(pathToSave);
                            }

                            string fileType = GetFileType(fileBytes);
                            string fileName = val + "-" + " UploadDeviationFile." + fileType;
                            // string fileName = val+x + ": AttachedFile." + fileType;
                            string filePath = Path.Combine(pathToSave, fileName);


                            var addedRow = addeddeviationrow[0];

                            System.IO.File.WriteAllBytes(filePath, fileBytes);
                            string type = GetContentType(fileType);
                            var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                    parsedSheetId7, addedRow.Id.Value, filePath, "application/msword");
                            if (System.IO.File.Exists(filePath))
                            {
                                System.IO.File.Delete(filePath);
                            }
                        }



                    }
                    catch (Exception ex)
                    {
                        return BadRequest(ex.Message);
                    }
                }





                foreach (var formData in formDataList.HcpList)
                {
                    var newRow1 = new Row();
                    newRow1.Cells = new List<Cell>();
                    newRow1.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet4, "HCPName"),
                        Value = formData.HcpName
                    });


                    newRow1.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet4, "MISCode"),
                        Value = formData.MisCode
                    });
                    newRow1.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet4, "HCP Type"),
                        Value = formData.HcpType
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
                        ColumnId = GetColumnIdByName(sheet4, "Medical Utility Cost"),
                        Value = formData.MedicalUtilityCostAmount
                    });
                    newRow1.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet4, "Medical Utility Type"),
                        Value = formDataList.MedicalUtilityData.MedicalUtilityType
                    });

                    newRow1.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet4, "Legitimate Need"),
                        Value = formData.Legitimate
                    });
                    newRow1.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet4, "Objective Criteria"),
                        Value = formData.Objective
                    });
                    newRow1.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet4, "Rationale"),
                        Value = formData.Rationale
                    });
                    newRow1.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet4, "FCPA Date"),
                        Value = formData.Fcpadate
                    });


                    newRow1.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet4, "EventId/EventRequestId"),
                        Value = val
                    });


                    var addeddatarows = smartsheet.SheetResources.RowResources.AddRows(parsedSheetId4, new Row[] { newRow1 });

                    var FCPAFile = "";
                    var UploadWrittenRequestDate = "";
                    var UploadHCPRequestDate = "";
                    var Invoice_Brouchere_Quotation = "";

                    if (formData.UploadFCPA != "")
                    {
                        FCPAFile = "Yes";
                    }
                    else
                    {
                        FCPAFile = "No";
                    }
                    if (formData.UploadWrittenRequestDate != "")
                    {
                        UploadWrittenRequestDate = "Yes";
                    }
                    else
                    {
                        UploadWrittenRequestDate = "No";
                    }
                    //if (formData.UploadHCPRequestDate != "")
                    //{
                    //    UploadHCPRequestDate = "Yes";
                    //}
                    //else
                    //{
                    //    UploadHCPRequestDate = "No";
                    //}
                    if (formData.Invoice_Brouchere_Quotation != "")
                    {
                        Invoice_Brouchere_Quotation = "Yes";
                    }
                    else
                    {
                        Invoice_Brouchere_Quotation = "No";
                    }



                    var columnId = GetColumnIdByName(sheet1, "EventId/EventRequestId");
                    var Cell = addedRows[0].Cells.FirstOrDefault(cell => cell.ColumnId == columnId);
                    var value = Cell.DisplayValue;

                    if (FCPAFile == "Yes")
                    {
                        byte[] fileBytes = Convert.FromBase64String(formData.UploadFCPA);
                        var folderName = Path.Combine("Resources", "Images");
                        var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                        if (!Directory.Exists(pathToSave))
                        {
                            Directory.CreateDirectory(pathToSave);
                        }

                        string fileType = GetFileType(fileBytes);
                        string fileName = value + "-" + " FCPA." + fileType;
                        // string fileName = val+x + ": AttachedFile." + fileType;
                        string filePath = Path.Combine(pathToSave, fileName);


                        var addedRow = addeddatarows[0];

                        System.IO.File.WriteAllBytes(filePath, fileBytes);
                        string type = GetContentType(fileType);
                        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                parsedSheetId4, addedRow.Id.Value, filePath, "application/msword");
                        if (System.IO.File.Exists(filePath))
                        {
                            System.IO.File.Delete(filePath);
                        }
                    }
                    if (UploadWrittenRequestDate == "Yes")
                    {
                        byte[] fileBytes = Convert.FromBase64String(formData.UploadWrittenRequestDate);
                        var folderName = Path.Combine("Resources", "Images");
                        var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                        if (!Directory.Exists(pathToSave))
                        {
                            Directory.CreateDirectory(pathToSave);
                        }

                        string fileType = GetFileType(fileBytes);
                        string fileName = value + "-" + " UploadWrittenRequestDate." + fileType;
                        // string fileName = val+x + ": AttachedFile." + fileType;
                        string filePath = Path.Combine(pathToSave, fileName);


                        var addedRow = addeddatarows[0];

                        System.IO.File.WriteAllBytes(filePath, fileBytes);
                        string type = GetContentType(fileType);
                        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                parsedSheetId4, addedRow.Id.Value, filePath, "application/msword");
                        if (System.IO.File.Exists(filePath))
                        {
                            System.IO.File.Delete(filePath);
                        }
                    }
                    //if (UploadHCPRequestDate == "Yes")
                    //{
                    //    byte[] fileBytes = Convert.FromBase64String(formData.UploadHCPRequestDate);
                    //    var folderName = Path.Combine("Resources", "Images");
                    //    var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                    //    if (!Directory.Exists(pathToSave))
                    //    {
                    //        Directory.CreateDirectory(pathToSave);
                    //    }

                    //    string fileType = GetFileType(fileBytes);
                    //    string fileName = value + "-" + " UploadHCPRequestDate." + fileType;
                    //    // string fileName = val+x + ": AttachedFile." + fileType;
                    //    string filePath = Path.Combine(pathToSave, fileName);


                    //    var addedRow = addeddatarows[0];

                    //    System.IO.File.WriteAllBytes(filePath, fileBytes);
                    //    string type = GetContentType(fileType);
                    //    var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                    //            parsedSheetId4, addedRow.Id.Value, filePath, "application/msword");
                    //    if (System.IO.File.Exists(filePath))
                    //    {
                    //        System.IO.File.Delete(filePath);
                    //    }
                    //}
                    if (Invoice_Brouchere_Quotation == "Yes")
                    {
                        byte[] fileBytes = Convert.FromBase64String(formData.Invoice_Brouchere_Quotation);
                        var folderName = Path.Combine("Resources", "Images");
                        var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                        if (!Directory.Exists(pathToSave))
                        {
                            Directory.CreateDirectory(pathToSave);
                        }

                        string fileType = GetFileType(fileBytes);
                        string fileName = value + "-" + " Invoice_Brouchere_Quotation." + fileType;
                        // string fileName = val+x + ": AttachedFile." + fileType;
                        string filePath = Path.Combine(pathToSave, fileName);


                        var addedRow = addeddatarows[0];

                        System.IO.File.WriteAllBytes(filePath, fileBytes);
                        string type = GetContentType(fileType);
                        var attachment = smartsheet.SheetResources.RowResources.AttachmentResources.AttachFile(
                                parsedSheetId4, addedRow.Id.Value, filePath, "application/msword");

                        if (System.IO.File.Exists(filePath))
                        {
                            System.IO.File.Delete(filePath);
                        }
                    }





                }

                foreach (var formdata in formDataList.BrandsList)
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


                foreach (var formdata in formDataList.ExpenseSheet)
                {
                    var newRow6 = new Row();
                    newRow6.Cells = new List<Cell>();
                    newRow6.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet6, "MisCode"),
                        Value = formdata.Expense
                    });

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
                        ColumnId = GetColumnIdByName(sheet6, "BTC/BTE"),
                        Value = formdata.BTC_BTE
                    });
                    newRow6.Cells.Add(new Cell
                    {
                        ColumnId = GetColumnIdByName(sheet6, "Amount"),
                        Value = formdata.TotalExpenseAmount
                    });

                    smartsheet.SheetResources.RowResources.AddRows(parsedSheetId6, new Row[] { newRow6 });
                }

                return Ok(new
                { Message = " Success!" });
            }

            catch (Exception ex)
            {
                return BadRequest($"Could not find {ex.Message}");
            }
        }


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
