using Aspose.Pdf.Plugins;
using IndiaEvents.Models.Models.EventTypeSheets;
using IndiaEventsWebApi.Helper;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartsheet.Api;
using Smartsheet.Api.Models;

namespace IndiaEventsWebApi.Controllers.EventsController
{
    [Route("api/[controller]")]
    [ApiController]
    public class ApprovalAndRejectionFlowController : ControllerBase
    {
        private readonly string accessToken;
        private readonly IConfiguration configuration;
        private readonly SmartsheetClient smartsheet;
        private readonly string sheetId1;
        private readonly string sheetId2;
        private readonly string sheetId3;
        private readonly string sheetId4;
        private readonly string sheetId5;
        private readonly string sheetId6;
        private readonly string sheetId7;
        private readonly string sheetId8;
        private readonly string sheetId9;

        private readonly string sheetId10;
        private readonly string sheetId11;
        private readonly string sheetId12;
        private readonly string sheetId13;
        private readonly string sheetId14;
        private readonly string sheetId15;

        public ApprovalAndRejectionFlowController(IConfiguration configuration)
        {
            this.configuration = configuration;
            accessToken = configuration.GetSection("SmartsheetSettings:AccessToken").Value;
            smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

            sheetId1 = configuration.GetSection("SmartsheetSettings:EventRequestProcess").Value;
            sheetId2 = configuration.GetSection("SmartsheetSettings:Deviation_Process").Value;
            sheetId3 = configuration.GetSection("SmartsheetSettings:HonorariumPayment").Value;
            sheetId4 = configuration.GetSection("SmartsheetSettings:EventSettlement").Value;
            sheetId5 = configuration.GetSection("SmartsheetSettings:EventRequestsHcpSlideKit").Value;
            sheetId6 = configuration.GetSection("SmartsheetSettings:EventRequestsExpensesSheet").Value;
            sheetId8 = configuration.GetSection("SmartsheetSettings:EventRequestBeneficiary").Value;
            sheetId9 = configuration.GetSection("SmartsheetSettings:EventRequestProductBrandsList").Value;

            sheetId10 = configuration.GetSection("SmartsheetSettings:ApprovedSpeakers").Value;
            sheetId11 = configuration.GetSection("SmartsheetSettings:ApprovedTrainers").Value;
            sheetId12 = configuration.GetSection("SmartsheetSettings:VendorMasterSheet").Value;
            sheetId13 = configuration.GetSection("SmartsheetSettings:SpeakerCodeCreation").Value;
            sheetId14 = configuration.GetSection("SmartsheetSettings:TrainerCodeCreation").Value;
            sheetId15 = configuration.GetSection("SmartsheetSettings:VendorCodeCreation").Value;
        }

        [HttpPut("ApprovalAndRejectionFlowInPreEvent")]
        public IActionResult ApprovalAndRejectionFlowInPreEvent(ApprovalAndRejectionFlowInPreEvent formDataList)
        {
            Discussion discussionSpecification = new Discussion
            {
                Comment = new Comment
                {
                    Text = formDataList.Comments
                },
                Comments = null
            };
            Comment commentSpecification = new Comment
            {
                Text = formDataList.Comments
            };

            var EventId = formDataList.EventId;
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId1);
            Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == EventId));

            if (targetRow != null)
            {
                try
                {
                    Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                    if (!string.IsNullOrEmpty(formDataList.RBMStatus))
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "PRE-RBM/BM Approval"), Value = formDataList.RBMStatus });
                    }
                    else if (!string.IsNullOrEmpty(formDataList.SalesHeadStatus))
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "PRE-Sales Head Approval"), Value = formDataList.SalesHeadStatus });
                    }
                    else if (!string.IsNullOrEmpty(formDataList.MarketingHeadStatus))
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "PRE-Marketing Head Approval"), Value = formDataList.MarketingHeadStatus });
                    }
                    else if (!string.IsNullOrEmpty(formDataList.FinanceTreasuryStatus))
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "PRE-Finance Treasury Approval"), Value = formDataList.FinanceTreasuryStatus });
                    }
                    else if (!string.IsNullOrEmpty(formDataList.MedicalAffairsHeadStatus))
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "PRE-Medical Affairs Head Approval"), Value = formDataList.MedicalAffairsHeadStatus });
                    }
                    else if (!string.IsNullOrEmpty(formDataList.ComplianceStatus))
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "PRE-Compliance Approval"), Value = formDataList.ComplianceStatus });
                    }

                    IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, new Row[] { updateRow });
                    if (!string.IsNullOrEmpty(formDataList.Comments))
                    {
                        if (targetRow.Discussions != null)
                        {
                            Comment newComment = smartsheet.SheetResources.DiscussionResources.CommentResources.AddComment(sheet.Id.Value, targetRow.Discussions[0].Id.Value, commentSpecification);
                        }
                        else
                        {
                            Discussion newDiscussion = smartsheet.SheetResources.RowResources.DiscussionResources.CreateDiscussion(sheet.Id.Value, targetRow.Id.Value, discussionSpecification);
                        }
                    }



                    return Ok(new { Message = "Updated Successfully" });
                }
                catch (Exception ex)
                {
                    return BadRequest(ex.Message);
                }
            }
            else
            {
                return BadRequest(new { Message = "Row data not found" });
            }

            //return Ok();

        }

        [HttpPut("ApprovalAndRejectionFlowInDeviation")]
        public IActionResult ApprovalAndRejectionFlowInDeviation(ApprovalAndRejectionFlowInDeviation formDataList)
        {
            Discussion discussionSpecification = new Discussion
            {
                Comment = new Comment
                {
                    Text = formDataList.Comments
                },
                Comments = null
            };
            Comment commentSpecification = new Comment
            {
                Text = formDataList.Comments
            };

            var EventId = formDataList.EventId;
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId2);
            Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == EventId));

            if (targetRow != null)
            {
                try
                {
                    Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                    if (!string.IsNullOrEmpty(formDataList.FinanceHeadStatus))
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance Head Approval"), Value = formDataList.FinanceHeadStatus });
                    }
                    else if (!string.IsNullOrEmpty(formDataList.SalesHeadStatus))
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Sales Head approval"), Value = formDataList.SalesHeadStatus });
                    }


                    IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, new Row[] { updateRow });
                    if (!string.IsNullOrEmpty(formDataList.Comments))
                    {
                        if (targetRow.Discussions != null)
                        {
                            Comment newComment = smartsheet.SheetResources.DiscussionResources.CommentResources.AddComment(sheet.Id.Value, targetRow.Discussions[0].Id.Value, commentSpecification);
                        }
                        else
                        {
                            Discussion newDiscussion = smartsheet.SheetResources.RowResources.DiscussionResources.CreateDiscussion(sheet.Id.Value, targetRow.Id.Value, discussionSpecification);
                        }
                    }



                    return Ok(new { Message = "Updated Successfully" });
                }
                catch (Exception ex)
                {
                    return BadRequest(ex.Message);
                }
            }
            else
            {
                return BadRequest(new { Message = "Row data not found" });
            }

            //return Ok();

        }

        [HttpPut("ApprovalAndRejectionFlowInHonorarium")]
        public IActionResult ApprovalAndRejectionFlowInHonorarium(ApprovalAndRejectionFlowInHonorarium formDataList)
        {
            Discussion discussionSpecification = new Discussion
            {
                Comment = new Comment
                {
                    Text = formDataList.Comments
                },
                Comments = null
            };
            Comment commentSpecification = new Comment
            {
                Text = formDataList.Comments
            };

            var EventId = formDataList.EventId;
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId3);
            Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == EventId));

            if (targetRow != null)
            {
                try
                {
                    Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                    if (!string.IsNullOrEmpty(formDataList.RBMStatus))
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "HON-RBM/BM Approval"), Value = formDataList.RBMStatus });
                    }
                    else if (!string.IsNullOrEmpty(formDataList.SalesHeadStatus))
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "HON-Sales Head Approval"), Value = formDataList.SalesHeadStatus });
                    }
                    else if (!string.IsNullOrEmpty(formDataList.MarketingHeadStatus))
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "HON-Marketing Head Approval"), Value = formDataList.MarketingHeadStatus });
                    }

                    else if (!string.IsNullOrEmpty(formDataList.MedicalAffairsHeadStatus))
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "HON-Medical Affairs Head Approval"), Value = formDataList.MedicalAffairsHeadStatus });
                    }
                    else if (!string.IsNullOrEmpty(formDataList.ComplianceStatus))
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "HON-Compliance Approval"), Value = formDataList.ComplianceStatus });
                    }

                    IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, new Row[] { updateRow });
                    if (!string.IsNullOrEmpty(formDataList.Comments))
                    {
                        if (targetRow.Discussions != null)
                        {
                            Comment newComment = smartsheet.SheetResources.DiscussionResources.CommentResources.AddComment(sheet.Id.Value, targetRow.Discussions[0].Id.Value, commentSpecification);
                        }
                        else
                        {
                            Discussion newDiscussion = smartsheet.SheetResources.RowResources.DiscussionResources.CreateDiscussion(sheet.Id.Value, targetRow.Id.Value, discussionSpecification);
                        }
                    }



                    return Ok(new { Message = "Updated Successfully" });
                }
                catch (Exception ex)
                {
                    return BadRequest(ex.Message);
                }
            }
            else
            {
                return BadRequest(new { Message = "Row data not found" });
            }

            //return Ok();

        }

        [HttpPut("ApprovalAndRejectionFlowInEventSettlement")]
        public IActionResult ApprovalAndRejectionFlowInEventSettlement(ApprovalAndRejectionFlowInPostSettlement formDataList)
        {
            Discussion discussionSpecification = new Discussion
            {
                Comment = new Comment
                {
                    Text = formDataList.Comments
                },
                Comments = null
            };
            Comment commentSpecification = new Comment
            {
                Text = formDataList.Comments
            };

            var EventId = formDataList.EventId;
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId4);
            Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == EventId));

            if (targetRow != null)
            {
                try
                {
                    Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                    if (!string.IsNullOrEmpty(formDataList.RBMStatus))
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventSettlement-RBM/BM Approval"), Value = formDataList.RBMStatus });
                    }
                    else if (!string.IsNullOrEmpty(formDataList.SalesHeadStatus))
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventSettlement-SalesHead Approval"), Value = formDataList.SalesHeadStatus });
                    }
                    else if (!string.IsNullOrEmpty(formDataList.MarketingHeadStatus))
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventSettlement-Marketing Head Approval"), Value = formDataList.MarketingHeadStatus });
                    }

                    else if (!string.IsNullOrEmpty(formDataList.MedicalAffairsHeadStatus))
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventSettlement-Medical Affairs Head Approval"), Value = formDataList.MedicalAffairsHeadStatus });
                    }
                    else if (!string.IsNullOrEmpty(formDataList.ComplianceStatus))
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "EventSettlement-Compliance Approval"), Value = formDataList.ComplianceStatus });
                    }

                    IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, new Row[] { updateRow });
                    if (!string.IsNullOrEmpty(formDataList.Comments))
                    {
                        if (targetRow.Discussions != null)
                        {
                            Comment newComment = smartsheet.SheetResources.DiscussionResources.CommentResources.AddComment(sheet.Id.Value, targetRow.Discussions[0].Id.Value, commentSpecification);
                        }
                        else
                        {
                            Discussion newDiscussion = smartsheet.SheetResources.RowResources.DiscussionResources.CreateDiscussion(sheet.Id.Value, targetRow.Id.Value, discussionSpecification);
                        }
                    }



                    return Ok(new { Message = "Updated Successfully" });
                }
                catch (Exception ex)
                {
                    return BadRequest(ex.Message);
                }
            }
            else
            {
                return BadRequest(new { Message = "Row data not found" });
            }

            //return Ok();

        }

        [HttpPut("ComplianceHeadCancellation")]
        public IActionResult ComplianceHeadCancellation(ComplianceRejectionFlow formDataList)
        {
            var EventId = formDataList.EventId;
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId1);
            Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == EventId));

            if (targetRow != null)
            {
                try
                {
                    Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };


                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Event Cancelled"), Value = formDataList.IsComplianceCancelledEvent });
                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Reason for Cancellation"), Value = formDataList.Comments });
                    IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, new Row[] { updateRow });
                    return Ok(new { Message = "Updated Successfully" });
                }
                catch (Exception ex)
                {
                    return BadRequest(ex.Message);
                }
            }
            else
            {
                return BadRequest(new { Message = "Row data not found" });
            }




            return Ok();
        }

        [HttpPut("PreEventFinanceTreasuryApproval")]
        public IActionResult PreEventFinanceTreasuryApproval(PreEventFinanceTreasuryApprovalFlow formDataList)
        {
            Discussion discussionSpecification = new Discussion
            {
                Comment = new Comment
                {
                    Text = formDataList.Comments
                },
                Comments = null
            };
            Comment commentSpecification = new Comment
            {
                Text = formDataList.Comments
            };

            var EventId = formDataList.EventId;
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId1);
            Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == EventId));

            if (targetRow != null)
            {
                try
                {
                    Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };

                    updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "PRE-Finance Treasury Approval"), Value = formDataList.PreFinanceTreasuryApprovalStatus });
                    if (formDataList.PreFinanceTreasuryApprovalStatus.ToLower() == "approved")
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Bank Reference Date"), Value = formDataList.BankReferenceDate });
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Advance Voucher Date"), Value = formDataList.AdvanceVoucherDate });
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Bank Reference Number"), Value = formDataList.BankReferenceNumber });
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Advance Voucher Number"), Value = formDataList.AdvanceVoucherNumber });

                    }




                    IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, new Row[] { updateRow });
                    if (!string.IsNullOrEmpty(formDataList.Comments))
                    {
                        if (targetRow.Discussions != null)
                        {
                            Comment newComment = smartsheet.SheetResources.DiscussionResources.CommentResources.AddComment(sheet.Id.Value, targetRow.Discussions[0].Id.Value, commentSpecification);
                        }
                        else
                        {
                            Discussion newDiscussion = smartsheet.SheetResources.RowResources.DiscussionResources.CreateDiscussion(sheet.Id.Value, targetRow.Id.Value, discussionSpecification);
                        }
                    }



                    return Ok(new { Message = "Updated Successfully" });
                }
                catch (Exception ex)
                {
                    return BadRequest(ex.Message);
                }
            }
            else
            {
                return BadRequest(new { Message = "Row data not found" });
            }
        }



        [HttpPut("ApprovedSpeakersApproval")]
        public IActionResult ApprovedSpeakersApproval(ApprovalAndRejectionFlowInApprovedSpeakers formDataList)
        {
            Discussion discussionSpecification = new()
            {
                Comment = new Comment
                {
                    Text = formDataList.Comments
                },
                Comments = null
            };
            Comment commentSpecification = new()
            {
                Text = formDataList.Comments
            };

            var EventId = formDataList.Id;
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId10);
            Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == EventId));

            if (targetRow != null)
            {
                try
                {
                    Row updateRow = new() { Id = targetRow.Id, Cells = new List<Cell>() };
                    if (!string.IsNullOrEmpty(formDataList.MedicalAffairsHeadStatus))
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Medical Affairs Head Approval"), Value = formDataList.MedicalAffairsHeadStatus });
                    }
                    else if (!string.IsNullOrEmpty(formDataList.SalesHeadStatus))
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Sales Head Approval"), Value = formDataList.SalesHeadStatus });
                    }

                    IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, new Row[] { updateRow });
                    if (!string.IsNullOrEmpty(formDataList.Comments))
                    {
                        if (targetRow.Discussions != null)
                        {
                            Comment newComment = smartsheet.SheetResources.DiscussionResources.CommentResources.AddComment(sheet.Id.Value, targetRow.Discussions[0].Id.Value, commentSpecification);
                        }
                        else
                        {
                            Discussion newDiscussion = smartsheet.SheetResources.RowResources.DiscussionResources.CreateDiscussion(sheet.Id.Value, targetRow.Id.Value, discussionSpecification);
                        }
                    }



                    return Ok(new { Message = "Updated Successfully" });
                }
                catch (Exception ex)
                {
                    return BadRequest(ex.Message);
                }
            }
            else
            {
                return BadRequest(new { Message = "Row data not found" });
            }
        }

        [HttpPut("ApprovedTrainerApproval")]
        public IActionResult ApprovedTrainerApproval(ApprovalAndRejectionFlowInApprovedSpeakers formDataList)
        {
            Discussion discussionSpecification = new Discussion
            {
                Comment = new Comment
                {
                    Text = formDataList.Comments
                },
                Comments = null
            };
            Comment commentSpecification = new Comment
            {
                Text = formDataList.Comments
            };

            var EventId = formDataList.Id;
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId11);
            Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == EventId));

            if (targetRow != null)
            {
                try
                {
                    Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                    if (!string.IsNullOrEmpty(formDataList.MedicalAffairsHeadStatus))
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Medical Affairs Head Approval"), Value = formDataList.MedicalAffairsHeadStatus });
                    }
                    else if (!string.IsNullOrEmpty(formDataList.SalesHeadStatus))
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Sales Head Approval"), Value = formDataList.SalesHeadStatus });
                    }

                    IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, new Row[] { updateRow });
                    if (!string.IsNullOrEmpty(formDataList.Comments))
                    {
                        if (targetRow.Discussions != null)
                        {
                            Comment newComment = smartsheet.SheetResources.DiscussionResources.CommentResources.AddComment(sheet.Id.Value, targetRow.Discussions[0].Id.Value, commentSpecification);
                        }
                        else
                        {
                            Discussion newDiscussion = smartsheet.SheetResources.RowResources.DiscussionResources.CreateDiscussion(sheet.Id.Value, targetRow.Id.Value, discussionSpecification);
                        }
                    }



                    return Ok(new { Message = "Updated Successfully" });
                }
                catch (Exception ex)
                {
                    return BadRequest(ex.Message);
                }
            }
            else
            {
                return BadRequest(new { Message = "Row data not found" });
            }
        }

        [HttpPut("SpeakersCodeCreationApproval")]
        public IActionResult SpeakersCodeCreationApproval(ApprovalAndRejectionFlowInApprovedSpeakers formDataList)
        {
            Discussion discussionSpecification = new Discussion
            {
                Comment = new Comment
                {
                    Text = formDataList.Comments
                },
                Comments = null
            };
            Comment commentSpecification = new Comment
            {
                Text = formDataList.Comments
            };

            var EventId = formDataList.Id;
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId13);
            Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == EventId));

            if (targetRow != null)
            {
                try
                {
                    Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                    if (!string.IsNullOrEmpty(formDataList.MedicalAffairsHeadStatus))
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Medical Affairs Head Approval"), Value = formDataList.MedicalAffairsHeadStatus });
                    }
                    else if (!string.IsNullOrEmpty(formDataList.SalesHeadStatus))
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Sales Head Approval"), Value = formDataList.SalesHeadStatus });
                    }

                    IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, new Row[] { updateRow });
                    if (!string.IsNullOrEmpty(formDataList.Comments))
                    {
                        if (targetRow.Discussions != null)
                        {
                            Comment newComment = smartsheet.SheetResources.DiscussionResources.CommentResources.AddComment(sheet.Id.Value, targetRow.Discussions[0].Id.Value, commentSpecification);
                        }
                        else
                        {
                            Discussion newDiscussion = smartsheet.SheetResources.RowResources.DiscussionResources.CreateDiscussion(sheet.Id.Value, targetRow.Id.Value, discussionSpecification);
                        }
                    }



                    return Ok(new { Message = "Updated Successfully" });
                }
                catch (Exception ex)
                {
                    return BadRequest(ex.Message);
                }
            }
            else
            {
                return BadRequest(new { Message = "Row data not found" });
            }
        }

        [HttpPut("TrainerCodeCreationApproval")]
        public IActionResult TrainerCodeCreationApproval(ApprovalAndRejectionFlowInApprovedSpeakers formDataList)
        {
            Discussion discussionSpecification = new Discussion
            {
                Comment = new Comment
                {
                    Text = formDataList.Comments
                },
                Comments = null
            };
            Comment commentSpecification = new Comment
            {
                Text = formDataList.Comments
            };

            var EventId = formDataList.Id;
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId14);
            Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == EventId));

            if (targetRow != null)
            {
                try
                {
                    Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                    if (!string.IsNullOrEmpty(formDataList.MedicalAffairsHeadStatus))
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Medical Affairs Head Approval"), Value = formDataList.MedicalAffairsHeadStatus });
                    }
                    else if (!string.IsNullOrEmpty(formDataList.SalesHeadStatus))
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Sales Head Approval"), Value = formDataList.SalesHeadStatus });
                    }

                    IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, new Row[] { updateRow });
                    if (!string.IsNullOrEmpty(formDataList.Comments))
                    {
                        if (targetRow.Discussions != null)
                        {
                            Comment newComment = smartsheet.SheetResources.DiscussionResources.CommentResources.AddComment(sheet.Id.Value, targetRow.Discussions[0].Id.Value, commentSpecification);
                        }
                        else
                        {
                            Discussion newDiscussion = smartsheet.SheetResources.RowResources.DiscussionResources.CreateDiscussion(sheet.Id.Value, targetRow.Id.Value, discussionSpecification);
                        }
                    }



                    return Ok(new { Message = "Updated Successfully" });
                }
                catch (Exception ex)
                {
                    return BadRequest(ex.Message);
                }
            }
            else
            {
                return BadRequest(new { Message = "Row data not found" });
            }
        }

        [HttpPut("ApprovedVendorApproval")]
        public IActionResult ApprovedVendorApproval(ApprovalAndRejectionFlowInApprovedVendor formDataList)
        {
            Discussion discussionSpecification = new Discussion
            {
                Comment = new Comment
                {
                    Text = formDataList.Comments
                },
                Comments = null
            };
            Comment commentSpecification = new Comment
            {
                Text = formDataList.Comments
            };

            var EventId = formDataList.Id;
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId12);
            Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == EventId));

            if (targetRow != null)
            {
                try
                {
                    Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                    if (!string.IsNullOrEmpty(formDataList.FinanceCheckerStatus))
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance Checker Approval"), Value = formDataList.FinanceCheckerStatus });
                    }

                    IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, new Row[] { updateRow });
                    if (!string.IsNullOrEmpty(formDataList.Comments))
                    {
                        if (targetRow.Discussions != null)
                        {
                            Comment newComment = smartsheet.SheetResources.DiscussionResources.CommentResources.AddComment(sheet.Id.Value, targetRow.Discussions[0].Id.Value, commentSpecification);
                        }
                        else
                        {
                            Discussion newDiscussion = smartsheet.SheetResources.RowResources.DiscussionResources.CreateDiscussion(sheet.Id.Value, targetRow.Id.Value, discussionSpecification);
                        }
                    }



                    return Ok(new { Message = "Updated Successfully" });
                }
                catch (Exception ex)
                {
                    return BadRequest(ex.Message);
                }
            }
            else
            {
                return BadRequest(new { Message = "Row data not found" });
            }
        }

        [HttpPut("VendorCodeCreationApproval")]
        public IActionResult VendorCodeCreationApproval(ApprovalAndRejectionFlowInApprovedVendor formDataList)
        {
            Discussion discussionSpecification = new Discussion
            {
                Comment = new Comment
                {
                    Text = formDataList.Comments
                },
                Comments = null
            };
            Comment commentSpecification = new Comment
            {
                Text = formDataList.Comments
            };

            var EventId = formDataList.Id;
            Sheet sheet = SheetHelper.GetSheetById(smartsheet, sheetId15);
            Row? targetRow = sheet.Rows.FirstOrDefault(r => r.Cells.Any(c => c.DisplayValue == EventId));

            if (targetRow != null)
            {
                try
                {
                    Row updateRow = new Row { Id = targetRow.Id, Cells = new List<Cell>() };
                    if (!string.IsNullOrEmpty(formDataList.FinanceCheckerStatus))
                    {
                        updateRow.Cells.Add(new Cell { ColumnId = SheetHelper.GetColumnIdByName(sheet, "Finance Checker  approval"), Value = formDataList.FinanceCheckerStatus });
                    }


                    IList<Row> updatedRow = smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, new Row[] { updateRow });
                    if (!string.IsNullOrEmpty(formDataList.Comments))
                    {
                        if (targetRow.Discussions != null)
                        {
                            Comment newComment = smartsheet.SheetResources.DiscussionResources.CommentResources.AddComment(sheet.Id.Value, targetRow.Discussions[0].Id.Value, commentSpecification);
                        }
                        else
                        {
                            Discussion newDiscussion = smartsheet.SheetResources.RowResources.DiscussionResources.CreateDiscussion(sheet.Id.Value, targetRow.Id.Value, discussionSpecification);
                        }
                    }



                    return Ok(new { Message = "Updated Successfully" });
                }
                catch (Exception ex)
                {
                    return BadRequest(ex.Message);
                }
            }
            else
            {
                return BadRequest(new { Message = "Row data not found" });
            }
        }








    }
}
