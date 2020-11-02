using Google.Apis.Sheets.v4.Data;
using Google.Sheets.Apis.Utility.Sheets.v4;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json.Linq;
using Server.Dtos;
using System.Collections.Generic;
using System.Linq;

namespace Server.Controllers
{
    /// <summary>
    /// AppendColumnLine、AppendRowLine、AppendRange (o)
    /// UpdateColumnLine、UpdateRowLine、UpdateRange (o)
    /// DeleteColumnLine、DeleteRowLine、DeleteRange (x)
    /// CreateEmptyColumnLines、CreateEmptyRowLines (o)
    /// 搜尋文字
    /// https://developers.google.com/sheets/api/samples/rowcolumn
    /// </summary>
    public class RangesController : ApiControllerBase
    {
        [HttpPost("column")]
        public IActionResult AppendColumnLine(CellColumnDto cellColumn)
        {
            return CheckWorkSheetExist(cellColumn.Sheetname, () =>
            {
                var result = SheetsService.AppendColumnLine(
                    GoogleSheetsApi.SpreadSheetId,
                    new SheetRange(cellColumn.Sheetname, cellColumn.Cellrange),
                    request => { },
                    cellColumn.Line);
                return Ok(result);
            });
        }

        [HttpPost("row")]
        public IActionResult AppendRowLine(CellRowDto cellRow)
        {
            return CheckWorkSheetExist(cellRow.Sheetname, () =>
            {
                var result = SheetsService.AppendRowLine(
                    GoogleSheetsApi.SpreadSheetId,
                    new SheetRange(cellRow.Sheetname, cellRow.Cellrange),
                    request =>
                    {
                        request.InsertDataOption = Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
                    },
                    cellRow.Line);
                return Ok(result);
            });
        }

        [HttpPost("range")]
        public IActionResult AppendRange(CellRangesDto cellRanges)
        {
            return CheckWorkSheetExist(cellRanges.Sheetname, () =>
            {
                var result = SheetsService.AppendMultiLine(
                    GoogleSheetsApi.SpreadSheetId,
                    new SheetRange(cellRanges.Sheetname, cellRanges.Cellrange).ToRange(),
                    request =>
                    {
                        request.InsertDataOption = Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
                    },
                    cellRanges.Values);
                return Ok(result);
            });
        }

        [HttpPut("column")]
        public IActionResult UpdateColumnLine(CellColumnDto cellColumn)
        {
            return CheckWorkSheetExist(cellColumn.Sheetname, () =>
            {
                var result = SheetsService.UpdateColumnLine(
                    GoogleSheetsApi.SpreadSheetId,
                    new SheetRange(cellColumn.Sheetname, cellColumn.Cellrange),
                    request => { },
                    cellColumn.Line);
                return Ok(result);
            });
        }

        [HttpPut("row")]
        public IActionResult UpdateRowLine(CellRowDto cellRow)
        {
            return CheckWorkSheetExist(cellRow.Sheetname, () =>
            {
                var result = SheetsService.UpdateRowLine(
                    GoogleSheetsApi.SpreadSheetId,
                    new SheetRange(cellRow.Sheetname, cellRow.Cellrange),
                    request => { },
                    cellRow.Line);
                return Ok(result);
            });
        }

        [HttpPut("range")]
        public IActionResult UpdateRange(CellRangesDto cellRanges)
        {
            return CheckWorkSheetExist(cellRanges.Sheetname, () =>
            {
                var result = SheetsService.UpdateMultiLine(
                    GoogleSheetsApi.SpreadSheetId,
                    new SheetRange(cellRanges.Sheetname, cellRanges.Cellrange).ToRange(),
                    cellRanges.Values,
                    request => { });
                return Ok(result);
            });
        }
         
        [HttpDelete("column")]
        public IActionResult DeleteColumnLine(CellRangeDeleteDto cellRange)
        {
            cellRange.IsColumn = true;
            return Ok(BatchUpdateSpreadsheet(new CellRangeDeleteDto[] { cellRange }));
        }

        [HttpDelete("row")]
        public IActionResult DeleteRowLine(CellRangeDeleteDto cellRange)
        {
            cellRange.IsColumn = false;
            return Ok(BatchUpdateSpreadsheet(new CellRangeDeleteDto[] { cellRange }));
        }

        [HttpDelete("range")]
        public IActionResult DeleteRangeEmpty(CellRangeDeleteDto[] cellRanges)
        {
            return Ok(BatchUpdateSpreadsheet(cellRanges));
        }

        [HttpPost("column/empty")]
        public IActionResult InsertColumnEmpty(CellRangeEmptyDto cellRange)
        {
            cellRange.IsColumn = true;
            return Created(nameof(InsertColumnEmpty), BatchUpdateSpreadsheet(new CellRangeEmptyDto[] { cellRange })[0]);
        }

        [HttpPost("row/empty")]
        public IActionResult InsertRowEmpty(CellRangeEmptyDto cellRange)
        {
            cellRange.IsColumn = false;
            return Created(nameof(InsertRowEmpty), BatchUpdateSpreadsheet(new CellRangeEmptyDto[] { cellRange })[0]);
        }

        [HttpPost("range/empty")]
        public IActionResult InsertRangeEmpty(CellRangeEmptyDto[] cellRanges)
        {
            return Created(nameof(InsertRangeEmpty), BatchUpdateSpreadsheet(cellRanges));
        }

        private IList<BatchUpdateSpreadsheetResponse> BatchUpdateSpreadsheet(CellRangeEmptyDto[] cellRanges)
        {
            if (cellRanges is null)
            {
                throw new System.ArgumentNullException(nameof(cellRanges));
            }

            return cellRanges.Select(c =>
            {
                return CheckWorkSheetExist(c.Sheetname, () =>
                {
                    var request = CreateInsertDimensionRequest(c);
                    return SheetsService.BatchUpdateSpreadsheet(GoogleSheetsApi.SpreadSheetId, request);
                });
            }).ToList();
        }

        private Request CreateInsertDimensionRequest(CellRangeEmptyDto cellRangeEmpty)
        {
            return new Request
            {
                InsertDimension = new InsertDimensionRequest
                {
                    Range = new DimensionRange
                    {
                        Dimension = cellRangeEmpty.IsColumn ? "COLUMNS" : "ROWS",
                        SheetId = GetSheetId(cellRangeEmpty.Sheetname),
                        StartIndex = cellRangeEmpty.StartIndex,
                        EndIndex = cellRangeEmpty.EndIndex
                    },
                    InheritFromBefore = cellRangeEmpty.InheritFromBefore
                }
            };
        }

        private IList<BatchUpdateSpreadsheetResponse> BatchUpdateSpreadsheet(CellRangeDeleteDto[] cellRanges)
        {
            if (cellRanges is null)
            {
                throw new System.ArgumentNullException(nameof(cellRanges));
            }

            return cellRanges.Select(c =>
            {
                return CheckWorkSheetExist(c.Sheetname, () =>
                {
                    var request = CreateDeleteDimensionRequest(c);
                    return SheetsService.BatchUpdateSpreadsheet(GoogleSheetsApi.SpreadSheetId, request);
                });
            }).ToList();
        }

        private Request CreateDeleteDimensionRequest(CellRangeDeleteDto cellRangeDelete)
        {
            return new Request
            {
                DeleteDimension = new DeleteDimensionRequest
                {
                    Range = new DimensionRange
                    {
                        Dimension = cellRangeDelete.IsColumn ? "COLUMNS" : "ROWS",
                        SheetId = GetSheetId(cellRangeDelete.Sheetname),
                        StartIndex = cellRangeDelete.StartIndex,
                        EndIndex = cellRangeDelete.EndIndex
                    }
                }
            };
        }


    }

}
