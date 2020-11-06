using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Sheets.Apis.Utility.Sheets.v4;
using Microsoft.AspNetCore.Mvc;
using Server.Dtos;
using System.Collections.Generic;
using System.Linq;

namespace Server.Controllers
{
    /// <summary>
    /// https://developers.google.com/sheets/api/samples/rowcolumn
    /// </summary>
    public class RangeController : ApiControllerBase
    {
        #region --- Read Range ---
        [HttpGet("column")]
        public IActionResult GetColumnLine(string sheetname, string range)
        {
            return CheckSheetExist(sheetname, () =>
            {
                var result = SheetsService.ReadRange(GoogleSheetsApi.SpreadSheetId,
                     SheetRange.ToRange(sheetname, range),
                     r =>
                     {
                         r.MajorDimension = SpreadsheetsResource.ValuesResource.GetRequest.MajorDimensionEnum.COLUMNS;
                     });
                return Ok(result);
            });
        }

        [HttpGet("row")]
        public IActionResult GetRowLine(string sheetname, string range)
        {
            return CheckSheetExist(sheetname, () =>
            {
                var result = SheetsService.ReadRange(GoogleSheetsApi.SpreadSheetId,
                     SheetRange.ToRange(sheetname, range),
                     r =>
                     {
                         r.MajorDimension = SpreadsheetsResource.ValuesResource.GetRequest.MajorDimensionEnum.ROWS;
                     });
                return Ok(result);
            });
        }
        #endregion

        #region --- Append Range ---
        [HttpPost("column")]
        public IActionResult AppendColumnLine(AppendColumnDto column)
        {
            return CheckSheetExist(column.Sheetname, () =>
            {
                var result = SheetsService.AppendColumnLine(
                    GoogleSheetsApi.SpreadSheetId,
                    SheetRange.ToRange(column.Sheetname, column.Range),
                    request => { },
                    column.Line);
                return Ok(result);
            });
        }

        [HttpPost("row")]
        public IActionResult AppendRowLine(AppendRowDto row)
        {
            return CheckSheetExist(row.Sheetname, () =>
            {
                var result = SheetsService.AppendRowLine(
                    GoogleSheetsApi.SpreadSheetId,
                    SheetRange.ToRange(row.Sheetname, row.Range),
                    request =>
                    {
                        request.InsertDataOption = Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
                    },
                    row.Line);
                return Ok(result);
            });
        }

        [HttpPost]
        public IActionResult AppendRange(AppendRangeDto range)
        {
            return CheckSheetExist(range.Sheetname, () =>
            {
                var result = SheetsService.AppendRange(
                    GoogleSheetsApi.SpreadSheetId,
                    SheetRange.ToRange(range.Sheetname, range.Range),
                    request =>
                    {
                        request.InsertDataOption = Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
                    },
                    range.Values);
                return Ok(result);
            });
        }
        #endregion

        #region --- Update Range ---
        [HttpPut("column")]
        public IActionResult UpdateColumnLine(AppendColumnDto column)
        {
            return CheckSheetExist(column.Sheetname, () =>
            {
                var result = SheetsService.UpdateColumnLine(
                    GoogleSheetsApi.SpreadSheetId,
                    SheetRange.ToRange(column.Sheetname, column.Range),
                    request => { },
                    column.Line);
                return Ok(result);
            });
        }

        [HttpPut("row")]
        public IActionResult UpdateRowLine(AppendRowDto row)
        {
            return CheckSheetExist(row.Sheetname, () =>
            {
                var result = SheetsService.UpdateRowLine(
                    GoogleSheetsApi.SpreadSheetId,
                    new SheetRange(row.Sheetname, row.Range).ToRange(),
                    request => { },
                    row.Line);
                return Ok(result);
            });
        }

        [HttpPut]
        public IActionResult UpdateRange(AppendRangeDto range)
        {
            return CheckSheetExist(range.Sheetname, () =>
            {
                var result = SheetsService.UpdateRange(
                    GoogleSheetsApi.SpreadSheetId,
                    new SheetRange(range.Sheetname, range.Range).ToRange(),
                    request => { },
                    range.Values);
                return Ok(result);
            });
        }
        #endregion

        #region --- Delete Range --
        [HttpDelete("column")]
        public IActionResult DeleteColumnLine(DeleteRangeDto range)
        {
            range.IsColumn = true;
            return Ok(BatchUpdateSpreadsheet(new DeleteRangeDto[] { range })[0]);
        }

        [HttpDelete("row")]
        public IActionResult DeleteRowLine(DeleteRangeDto range)
        {
            range.IsColumn = false;
            return Ok(BatchUpdateSpreadsheet(new DeleteRangeDto[] { range })[0]);
        }

        [HttpDelete]
        public IActionResult DeleteRange(DeleteRangeDto[] ranges)
        {
            return Ok(BatchUpdateSpreadsheet(ranges));
        }

        private IList<BatchUpdateSpreadsheetResponse> BatchUpdateSpreadsheet(DeleteRangeDto[] ranges)
        {
            if (ranges is null)
            {
                throw new System.ArgumentNullException(nameof(ranges));
            }

            return ranges.Select(c =>
            {
                return CheckSheetExist(c.Sheetname, () =>
                {
                    var request = CreateDeleteDimensionRequest(c);
                    return SheetsService.BatchUpdateSpreadsheet(GoogleSheetsApi.SpreadSheetId, request);
                });
            }).ToList();
        }

        private Request CreateDeleteDimensionRequest(DeleteRangeDto range)
        {
            return new Request
            {
                DeleteDimension = new DeleteDimensionRequest
                {
                    Range = new DimensionRange
                    {
                        Dimension = range.IsColumn ? "COLUMNS" : "ROWS",
                        SheetId = GetSheetId(range.Sheetname),
                        StartIndex = range.StartIndex,
                        EndIndex = range.EndIndex
                    }
                }
            };
        }
        #endregion

        #region --- Clear Range ---
        [HttpPost("clear")]
        public IActionResult ClearRange(ClearRangeDto range)
        {
            return CheckSheetExist(range.Sheetname, () =>
            {
                var result = SheetsService.ClearRange(GoogleSheetsApi.SpreadSheetId, range.Range);
                return Ok(result);
            });
        }
        #endregion

        #region --- Find Replace Range ---
        // TODO Find Replace Range
        #endregion

        #region --- Append Empty ---
        [HttpPost("column/empty")]
        public IActionResult AppendColumnEmpty(AppendRangeEmptyDto range)
        {
            range.IsColumn = true;
            return Created(nameof(AppendColumnEmpty), BatchUpdateSpreadsheet(new AppendRangeEmptyDto[] { range })[0]);
        }

        [HttpPost("row/empty")]
        public IActionResult AppendRowEmpty(AppendRangeEmptyDto range)
        {
            range.IsColumn = false;
            return Created(nameof(AppendRowEmpty), BatchUpdateSpreadsheet(new AppendRangeEmptyDto[] { range })[0]);
        }

        [HttpPost("empty")]
        public IActionResult AppendRangeEmpty(AppendRangeEmptyDto[] range)
        {
            return Created(nameof(AppendRangeEmpty), BatchUpdateSpreadsheet(range));
        }

        private IList<BatchUpdateSpreadsheetResponse> BatchUpdateSpreadsheet(AppendRangeEmptyDto[] ranges)
        {
            if (ranges is null)
            {
                throw new System.ArgumentNullException(nameof(ranges));
            }

            return ranges.Select(r =>
            {
                return CheckSheetExist(r.Sheetname, () =>
                {
                    var request = CreateAppendDimensionRequest(r);
                    return SheetsService.BatchUpdateSpreadsheet(GoogleSheetsApi.SpreadSheetId, request);
                });
            }).ToList();
        }

        private Request CreateAppendDimensionRequest(AppendRangeEmptyDto range)
        {
            return new Request
            {
                AppendDimension = new AppendDimensionRequest
                {
                    Dimension = range.IsColumn ? "COLUMNS" : "ROWS",
                    SheetId = GetSheetId(range.Sheetname),
                    Length = range.Length
                }
            };
        }
        #endregion

        #region --- Insert Empty --
        [HttpPost("column/empty/insert")]
        public IActionResult InsertColumnEmpty(InsertRangeEmptyDto range)
        {
            range.IsColumn = true;
            return Created(nameof(InsertColumnEmpty), BatchUpdateSpreadsheet(new InsertRangeEmptyDto[] { range })[0]);
        }

        [HttpPost("row/empty/insert")]
        public IActionResult InsertRowEmpty(InsertRangeEmptyDto range)
        {
            range.IsColumn = false;
            return Created(nameof(InsertRowEmpty), BatchUpdateSpreadsheet(new InsertRangeEmptyDto[] { range })[0]);
        }

        [HttpPost("empty/insert")]
        public IActionResult InsertRangeEmpty(InsertRangeEmptyDto[] ranges)
        {
            return Created(nameof(InsertRangeEmpty), BatchUpdateSpreadsheet(ranges));
        }

        private IList<BatchUpdateSpreadsheetResponse> BatchUpdateSpreadsheet(InsertRangeEmptyDto[] ranges)
        {
            if (ranges is null)
            {
                throw new System.ArgumentNullException(nameof(ranges));
            }

            return ranges.Select(c =>
            {
                return CheckSheetExist(c.Sheetname, () =>
                {
                    var request = CreateInsertDimensionRequest(c);
                    return SheetsService.BatchUpdateSpreadsheet(GoogleSheetsApi.SpreadSheetId, request);
                });
            }).ToList();
        }

        private Request CreateInsertDimensionRequest(InsertRangeEmptyDto range)
        {
            return new Request
            {
                InsertDimension = new InsertDimensionRequest
                {
                    Range = new DimensionRange
                    {
                        Dimension = range.IsColumn ? "COLUMNS" : "ROWS",
                        SheetId = GetSheetId(range.Sheetname),
                        StartIndex = range.StartIndex,
                        EndIndex = range.EndIndex
                    },
                    InheritFromBefore = range.InheritFromBefore
                }
            };
        }
        #endregion

      
    }
}
