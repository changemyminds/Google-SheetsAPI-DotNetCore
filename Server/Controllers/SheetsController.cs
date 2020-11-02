using Google.Apis.Sheets.v4;
using Google.Sheets.Apis.Utility.Sheets.v4;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Server.Application;
using System.Net;

namespace Server.Controllers
{
    public class SheetsController : ApiControllerBase
    {
        [HttpGet("all")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status204NoContent)]
        public IActionResult Names()
            => Ok(SheetsService.GetSheetsName(GoogleSheetsApi.SpreadSheetId));

        [HttpGet("exist")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        public IActionResult Exist(string sheetname)
        {
            return CheckWorkSheetExist(sheetname, () => Ok(sheetname));
        }

        [HttpPost("create")]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        [ProducesResponseType(StatusCodes.Status201Created)]
        public IActionResult Create(string sheetname)
        {
            if (IsWorkSheetExist(sheetname))
            {
                throw new RESTfulException(HttpStatusCode.BadRequest, $"The sheet {sheetname} is exist");
            }

            SheetsService.AddSheet(GoogleSheetsApi.SpreadSheetId, sheetname);
            return CreatedAtAction(nameof(Create), sheetname);
        }

        [HttpDelete("delete")]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        [ProducesResponseType(StatusCodes.Status200OK)]
        public IActionResult Delete(string sheetname)
        {
            return CheckWorkSheetExist(sheetname, () =>
            {
                SheetsService.DeleteSheet(GoogleSheetsApi.SpreadSheetId, sheetname);
                return Ok();
            }, $"Delete failed. Can't find the sheet name with {sheetname}");
        }

        [HttpPut("update")]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        [ProducesResponseType(StatusCodes.Status200OK)]
        public IActionResult Update(string sheetname, string newsheetname)
        {
            return CheckWorkSheetExist(sheetname, () =>
            {
                SheetsService.UpdateSheetTitle(GoogleSheetsApi.SpreadSheetId, sheetname, newsheetname);
                return Ok();
            }, $"Update failed. Can't find the sheet name with {sheetname}");
        }


    }
}
