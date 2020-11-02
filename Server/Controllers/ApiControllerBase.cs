using Google.Apis.Sheets.v4;
using Google.Sheets.Apis.Utility.Sheets.v4;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Server.Application;
using Server.Settings;
using System;
using System.Net;

namespace Server.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ApiControllerBase : ControllerBase
    {
        private SheetsService _sheetsService;
        private GoogleSheetsApi _googleSheetsApi;

        protected SheetsService SheetsService => _sheetsService ??= GetService<SheetsService>();

        protected GoogleSheetsApi GoogleSheetsApi => _googleSheetsApi ??= GetService<GoogleSheetsApi>();

        private T GetService<T>() where T : class => (T)HttpContext.RequestServices.GetService(typeof(T));

        protected bool IsSheetExist(string sheetname)
            => SheetsService.ISheetExist(GoogleSheetsApi.SpreadSheetId, sheetname);

        protected int? GetSheetId(string sheetname) 
            => SheetsService.GetSheetId(GoogleSheetsApi.SpreadSheetId, sheetname);

        protected T CheckSheetExist<T>(string sheetname, Func<T> func, string description = "")
        {
            if (!IsSheetExist(sheetname))
            {
                throw new RESTfulException(HttpStatusCode.NotFound, string.IsNullOrEmpty(description) ? $"Can't find the {sheetname}" : description);
            }

            return func();
        }
    }
}
