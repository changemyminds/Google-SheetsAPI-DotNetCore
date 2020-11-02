using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Google.Sheets.Apis.Utility.Sheets.v4
{
    /// <summary>
    /// https://developers.google.com/sheets/api/samples
    /// </summary>
    public static class SheetsServiceExtension
    {
        #region --- Sheets ---         
        public static bool ISheetExist(this SheetsService service, string spreadsheetId, string sheetName)
            => service.GetSheet(spreadsheetId, sheetName) != null;

        public static IEnumerable<string> GetSheetsName(this SheetsService service, string spreadsheetId)
            => service.GetSheets(spreadsheetId).Select(s => s.Properties.Title);

        /// <summary>
        /// 新增工作表
        /// </summary>
        public static void AddSheet(this SheetsService service, string spreadsheetId, string sheetName)
        {
            // 設定工作表名稱到工作表屬性中
            var sheetProperties = new SheetProperties
            {
                Title = sheetName
            };

            // 新增工作表的請求
            var addSheetRequest = new AddSheetRequest
            {
                Properties = sheetProperties
            };

            // 創建一個新的請求
            var request = new Request
            {
                AddSheet = addSheetRequest
            };

            // 送出請求
            service.BatchUpdateSpreadsheet(spreadsheetId, request);
        }

        /// <summary>
        /// 刪除工作表
        /// </summary>         
        public static void DeleteSheet(this SheetsService service, string spreadsheetId, string sheetName)
        {
            var sheet = service.GetSheet(spreadsheetId, sheetName);
            if (sheet == null) throw new NullReferenceException($"Can't find the spreadsheetId {spreadsheetId}, sheetName {sheetName}");

            // 創建刪除Sheet指令
            var deleteSheetRequest = new DeleteSheetRequest
            {
                SheetId = sheet.Properties.SheetId
            };

            // 創建一個要求
            var request = new Request
            {
                DeleteSheet = deleteSheetRequest
            };

            // 送出請求
            service.BatchUpdateSpreadsheet(spreadsheetId, request);
        }

        /// <summary>
        /// 更新工作表名稱
        /// </summary>
        public static void UpdateSheetTitle(this SheetsService service, string spreadsheetId, string oldName, string newName)
        {
            var sheet = service.GetSheet(spreadsheetId, oldName);
            if (sheet == null) throw new NullReferenceException($"Can't find the spreadsheetId {spreadsheetId}, sheetName {oldName}");

            var sheetProperties = new SheetProperties
            {
                SheetId = sheet.Properties.SheetId,
                Title = newName
            };

            var updateSheetPropertiesRequest = new UpdateSheetPropertiesRequest
            {
                Properties = sheetProperties,
                Fields = "title"
            };

            //Create a new request with containing the addSheetRequest and add it to the requestList
            var request = new Request
            {
                UpdateSheetProperties = updateSheetPropertiesRequest
            };

            service.BatchUpdateSpreadsheet(spreadsheetId, request);
        }

        /// <summary>
        /// 取得藉由spreadsheetId可以取得Google Spreadsheet，再透過<see cref="GetSheets(SheetsService, string)"/>
        /// 可以取得Google Sheet工作表
        /// </summary>
        public static Spreadsheet GetSpreadsheet(this SheetsService service, string spreadsheetId)
            => service.Spreadsheets.Get(spreadsheetId).Execute();

        /// <summary>
        /// 取得Google Sheet所有的工作表
        /// </summary>
        public static IEnumerable<Sheet> GetSheets(this SheetsService service, string spreadsheetId)
            => service.GetSpreadsheet(spreadsheetId).Sheets;

        /// <summary>
        /// 取得Google Sheet工作表
        /// </summary>
        public static Sheet GetSheet(this SheetsService service, string spreadsheetId, string sheetName)
            => service.GetSheets(spreadsheetId).SingleOrDefault(sheet => sheet.Properties.Title.Equals(sheetName));

        public static int? GetSheetId(this SheetsService service, string spreadsheetId, string sheetName)
            => service.GetSheet(spreadsheetId, sheetName).Properties.SheetId;
        #endregion

        #region --- Cell Range ---
        public static AppendValuesResponse AppendRowLine(this SheetsService service, string spreadsheetId, SheetRange sheetRange, Action<SpreadsheetsResource.ValuesResource.AppendRequest> onRequest, params object[] rowValues)
            => service.AppendRowLine(spreadsheetId, sheetRange.ToRange(), onRequest, rowValues);

        public static AppendValuesResponse AppendRowLine(this SheetsService service, string spreadsheetId, string range, Action<SpreadsheetsResource.ValuesResource.AppendRequest> onRequest, params object[] rowValues)
        {
            var values = new List<IList<object>>
            {
                rowValues.ToList()
            };

            return service.AppendMultiLine(spreadsheetId, range, onRequest, values);
        }

        public static UpdateValuesResponse UpdateRowLine(this SheetsService service, string spreadsheetId, SheetRange sheetRange, Action<SpreadsheetsResource.ValuesResource.UpdateRequest> onRequest, params object[] rowValues)
            => service.UpdateRowLine(spreadsheetId, sheetRange.ToRange(), onRequest, rowValues);

        public static UpdateValuesResponse UpdateRowLine(this SheetsService service, string spreadsheetId, string range, Action<SpreadsheetsResource.ValuesResource.UpdateRequest> onRequest, params object[] rowValues)
        {
            var values = new List<IList<object>>
            {
                rowValues.ToList()
            };

            return service.UpdateMultiLine(spreadsheetId, range, values, onRequest);
        }

        public static AppendValuesResponse AppendColumnLine(this SheetsService service, string spreadsheetId, SheetRange sheetRange, Action<SpreadsheetsResource.ValuesResource.AppendRequest> onRequest, params object[] columnValues)
            => service.AppendColumnLine(spreadsheetId, sheetRange.ToRange(), onRequest, columnValues);

        public static AppendValuesResponse AppendColumnLine(this SheetsService service, string spreadsheetId, string range, Action<SpreadsheetsResource.ValuesResource.AppendRequest> onRequest, params object[] columnValues)
        {
            // convert columnValues to columList
            var columList = columnValues.Select(v => new List<object> { v });

            // Add columList to values and input to valueRange
            var values = new List<IList<object>>();
            values.AddRange(columList.ToList());

            return service.AppendMultiLine(spreadsheetId, range, onRequest, values);
        }

        public static UpdateValuesResponse UpdateColumnLine(this SheetsService service, string spreadsheetId, SheetRange sheetRange, Action<SpreadsheetsResource.ValuesResource.UpdateRequest> onRequest, params object[] columnValues)
           => service.UpdateColumnLine(spreadsheetId, sheetRange.ToRange(), onRequest, columnValues);

        public static UpdateValuesResponse UpdateColumnLine(this SheetsService service, string spreadsheetId, string range, Action<SpreadsheetsResource.ValuesResource.UpdateRequest> onRequest, params object[] columnValues)
        {
            // convert columnValues to columList
            var columList = columnValues.Select(v => new List<object> { v });

            // Add columList to values and input to valueRange
            var values = new List<IList<object>>();
            values.AddRange(columList.ToList());

            return service.UpdateMultiLine(spreadsheetId, range, values, onRequest);
        }

        /// <summary>
        /// 批次更新數值
        /// </summary>
        public static BatchUpdateValuesResponse BatchUpdateValues(this SheetsService service, string spreadsheetId, SheetOption.ValueInputOption option, List<ValueRange> valueRanges)
        {
            var request = new BatchUpdateValuesRequest
            {
                ValueInputOption = option.ToString(),
                Data = valueRanges
            };

            return service.BatchUpdateValues(spreadsheetId, request);
        }

        /// <summary>
        /// 批次更新數值
        /// </summary>
        public static BatchUpdateValuesResponse BatchUpdateValues(this SheetsService service, string spreadsheetId, BatchUpdateValuesRequest request)
            => service.Spreadsheets.Values.BatchUpdate(request, spreadsheetId).Execute();

        public static AppendValuesResponse AppendMultiLine(this SheetsService service,
            string spreadsheetId,
            string range,
            Action<SpreadsheetsResource.ValuesResource.AppendRequest> onRequest,
            IList<IList<object>> values)
        {
            var body = new ValueRange()
            {
                Values = values
            };

            var appendRequest = service.Spreadsheets.Values.Append(body, spreadsheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            onRequest?.Invoke(appendRequest);
            return appendRequest.Execute();
        }

        public static UpdateValuesResponse UpdateMultiLine(this SheetsService service,
           string spreadsheetId,
           string range,
           IList<IList<object>> values,
           Action<SpreadsheetsResource.ValuesResource.UpdateRequest> onRequest)
        {
            var body = new ValueRange()
            {
                Values = values
            };

            var updateRequest = service.Spreadsheets.Values.Update(body, spreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            onRequest?.Invoke(updateRequest);
            return updateRequest.Execute();
        }
        #endregion

        /// <summary>
        /// 執行批次工作表的請求
        /// </summary>
        public static BatchUpdateSpreadsheetResponse BatchUpdateSpreadsheet(this SheetsService service, string spreadsheetId, params Request[] requests)
        {
            // 創建批次請求
            var batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest
            {
                // 將請求加入到批次請求內
                Requests = requests.ToList()
            };

            // 呼叫工作表的API去執行批次請求
            return service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, spreadsheetId).Execute();
        }
    }
}
