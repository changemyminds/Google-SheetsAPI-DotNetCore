using System.Linq;

namespace Google.Sheets.Apis.Utility.Sheets.v4
{
    public class SheetRange
    {
        public string WorkSheetName { get; set; }

        public string CellRange { get; set; }

        public SheetRange()
        {
        }

        public SheetRange(string workSheetName, string cellRange)
        {
            WorkSheetName = workSheetName;
            CellRange = cellRange;
        }

        /// <summary>
        /// 取得範圍
        /// </summary>
        public string ToRange()
        {
            var separated = string.IsNullOrEmpty(CellRange) ? "" : "!";
            var cellRange = CellRange + (CellRange.Any(char.IsDigit) ? "" : "1");
            return WorkSheetName + separated + cellRange;
        }
    }
}
