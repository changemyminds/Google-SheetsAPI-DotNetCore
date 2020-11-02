using System.Linq;

namespace Google.Sheets.Apis.Utility.Sheets.v4
{
    public class SheetRange
    {
        public string SheetName { get; set; }

        public string Range { get; set; }

        public SheetRange()
        {
        }

        public SheetRange(string sheetName, string range)
        {
            SheetName = sheetName;
            Range = range;
        }

        /// <summary>
        /// 取得範圍
        /// </summary>
        public string ToRange()
        {
            var separated = string.IsNullOrEmpty(Range) ? "" : "!";
            var range = Range + (Range.Any(char.IsDigit) ? "" : "1");
            return SheetName + separated + range;
        }

        public static string ToRange(string sheetName, string range)
        {
            return new SheetRange(sheetName, range).ToRange();
        }
    }
}
