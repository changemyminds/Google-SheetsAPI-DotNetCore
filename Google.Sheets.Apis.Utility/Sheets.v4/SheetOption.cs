namespace Google.Sheets.Apis.Utility.Sheets.v4
{
    public partial class SheetOption
    {
        /**
         * Determines how existing data is changed when new data is message.
         */
        public enum InsertDataOption
        {
            /**
             * The new data overwrites existing data in the areas it is written.
             * (Note: adding data to the end of the sheet will still insert new rows or columns so the data can be written.)
             */
            OVERWRITE,

            /**
             * Rows are inserted for the new data.
             */
            INSERT_ROWS,
        }

        /**
         * Determines how message data should be interpreted.
         */
        public enum ValueInputOption
        {
            /**
             * Default message value. This value must not be used.
             */
            INPUT_VALUE_OPTION_UNSPECIFIED,

            /**
             * The values the user has entered will not be parsed and will be stored as-is.
             * 輸入的數值不會解析，會按照原來的數值輸入
             */
            RAW,

            /**
             * The values will be parsed as if the user typed them into the UI.
             * Numbers will stay as numbers, but strings may be converted to numbers, dates,
             * etc. following the same rules that are applied when entering text into a cell via the Google Sheets UI.
             * 輸入的數值會解析，會根據Google表單內的格式，進行輸入
             */
            USER_ENTERED
        }

        /**
         * Determines how values should be rendered in the output.
         */
        public enum ValueRenderOption
        {
            /**
             * Values will be calculated & formatted in the reply according to the cell's formatting.
             * Formatting is based on the spreadsheet's locale, not the requesting user's locale.
             * For example, if A1 is 1.23 and A2 is =A1 and formatted as currency, then A2 would return "$1.23".
             */
            FORMATTED_VALUE,

            /**
             * Values will be calculated, but not formatted in the reply.
             * For example, if A1 is 1.23 and A2 is =A1 and formatted as currency, then A2 would return the number 1.23.
             */
            UNFORMATTED_VALUE,

            /**
             * Values will not be calculated. The reply will include the formulas.
             * For example, if A1 is 1.23 and A2 is =A1 and formatted as currency, then A2 would return "=A1".
             */
            FORMULA,
        }

        /**
         * Determines how dates should be rendered in the output.
         */
        public enum DateTimeRenderOption
        {
            /**
             * Instructs date, time, datetime, and duration fields to be output as doubles in "serial number" format,
             * as popularized by Lotus 1-2-3.
             * The whole number portion of the value (left of the decimal) counts the days since December 30th 1899.
             * The fractional portion (right of the decimal) counts the time as a fraction of the day.
             * For example, January 1st 1900 at noon would be 2.5, 2 because it's 2 days after December 30st 1899,
             * and .5 because noon is half a day. February 1st 1900 at 3pm would be 33.625.
             * This correctly treats the year 1900 as not a leap year.
             */
            SERIAL_NUMBER,

            /**
             * Instructs date, time, datetime, and duration fields to be output as strings in their given number format
             * (which is dependent on the spreadsheet locale).
             */
            FORMATTED_STRING,
        }
    }
}
