namespace Server.Dtos
{
    public class InsertRangeEmptyDto
    {
        /// <summary>
        /// 工作表名稱
        /// </summary>
        public string Sheetname { get; set; }

        /// <summary>
        /// 是否為行或列
        /// </summary>
        public bool IsColumn { get; set; }

        /// <summary>
        /// 從哪一個位置開始插入
        /// </summary>
        public int StartIndex { get; set; }

        /// <summary>
        /// 從哪一個位置結束插入
        /// </summary>
        public int EndIndex { get; set; }

        /// <summary>
        /// 繼承插入格的欄位屬性，如果為第一列，獲第一行(A)時，此數值必須為false，因為沒有繼承更早之前的欄位
        /// </summary>
        public bool InheritFromBefore { get; set; }
    }
}
