﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Server.Dtos
{
    public class AppendRangeEmptyDto
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
        /// 從哪一個位置結束插入
        /// </summary>
        public int Length { get; set; }
    }
}
