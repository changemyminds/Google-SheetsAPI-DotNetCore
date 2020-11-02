using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Server.Dtos
{
    public class CellColumnDto
    {
        public string Sheetname { get; set; }

        public string Cellrange { get; set; }

        public string[] Line { get; set; }
    }
}
