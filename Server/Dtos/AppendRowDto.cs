using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Server.Dtos
{
    public class AppendRowDto
    {
        public string Sheetname { get; set; }

        public string Range { get; set; }

        public string[] Line { get; set; }
    }
}
