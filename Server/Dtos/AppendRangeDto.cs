using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Server.Dtos
{
    public class AppendRangeDto
    {
        public string Sheetname { get; set; }

        public string Range { get; set; }

        public IList<IList<object>> Values { get; set; }
    }
}
