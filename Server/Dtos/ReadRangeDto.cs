namespace Server.Dtos
{
    public class ReadColumnDto
    {
        public string Sheetname { get; set; }

        public string Range { get; set; }
    }

    public class ReadRowDto
    {
        public string Sheetname { get; set; }

        public string Range { get; set; }
    }

    public class ReadRangeDto
    {
        public string Sheetname { get; set; }

        public string Range { get; set; }

        public bool IsColumn { get; set; }
    }
}
