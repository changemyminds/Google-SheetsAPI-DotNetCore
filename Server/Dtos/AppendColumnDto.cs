namespace Server.Dtos
{
    public class AppendColumnDto
    {
        public string Sheetname { get; set; }

        public string Range { get; set; }

        public string[] Line { get; set; }
    }
}
