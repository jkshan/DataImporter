namespace DataImporter.Models
{
    public class Record
    {
        public string Account { get; set; }
        public string Description { get; set; }
        public string CurrencyCode { get; set; }
        public string Amount { get; set; }
        public string ErrorMessages { get; set; }
    }
}