namespace ExcelYukleme.Models
{
    public class KodIlcelerModel
    {
        public int Id { get; set; }
        public string IlceAdi { get; set; }
        public int UstIlceId { get;set; }
        public int MulkiBirimId { get; set; }
        public bool IptalMi { get; set; }
        public string? KisaAd { get; set; }  


    }
}
