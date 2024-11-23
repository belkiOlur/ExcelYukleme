namespace ExcelYukleme.Models
{
    public class PersonelModel
    {
        public Guid Id { get; set; }
        public string Sicil { get;set; }
        public string Ad { get; set; }
        public string Soyad { get; set; }
        public string Sifre { get; set; }
        public int RutbeId { get; set; }
        public int BirimId { get; set; }
        public int CinsiyetId { get; set; }
        public string TcNo { get; set; }
        public string IbanNo { get; set; }
        public int KanGrubuId { get; set; }
        public int HataliGirisSayisi { get; set; }
        public string? FotoIsim { get; set; }
        public string? TelsizKodu { get; set; }
        public int MedeniDurumId { get; set; }
        public string? Mail { get; set; }
        public string CepTelefonu { get; set; }
        public string? SilahMarka { get; set; }
        public string? SilahSeriNo { get; set; }
        public string? EsSicil { get; set; }
        public DateTime KayitTarihi { get; set; }
        public bool IptalMi { get; set; }
        public string? Adres { get; set; }
        public int IlceId { get; set; }
        public int IstihkakDurumu {  get; set; }
        public DateTime DogumTarihi { get; set; }

    }
}
