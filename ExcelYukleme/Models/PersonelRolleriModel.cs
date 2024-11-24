namespace ExcelYukleme.Models
{
    public class PersonelRolleriModel
    {
        public Guid Id { get; set; }
        public Guid PersonelId { get; set; }
        public int RolId { get; set; }
        public Guid TanimlayanPersonelId { get; set; }
        public bool IptalMi {  get; set; }
    }
}
