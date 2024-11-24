using ExcelYukleme.Repository.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;

namespace ExcelYukleme.Service
{
    public class ExcelService:IExcelService
    {
        private readonly AppDbContext _context;
        private readonly ICalculateService _calculate;
        public ExcelService(ICalculateService calculate, AppDbContext context)
        {
            _calculate = calculate;
            _context = context;
        }
        public async Task<byte[]> IdIsle(IFormFile uploadedFilee)
        {
            string bilgi = "";
            string ilceId = "";
            string rutbeId = "";
            string birimId = "";
            string cinsiyetId = "";
            string kanGrubuId = "";
            string medeniDurumId = "";
            string istihkakId = "";
            int i = 0;
            List<string> satir = new();
            List<List<string>> list = new();
            if (uploadedFilee != null && uploadedFilee.Length > 0)
            {
                using (var stream = new MemoryStream())
                {

                    await uploadedFilee.CopyToAsync(stream);
                    using (var package = new ExcelPackage(stream))
                    {
                        var worksheet = package.Workbook.Worksheets[0];
                        var columnCount = 0;
                        var rowCount = 0;
                        for (int row = 1; row <= worksheet.Dimension.Rows; row++)
                        {
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[row, 2].Text))
                            {
                                rowCount = row;
                            }
                        }
                        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        {
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[1, col].Text))
                            {
                                columnCount = col;
                            }
                        }
                        var ilceler = _context.KodIlceler;
                        var rutbeler = _context.KodRutbeler;
                        var birimler = _context.KodBirimler;
                        var kanGruplari = _context.KodKanGruplari;
                        satir.Add("Sıra No");
                        satir.Add("Sicil");
                        satir.Add("Ad");
                        satir.Add("Soyad");
                        satir.Add("Sifre");
                        satir.Add("RutbeId");
                        satir.Add("BirimId");
                        satir.Add("CinsiyetId");
                        satir.Add("TcNo");
                        satir.Add("IbanNo");
                        satir.Add("KanGrubuId");
                        satir.Add("HataliGirisSayisi");
                        satir.Add("FotoIsim");
                        satir.Add("TelsizKodu");
                        satir.Add("MedeniDurumId");
                        satir.Add("Mail");
                        satir.Add("CepTelefonu");
                        satir.Add("SilahMarka");
                        satir.Add("SilahSeriNo");
                        satir.Add("EsSicil");
                        satir.Add("KayitTarihi");
                        satir.Add("IptalMi");
                        satir.Add("Adres");
                        satir.Add("IlceId");
                        satir.Add("IstihkakDurumu");
                        satir.Add("DogumTarihi");
                        list.Add(satir);
                        for (int row = 2; row <= rowCount; row++)
                        {
                            satir = new List<string>();
                            i++;
                            try
                            {
                                string adres = worksheet.Cells[row, 23].Text.ToLower();
                                List<string> adresKelimeleri = adres.Split(' ', ',', '.', '/', '-').ToList();
                                string ilce = worksheet.Cells[row, 24].Text.ToLower();
                                string rutbe = worksheet.Cells[row,6].Text.ToLower();
                                string birim = worksheet.Cells[row, 7].Text.ToLower();
                                string cinsiyet = worksheet.Cells[row, 8].Text.ToLower();
                                string kan = worksheet.Cells[row, 11].Text.ToLower();
                                string evlilikDurumu = worksheet.Cells[row, 15].Text.ToLower();
                                var istihkakDurumu = worksheet.Cells[row, 25].Text.ToLower(); 
                                if (ilce == "" || ilce == null)
                                {
                                    if (adres != null && adres != "")
                                    {
                                        foreach (var item in ilceler)
                                        {
                                            foreach (var kelime in adresKelimeleri)
                                            {
                                                double similarity = _calculate.CalculateSimilarity(kelime.ToLower(), item.IlceAdi.ToLower());
                                                if (similarity >= 0.6)
                                                {
                                                    ilceId = Convert.ToString(item.Id);
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    foreach (var item in ilceler)
                                    {
                                        double similarity = _calculate.CalculateSimilarity(ilce.ToLower(), item.IlceAdi.ToLower());
                                        if (similarity >= 0.6)
                                        {
                                            ilceId = Convert.ToString(item.Id);
                                            break;
                                        }
                                        else
                                        {
                                            ilceId = ilce;
                                        }
                                    }
                                }
                                if ((adres == null || adres == "") && (ilce == null || ilce == ""))
                                {
                                    ilceId = "0";
                                }
                                if (rutbe != null && rutbe != "")
                                {
                                    foreach (var item in rutbeler)
                                    {
                                        double similarity = _calculate.CalculateSimilarity(rutbe.ToLower(), item.Ad.ToLower());
                                        if (similarity >= 0.83)
                                        {
                                            rutbeId = Convert.ToString(item.Id);
                                            break;
                                        }
                                        if (rutbe.Contains("1.sınıf"))
                                        {
                                            rutbeId = "8";
                                            break;
                                        }
                                        if (rutbe.Contains("2.sınıf"))
                                        {
                                            rutbeId = "11";
                                            break;
                                        }
                                        if (rutbe.Contains("3.sınıf"))
                                        {
                                            rutbeId = "3";
                                            break;
                                        }
                                        if (rutbe.Contains("4.sınıf"))
                                        {
                                            rutbeId = "12";
                                            break;
                                        }                                       
                                        else
                                        {
                                            rutbeId = rutbe;
                                        }
                                    }
                                }
                                else
                                {
                                    rutbeId = "0";
                                }
                                if (birim != null && birim != "")
                                {
                                    foreach (var item in birimler)
                                    {
                                        double similarity = _calculate.CalculateSimilarity(birim.ToLower(), item.Ad.ToLower());
                                        if (similarity >= 0.95)
                                        {
                                            birimId = Convert.ToString(item.Id);
                                            break;
                                        }                                       
                                        else
                                        {
                                            birimId = birim;
                                        }
                                    }
                                }
                                else
                                {
                                    birimId = "0";
                                }
                                if (cinsiyet != null && cinsiyet != "")
                                {
                                        double similarity = _calculate.CalculateSimilarity(cinsiyet.ToLower(), "erkek");
                                        if (similarity >= 0.6)
                                        {
                                            cinsiyetId = "1";
                                        }
                                        else
                                        {
                                        cinsiyetId = "2";
                                        }                                   
                                }
                                else
                                {
                                    cinsiyetId = "0";
                                }
                                if (evlilikDurumu != null && evlilikDurumu != "")
                                {
                                    double similarity = _calculate.CalculateSimilarity(evlilikDurumu.ToLower(), "evli");
                                    if (similarity >= 0.6)
                                    {
                                        medeniDurumId = "1";
                                    }
                                    else
                                    {
                                        medeniDurumId = "2";
                                    }

                                }
                                else
                                {
                                   medeniDurumId = "0";
                                }
                                if (kan != null && kan != "")
                                {
                                    kan = kan.Replace(" ", "").ToLower();
                                    foreach (var item in kanGruplari)
                                    {
                                        var kanGrb = item.Ad.Replace(" ", "").ToLower();
                                        double similarity = _calculate.CalculateSimilarity(kan, kanGrb);
                                        if (similarity >= 0.85)
                                        {
                                            kanGrubuId = Convert.ToString(item.Id); ;
                                            break;
                                        }
                                        else
                                        {
                                            kanGrubuId = kan;
                                        }
                                    }
                                }
                                else
                                {
                                    kanGrubuId = "0";
                                }
                                if (istihkakDurumu != null && istihkakDurumu != "")
                                {
                                    double similarity = _calculate.CalculateSimilarity(istihkakDurumu.ToLower(), "sivil");
                                    if (similarity >= 0.6)
                                    {
                                        istihkakId = "1";
                                    }
                                    else
                                    {
                                        istihkakId = "2";
                                    }
                                }
                                else
                                {
                                    istihkakId = "0";
                                }
                            }
                            catch
                            {
                                bilgi += $"{row}. Satırda Hata oluştu";
                            }


                            for (int column = 1; column <= columnCount; column++)
                            {
                                if (column == 6)
                                {
                                    satir.Add(rutbeId);
                                }
                                else if (column == 7)
                                {
                                    satir.Add(birimId);
                                }
                               else if (column == 8)
                                {
                                    satir.Add(cinsiyetId);
                                }
                                else if (column == 9)
                                {
                                    satir.Add(worksheet.Cells[row, column].Value?.ToString()!.Trim()!);
                                }
                                else if (column == 11)
                                {
                                    satir.Add(kanGrubuId);
                                }
                                else if (column == 15)
                                {
                                    satir.Add(medeniDurumId);
                                }
                                else if (column == 24)
                                {
                                    satir.Add(ilceId);
                                }
                                else if (column == 25)
                                {
                                    satir.Add(istihkakId);
                                }
                                else if (column == 26)
                                {
                                    worksheet.Cells[row, 26].Style.Numberformat.Format = "yyyy-MM-dd";
                                    satir.Add(worksheet.Cells[row, column].Text);
                                }
                                else
                                {
                                    satir.Add(worksheet.Cells[row, column].Text);
                                }

                            }
                            list.Add(satir);
                        }

                    }
                }
            }
            return await ExcelHazirla(list);
        }

        private async Task<byte[]> ExcelHazirla(List<List<string>> list)
        {
            int i = 0;
            int column = list[0].Count;
            using (var workbook = new ExcelPackage())
            {

                var worksheet = workbook.Workbook.Worksheets.Add("Sayfa");
                foreach (var item in list)
                {
                    i++;
                    for (int j = 0; j < column; j++)
                    {

                        if (j == 0)
                        {
                            worksheet.Cells[i, j + 1].Value = i - 1;
                        }
                        else
                        {
                            worksheet.Cells[i, j + 1].Value = item[j];
                        }
                        if (i == 1 && j == 0)
                        {
                            worksheet.Cells[1, 1].Value = item[j];
                        }
                    }
                }
                workbook.Save();
                return workbook.GetAsByteArray();
            }

        }
        public string ExceliDatabaseIsleme(IFormFile uploadedFile)
        {
            string bilgi = "";
            if (uploadedFile != null && uploadedFile.Length > 0)
            {
                using (var stream = new MemoryStream())
                {
                    uploadedFile.CopyToAsync(stream);
                    using (var package = new ExcelPackage(stream))
                    {
                        var worksheet = package.Workbook.Worksheets[0];
                        var columnCount = 0;
                        var rowCount = 0;
                        for (int row = 1; row <= worksheet.Dimension.Rows; row++)
                        {
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[row, 2].Text))
                            {
                                rowCount = row;
                            }
                        }
                        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        {
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[1, col].Text))
                            {
                                columnCount = col;
                            }
                        }
                        PersonelModel model = new();
                        for (int row = 2; row <= rowCount; row++)
                        {
                            string sicil = worksheet.Cells[row, 2].Text;
                            if (sicil != null && sicil != "")
                            {
                                model = _context.Personeller.AsNoTracking().Where(x => x.Sicil == sicil).FirstOrDefault()!;
                                if (model != null)
                                {
                                    try
                                    {
                                        model.Adres = worksheet.Cells[row, 23].Text;
                                        string excelIlce = worksheet.Cells[row, 24].Text;
                                        if (excelIlce == null || excelIlce == "")
                                        {
                                            model.IlceId = 0;
                                        }
                                        else
                                        {
                                            model.IlceId = Convert.ToInt32(worksheet.Cells[row, 24].Text);
                                        }
                                        model.IstihkakDurumu = Convert.ToInt32(worksheet.Cells[row, 25].Text);
                                        string dogumTarihi = worksheet.Cells[row, 26].Text;
                                        string[] dateParts = dogumTarihi.Split('-', '.', '/');
                                        if (dateParts.Length >= 3)
                                        {
                                            model.DogumTarihi = Convert.ToDateTime(worksheet.Cells[row, 26].Text);
                                        }
                                        else
                                        {
                                            model.DogumTarihi = new DateTime(1970, 1, 1);
                                        }
                                    }
                                    catch
                                    {
                                        bilgi = $"{row}. Satırdaki {model.Sicil} Sicilli Personle Ait Veri Hatalı.<br/> {row - 1} Satıra Kadar Güncelleme ve Ekleme Başarılı.";
                                        return bilgi;
                                    }
                                    _context.Personeller.Update(model);
                                    _context.SaveChanges();
                                }
                                else
                                {
                                    PersonelModel personel = new();
                                    try
                                    {

                                        personel.Id = Guid.NewGuid();
                                        personel.Sicil = worksheet.Cells[row, 2].Text;
                                        personel.Ad = worksheet.Cells[row, 3].Text;
                                        personel.Soyad = worksheet.Cells[row, 4].Text;
                                        personel.Sifre = "B61FEF74D1E1C848DD109B93DAE4C9CEB7CB5E362F24CF4D1AAA50DD30D1305F";
                                        personel.RutbeId = Convert.ToInt32(worksheet.Cells[row, 6].Text);
                                        personel.BirimId = Convert.ToInt32(worksheet.Cells[row, 7].Text);
                                        personel.CinsiyetId = Convert.ToInt32(worksheet.Cells[row, 8].Text);
                                        personel.TcNo = worksheet.Cells[row, 9].Value?.ToString()!.Trim()!;
                                        personel.IbanNo = worksheet.Cells[row, 10].Text;
                                        personel.KanGrubuId = Convert.ToInt32(worksheet.Cells[row, 11].Text);
                                        personel.HataliGirisSayisi = Convert.ToInt32(worksheet.Cells[row, 12].Text);
                                        personel.FotoIsim = worksheet.Cells[row, 13].Text;
                                        personel.TelsizKodu = worksheet.Cells[row, 14].Text;
                                        personel.MedeniDurumId = Convert.ToInt32(worksheet.Cells[row, 15].Text);
                                        personel.Mail = worksheet.Cells[row, 16].Text;
                                        personel.CepTelefonu = worksheet.Cells[row, 17].Text;
                                        personel.SilahMarka = worksheet.Cells[row, 18].Text;
                                        personel.SilahSeriNo = worksheet.Cells[row, 19].Text;
                                        personel.EsSicil = worksheet.Cells[row, 20].Text;
                                        personel.KayitTarihi = DateTime.Now;
                                        personel.IptalMi = false;
                                        personel.Adres = worksheet.Cells[row, 23].Text;
                                        string excelIlce = worksheet.Cells[row, 24].Text;
                                        if (excelIlce == null || excelIlce == "")
                                        {
                                            personel.IlceId = 0;
                                        }
                                        else
                                        {
                                            personel.IlceId = Convert.ToInt32(worksheet.Cells[row, 24].Text);
                                        }
                                        personel.IstihkakDurumu = Convert.ToInt32(worksheet.Cells[row, 25].Text);
                                        string dogumTarihi = worksheet.Cells[row, 26].Text;
                                        string[] dateParts = dogumTarihi.Split('-', '.', '/');
                                        if (dateParts.Length >= 3)
                                        {
                                            personel.DogumTarihi = Convert.ToDateTime(worksheet.Cells[row, 26].Text);
                                        }
                                        else
                                        {
                                            personel.DogumTarihi = new DateTime(1970, 1, 1);
                                        }
                                        _context.Personeller.Add(personel);
                                        _context.SaveChanges();

                                        PersonelRolleriModel rol = new();
                                        List<int> roller = new List<int> { 35, 6 };
                                        foreach (int i in roller)
                                        {
                                            rol.Id = Guid.NewGuid();
                                            rol.PersonelId = personel.Id;
                                            rol.IptalMi = false;
                                            rol.RolId = i;
                                            _context.PersonelRolleri.Add(rol);
                                            _context.SaveChanges();
                                        }
                                        bilgi += $"{personel.Sicil} Sicilli Personel Sisteme Eklendi. <br/>";
                                    }
                                    catch
                                    {
                                        bilgi = $"{personel.Sicil} Sicilli Personel Sisteme Eklenemedi. <br/> {row - 1} Satıra Kadar Güncelleme ve Ekleme Başarılı";
                                        return bilgi;
                                    }
                                }
                            }
                        }

                    }
                }
            }
            bilgi += "Tüm Personel Başarıyla Güncellendi. <br/>";
            return bilgi;
        }


    }
}
