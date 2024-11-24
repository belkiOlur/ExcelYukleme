using ExcelYukleme.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion.Internal;
using Microsoft.IdentityModel.Abstractions;
using OfficeOpenXml;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using static OfficeOpenXml.ExcelErrorValue;

namespace ExcelYukleme.Controllers
{
    public class HomeController : Controller
    {
        private readonly AppDbContext _context;
        public HomeController(AppDbContext context)
        {
            _context = context;
        }

        public IActionResult Index()
        {
            return View();
        }
        public async Task<ActionResult> IlceIdIsleme(IFormFile uploadedFilee)
        {

            string bilgi = "";
            string ilceId = "";
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
                        satir.Add("Sýra No");
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
                                string ilce = worksheet.Cells[row, 24].Text.ToLower();
                                if (ilce == "" || ilce == null)
                                {
                                    if (adres != null && adres != "")
                                    {
                                        foreach (var item in ilceler)
                                        {
                                            if (adres.Contains(item.IlceAdi.ToLower()))
                                            {
                                                ilceId = Convert.ToString(item.Id);
                                                break;
                                            }
                                        }

                                    }
                                }
                                else
                                {
                                    foreach (var item in ilceler)
                                    {
                                        double similarity = CalculateSimilarity(ilce.ToLower(), item.IlceAdi.ToLower());
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
                                    ilceId = "";
                                }
                            }
                            catch
                            {
                                bilgi += $"{row}. Satýrda Hata oluþtu";
                            }

                            for (int column = 1; column <= columnCount; column++)
                            {
                                if (column == 9)
                                {
                                    satir.Add(worksheet.Cells[row, column].Value?.ToString()!.Trim()!);
                                }
                                else if (column == 24)
                                {
                                    satir.Add(ilceId);
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
            var fileContent = await Indir(list);
            var fileName = "EBSISPersonel.xlsx";
            var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            var cd = new System.Net.Mime.ContentDisposition
            {
                FileName = fileName,
                Inline = false,
            };
            Response.Headers.Append("Content-Disposition", cd.ToString());
            Response.Cookies.Append("DownloadToken", "true", new CookieOptions
            {
                Expires = DateTime.Now.AddMinutes(1),
                HttpOnly = false
            });
            return File(fileContent, contentType);
        }
        private async Task<byte[]> Indir(List<List<string>> list)
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
        public IActionResult ExcelYukle(IFormFile uploadedFile)
        {
            string bilgi = ExceliDatabaseIsleme(uploadedFile);
            if (bilgi.Contains("Güncellendi."))
            {
                TempData["Status"] = bilgi;
            }
            if (bilgi.Contains("Hatalý."))
            {
                TempData["Error"] = bilgi;
            }
            if (bilgi.Contains("Eklenemedi."))
            {
                TempData["KismiHata"] = bilgi;
            }
            return RedirectToAction("Index");
        }
        private string ExceliDatabaseIsleme(IFormFile uploadedFile)
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
                                        bilgi = $"{row}. Satýrdaki {model.Sicil} Sicilli Personle Ait Veri Hatalý.<br/> {row - 1} Satýra Kadar Güncelleme ve Ekleme Baþarýlý.";
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
                                        bilgi = $"{personel.Sicil} Sicilli Personel Sisteme Eklenemedi. <br/> {row - 1} Satýra Kadar Güncelleme ve Ekleme Baþarýlý";
                                        return bilgi;
                                    }
                                }
                            }
                        }

                    }
                }
            }
            bilgi += "Tüm Personel Baþarýyla Güncellendi. <br/>";
            return bilgi;
        }
        private static double CalculateSimilarity(string source, string target)
        {
            int distance = LevenshteinDistance(source, target);
            int maxLength = Math.Max(source.Length, target.Length);

            return 1.0 - (double)distance / maxLength;
        }

        private static int LevenshteinDistance(string source, string target)
        {
            int[,] matrix = new int[source.Length + 1, target.Length + 1];

            for (int i = 0; i <= source.Length; i++) matrix[i, 0] = i;
            for (int j = 0; j <= target.Length; j++) matrix[0, j] = j;

            for (int i = 1; i <= source.Length; i++)
            {
                for (int j = 1; j <= target.Length; j++)
                {
                    int cost = (source[i - 1] == target[j - 1]) ? 0 : 1;

                    matrix[i, j] = Math.Min(
                        Math.Min(matrix[i - 1, j] + 1, matrix[i, j - 1] + 1),
                        matrix[i - 1, j - 1] + cost
                    );
                }
            }

            return matrix[source.Length, target.Length];
        }
        public IActionResult ExcelIndir()
        {
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/excel", "EBSISPersonelGuncelleme.xlsx");
            byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "EBSISPersonelGuncelleme.xlsx");
        }
        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
