using ExcelYukleme.Core.ViewModels;
using ExcelYukleme.Service;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
namespace ExcelYukleme.Controllers
{
    public class HomeController : Controller
    {
        private readonly ICalculateService _calculate;
        private readonly IExcelService _excelService;
        public HomeController(ICalculateService calculate, IExcelService excelService)
        {
            _calculate = calculate;
            _excelService = excelService;
        }

        public IActionResult Index()
        {
            return View();
        }
        public async Task<ActionResult> IlceIdIsleme(IFormFile uploadedFilee)
        {
            var fileContent = await _excelService.IlceIdIsle(uploadedFilee);
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
       
        public IActionResult ExcelYukle(IFormFile uploadedFile)
        {
            string bilgi = _excelService.ExceliDatabaseIsleme(uploadedFile);
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
