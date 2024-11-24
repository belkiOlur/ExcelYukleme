using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelYukleme.Service
{
    public interface IExcelService
    {
       Task<byte[]> IlceIdIsle(IFormFile uploadedFilee);
       string ExceliDatabaseIsleme(IFormFile uploadedFile)
    }
}
