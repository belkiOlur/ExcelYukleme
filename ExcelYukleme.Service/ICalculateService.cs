using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelYukleme.Service
{
    public interface ICalculateService
    {
       double CalculateSimilarity(string source, string target);
    }
}
