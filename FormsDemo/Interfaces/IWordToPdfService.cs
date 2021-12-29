using FormsDemo.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace FormsDemo.Interfaces
{
   public interface IWordToPdfService
    {
        Task<byte[]> GenerateReport(ReportRequest param);
    }
}
