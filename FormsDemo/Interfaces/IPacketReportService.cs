using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace FormsDemo.Interfaces
{
   public interface IPacketReportService
    {
        Dictionary<string, string> GetFormValues(string formName);
        Dictionary<string, List<Dictionary<string, string>>> GetTableValues(string formName);
        Task<byte[]> GeneratePdfFileAsync(Dictionary<string, List<Dictionary<string, string>>> tableValues,
            Dictionary<string, string> formValues, string template);
    }
}
