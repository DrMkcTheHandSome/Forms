using FormsDemo.ErrorHandlers;
using FormsDemo.Interfaces;
using FormsDemo.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace FormsDemo.Services
{
    public class PacketReportService : IPacketReportService
    {
        private readonly IWordToPdfService _wordToPdfService;

        public PacketReportService(IWordToPdfService wordToPdfService)
        {
            _wordToPdfService = wordToPdfService;
        }

        public Dictionary<string, string> GetFormValues(string formName)
        {
            switch (formName)
            {
                case "NJ Amendatory": { return GetNJAmendatoryFormValues(); };
                case "NE Amendatory": { return GetNJAmendatoryFormValues(); };
                default: return new Dictionary<string, string>();
            };
        }

        #region Form Values 
        private Dictionary<string, string> GetNJAmendatoryFormValues()
        {
            return new Dictionary<string, string>() {
                {"NewJersey", "Hello Marc Kenneth Lomio!"},
                {"Marius", "Hello Marius!"}
            };
        }
        #endregion

        public Dictionary<string, List<Dictionary<string, string>>> GetTableValues(string formName)
        {
            switch (formName)
            {
                default: return new Dictionary<string, List<Dictionary<string, string>>>();
            };
        }

        public async Task<byte[]> GeneratePdfFileAsync(Dictionary<string, List<Dictionary<string, string>>> tableValues,
            Dictionary<string, string> formValues, string template)
        {
            try
            {
                var param = new ReportRequest()
                {
                    Values = formValues,
                    TableValues = tableValues,
                    TemplateName = template
                };
                var pdfByte = await _wordToPdfService.GenerateReport(param);
                return pdfByte;
            }
            catch (Exception ex)
            {
                throw new CustomErrorException($"{ex.Message}");
            }
        }
    }
}
