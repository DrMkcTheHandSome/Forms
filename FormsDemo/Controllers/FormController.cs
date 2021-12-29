using FormsDemo.Helpers;
using FormsDemo.Interfaces;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace FormsDemo.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class FormController : ControllerBase
    {
      
        private readonly ILogger<FormController> _logger;
        private readonly IFormService _formService;
        public FormController(ILogger<FormController> logger,
             IFormService formService)
        {
            _logger = logger;
            _formService = formService;
        }

        [HttpGet]
        public async Task<IActionResult> GetFormDocuments()
        {
            try
            {
                var packetFile = await _formService.GenerateViewPacketReport();

                var packetReportFile = new FileResultFromStream(
                    $"policy_packet_report.pdf",
                    new MemoryStream(packetFile),
                    "application/pdf"
                );

                return packetReportFile;
            }
            catch (Exception e)
            {
                return BadRequest(e.Message);
            }
        }
    }
}
