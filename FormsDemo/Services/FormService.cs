using FormsDemo.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace FormsDemo.Services
{
    public class FormService : IFormService
    {
        private readonly IPacketBuilderService _packetBuilderService;
        public FormService(IPacketBuilderService packetBuilderService)
        {
            _packetBuilderService = packetBuilderService;
        }
        public async Task<byte[]> GenerateViewPacketReport()
        {
            var packetPages = await _packetBuilderService.BuildPolicyPacket();

            var consolidatedPagesBytes = _packetBuilderService.GetConsolidatedPagesBytes(packetPages);

            return consolidatedPagesBytes;
        }
    }
}
