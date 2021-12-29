using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace FormsDemo.Interfaces
{
   public interface IPacketBuilderService
    {
        Task<List<byte[]>> BuildPolicyPacket();
        byte[] GetConsolidatedPagesBytes(List<byte[]> packetLists);
    }
}
