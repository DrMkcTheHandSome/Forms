using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace FormsDemo.Interfaces
{
    public interface IFormService
    {
        Task<byte[]> GenerateViewPacketReport();
    }
}
