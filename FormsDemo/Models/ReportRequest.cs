using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace FormsDemo.Models
{
    public class ReportRequest
    {
        public Dictionary<string, string> Values { get; set; }
        public Dictionary<string, List<Dictionary<string, string>>> TableValues { get; set; }
        public string TemplateName { get; set; }
        public string AuthorName { get; set; }
        public string Keywords { get; set; }
    }
}
