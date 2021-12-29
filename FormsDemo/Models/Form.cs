using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace FormsDemo.Models
{
    public class Form
    {
        public string FormName { get; set; }
        public int FormOrder { get; set; }
        public string TemplateName { get; set; }
        public bool IsSelected { get; set; }
        public bool IsMandatory { get; set; }
    }
}
