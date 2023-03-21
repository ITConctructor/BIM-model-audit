using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Audit
{
    public class FileCheckingResult
    {
        public string FileName;
        public BindingList<CheckingTemplate> Checkings { get; set; } = new BindingList<CheckingTemplate>();
    }
}
