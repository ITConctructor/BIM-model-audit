using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Audit
{
    public class ElementCheckingResult
    {
        public string Name { get; set; }
        public string Status { get; set; }
        public string ID { get; set; }
        public string Comment { get; set; }
        public string Time { get; set; }
        public sealed override bool Equals(object obj)
        {
            ElementCheckingResult result = obj as ElementCheckingResult;
            if (result?.Name == this?.Name && result?.ID == this?.ID) return true;
            return false;
        }
    }
}
