using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace Audit
{
    public class CheckingTemplate
    {
        public string Name { get; set; }
        public string Status { get; set; }
        public string Amount { get; set; }
        public string Dep { get; set; }
        public virtual Result Run()
        {
            return Result.Succeeded;
        }
    }
}
