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
    [Transaction(TransactionMode.Manual)]
    public class IfLinkInCorrectWorksetChecking : CheckingTemplate
    {
        public IfLinkInCorrectWorksetChecking()
        {
            Name = "ОБЩ_Корректность рабочих наборов связей";
            Dep = "ОБЩ";
        }
        public Result Run()
        {
            Name = "ОБЩ_Корректность рабочих наборов связей";
            return Result.Succeeded;
        }
    }
}
