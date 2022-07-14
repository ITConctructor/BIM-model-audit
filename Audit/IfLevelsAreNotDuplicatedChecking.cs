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
    public class IfLevelsAreNotDuplicated : CheckingTemplate
    {
        //pushing test after online merging
        public IfLevelsAreNotDuplicated()
        {
            Name = "ОБЩ_Уровни не дублируются";
            Dep = "ОБЩ";
        }
        public Result Run()
        {
            return Result.Succeeded;
        }
    }
}
