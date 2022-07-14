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
    public class IfLevelNamesAreCorrectChecking : CheckingTemplate
    {
        //pushing test after online merging
        public IfLevelNamesAreCorrectChecking()
        {
            Name = "ОБЩ_Корректность наименований уровней";
            Dep = "ОБЩ";
        }
        public Result Run()
        {
            return Result.Succeeded;
        }
    }
}
