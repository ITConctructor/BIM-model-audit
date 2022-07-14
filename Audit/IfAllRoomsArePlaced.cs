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
    public class IfAllRoomsArePlaced : CheckingTemplate
    {
        //pushing test after online merging
        public IfAllRoomsArePlaced()
        {
            Name = "АР_Нет неразмещенных помещений";
            Dep = "АР";
        }
        public Result Run()
        {
            return Result.Succeeded;
        }
    }
}

