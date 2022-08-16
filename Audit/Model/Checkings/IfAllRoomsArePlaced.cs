using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace Audit.Checkings
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
        public override Result Run(string filePath)
        {
            TaskDialog dialog = new TaskDialog("Test");
            dialog.MainContent = Name;
            dialog.Show();
            return Result.Succeeded;
        }
    }
}

