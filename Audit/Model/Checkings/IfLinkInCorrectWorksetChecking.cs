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
    public class IfLinkInCorrectWorksetChecking : CheckingTemplate
    {
        public IfLinkInCorrectWorksetChecking()
        {
            Name = "ОБЩ_Корректность рабочих наборов связей";
            Dep = "ОБЩ";
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
