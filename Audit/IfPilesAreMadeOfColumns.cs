﻿using System;
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
    internal class IfPilesAreMadeOfColumns : CheckingTemplate
    {
        public IfPilesAreMadeOfColumns()
        {
            Name = "КР_Сваи замоделированы колоннами";
            Dep = "КР";
        }
        public override Result Run()
        {
            TaskDialog dialog = new TaskDialog("Test");
            dialog.MainContent = Name;
            dialog.Show();
            return Result.Succeeded;
        }
    }
}
