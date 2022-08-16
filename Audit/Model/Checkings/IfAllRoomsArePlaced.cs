﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.ComponentModel;

namespace Audit.Checkings
{
    [Transaction(TransactionMode.Manual)]
    public class IfAllRoomsArePlaced : CheckingTemplate
    {
        public IfAllRoomsArePlaced()
        {
            Name = "АР_Нет неразмещенных помещений";
            Dep = "АР";
        }
        public override void Run(string filePath, BindingList<ElementCheckingResult> oldResults)
        {
            ElementCheckingResult newResult = new ElementCheckingResult() { Name = "elementName", ID = "elementID", Time = System.DateTime.Now.ToString() };
            ElementCheckingResult newResult2 = new ElementCheckingResult() { Name = "elementName2", ID = "elementID2", Time = System.DateTime.Now.ToString() };
            ApplicationViewModel.AddElementCheckingResult(newResult2, oldResults);
            ApplicationViewModel.AddElementCheckingResult(newResult, oldResults);
            TaskDialog dialog = new TaskDialog("Test");
            dialog.MainContent = Name;
            dialog.Show();
        }

    }
}

