using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.ComponentModel;
using static Audit.Model.Utils;

namespace Audit.Model.Checkings
{
    [Transaction(TransactionMode.Manual)]
    public class IfPilesAreMadeOfColumns : CheckingTemplate
    {
        public IfPilesAreMadeOfColumns()
        {
            Name = "КР_Сваи замоделированы колоннами";
            Dep = "КР";
            ResultType = ApplicationViewModel.CheckingResultType.ElementsList;
        }
        public override ApplicationViewModel.CheckingStatus Run(Document doc, BindingList<ElementCheckingResult> oldResults)
        {
            ElementCheckingResult newResult = new ElementCheckingResult() { Name = "elementName", ID = "elementID", Time = System.DateTime.Now.ToString() };
            ElementCheckingResult newResult2 = new ElementCheckingResult() { Name = "elementName2", ID = "elementID2", Time = System.DateTime.Now.ToString() };
            AddElementCheckingResult(newResult2, oldResults);
            AddElementCheckingResult(newResult, oldResults);
            IList<Element> results = new List<Element>();

            //Возвращаем результат проверки - пройдена или нет
            if (results.Count == 0)
            {
                return ApplicationViewModel.CheckingStatus.CheckingSuccessful;
            }
            else
            {
                return ApplicationViewModel.CheckingStatus.CheckingFailed;
            }
        }
    }
}
