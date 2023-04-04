using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.ComponentModel;
using Autodesk.Revit.ApplicationServices;
using static Audit.Model.Utils;

namespace Audit.Model.Checkings
{
    [Transaction(TransactionMode.Manual)]
    public class IfLevelsAreNotDuplicated : CheckingTemplate
    {
        public IfLevelsAreNotDuplicated()
        {
            Name = "ОБЩ_Уровни не дублируются";
            Dep = "ОБЩ";
            ResultType = ApplicationViewModel.CheckingResultType.ElementsList;
        }
        public override ApplicationViewModel.CheckingStatus Run(Document doc, BindingList<ElementCheckingResult> oldResults)
        {
            //Получаем уровни
            IList<Element> levels = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements();
            IList<Element> results = new List<Element>();

            //Получаем отметки уровней
            List<double> marks = new List<double>();
            foreach (Element level in levels)
            {
                marks.Add(Math.Round(level.LookupParameter("Фасад").AsDouble()*304.8, 0));
            }

            //Проверяем наличие повторяющихся отметок
            for (int i = 0; i < marks.Count; i = i)
            {
                double mark = marks[i];
                marks.RemoveAt(i);
                Element level = levels[i];
                levels.RemoveAt(i);
                if (marks.Contains(mark))
                {
                    results.Add(level);
                }
            }

            //Из списка элементов заполняем отчет
            foreach (Element element in results)
            {
                ElementCheckingResult result = new ElementCheckingResult() { Name = element.Name, ID = element.Id.ToString(), Time = System.DateTime.Now.ToString() };
                AddElementCheckingResult(result, oldResults);
            }

            //Проверяем, есть ли среди прошлого результата проверок какой-либо результат из новой. Если нет, то ставим этому результату статус "Исправленная"
            foreach (ElementCheckingResult item in oldResults)
            {
                int flag = 0;
                foreach (Element level in results)
                {
                    if (item.Name == level.Name)
                    {
                        flag = 1;
                    }
                    if (flag == 1)
                    {
                        break;
                    }
                }
                if (flag == 0)
                {
                    item.Status = "Исправленная";
                }
            }

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
