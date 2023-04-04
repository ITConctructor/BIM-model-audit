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
    public class IfLinkInCorrectWorksetChecking : CheckingTemplate
    {
        public IfLinkInCorrectWorksetChecking()
        {
            Name = "ОБЩ_Корректность рабочих наборов связей";
            Dep = "ОБЩ";
            ResultType = ApplicationViewModel.CheckingResultType.ElementsList;
        }
        public override ApplicationViewModel.CheckingStatus Run(Document doc, BindingList<ElementCheckingResult> oldResults)
        {
            //Получаем связи в файле
            IList<Element> links = new List<Element>();
            links = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType().ToElements();

            //Проверяем корректность рабочего набора связи, если документ поддерживает совместную работу
            IList<Element> results = new List<Element>();
            if (doc.IsWorkshared)
            {
                foreach (Element elem in links)
                {
                    if (!elem.LookupParameter("Рабочий набор").AsValueString().Contains("Связи") || !elem.Name.Contains(elem.LookupParameter("Рабочий набор").AsValueString().Substring(8)))
                    {
                        results.Add(elem);
                    }
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
