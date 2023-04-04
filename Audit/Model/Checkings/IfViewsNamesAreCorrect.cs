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
    public class IfViewsNamesAreCorrect : CheckingTemplate
    {
        public IfViewsNamesAreCorrect()
        {
            Name = "ОБЩ_Наличие копий видов";
            Dep = "ОБЩ";
            ResultType = ApplicationViewModel.CheckingResultType.ElementsList;
        }
        public override ApplicationViewModel.CheckingStatus Run(Document doc, BindingList<ElementCheckingResult> oldResults)
        {
            //Получение всех видов из файла
            IList<Element> viewPlans = new FilteredElementCollector(doc).OfClass(typeof(ViewPlan)).WhereElementIsNotElementType().ToElements();
            IList<Element> viewSections = new FilteredElementCollector(doc).OfClass(typeof(ViewSection)).WhereElementIsNotElementType().ToElements();
            IList<Element> views3D = new FilteredElementCollector(doc).OfClass(typeof(View3D)).WhereElementIsNotElementType().ToElements();
            IList<Element> shedules = new FilteredElementCollector(doc).OfClass(typeof(ViewSchedule)).WhereElementIsNotElementType().ToElements();

            IList<Element> results = new List<Element>();
            //Проверка планов
            foreach (Element plan in viewPlans)
            {
                if (plan.Name.EndsWith("копия 1") || plan.Name.EndsWith("(1)") || plan.Name.EndsWith("(1") || plan.Name.Contains("копия 1"))
                {
                    results.Add(plan);
                }
            }

            //Проверка фасадов и разрезов
            foreach (Element fas in viewSections)
            {
                if (fas.Name.EndsWith("копия 1") || fas.Name.EndsWith("(1)") || fas.Name.EndsWith("(1") || fas.Name.Contains("копия 1"))
                {
                    results.Add(fas);
                }
            }

            //Проверка 3D видов
            foreach (Element view3D in views3D)
            {
                if (view3D.Name.EndsWith("копия 1") || view3D.Name.EndsWith("(1)") || view3D.Name.EndsWith("(1") || view3D.Name.Contains("копия 1"))
                {
                    results.Add(view3D);
                }
            }

            //Проверка спецификаций
            foreach (Element schedule in shedules)
            {
                if (schedule.Name.EndsWith("копия 1") || schedule.Name.EndsWith("(1)") || schedule.Name.EndsWith("(1") || schedule.Name.Contains("копия 1"))
                {
                    results.Add(schedule);
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
