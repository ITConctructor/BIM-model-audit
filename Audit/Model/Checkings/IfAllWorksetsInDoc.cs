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
    public class IfAllWorksetsInDoc : CheckingTemplate
    {
        public IfAllWorksetsInDoc()
        {
            Name = "ОБЩ_Наличие рабочих наборов";
            Dep = "ОБЩ";
            ResultType = ApplicationViewModel.CheckingResultType.ElementsList;
        }
        public override ApplicationViewModel.CheckingStatus Run(Document doc, BindingList<ElementCheckingResult> oldResults)
        {
            IList<ElementCheckingResult> results = new List<ElementCheckingResult>();
            //Получение всех рабочих наборов
            IList<Workset> worksets = new FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets();

            //Составляем список из имен существующих рабочих наборов
            List<string> worksetnames = new List<string>();
            foreach (Workset set in worksets)
            {
                worksetnames.Add(set.Name);
            }

            //Список обязательных рабочих наборов
            List<string> allWorksets = new List<string>() { "#_Общие уровни и сетки", "#_Связи АР", "#_Связи КЖ", "#_Связи ОВ", "#_Связи ВК", "#_Связи СС", "#_Связи ЭО", "#_Связи DWG", "#_Задание входящее", "#_Задание исходящее" };

            //Проверяем наличие обязательного рабочего набора в списке существующих, и если он отсутствует, то записываем в результаты отчета
            foreach (string set in allWorksets)
            {
                if (!worksetnames.Contains(set))
                {
                    results.Add(new ElementCheckingResult() { Name=set, ID="-", Status="Созданная"});
                }
            }

            //Из списка элементов заполняем отчет
            foreach (ElementCheckingResult element in results)
            {
                AddElementCheckingResult(element, oldResults);
            }

            //Проверяем, есть ли среди прошлого результата проверок какой-либо результат из новой. Если нет, то ставим этому результату статус "Исправленная"
            foreach (ElementCheckingResult item in oldResults)
            {
                int flag = 0;
                foreach (ElementCheckingResult level in results)
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
