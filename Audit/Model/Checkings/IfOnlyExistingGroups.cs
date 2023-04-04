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
    public class IfOnlyExistingGroups : CheckingTemplate
    {
        public IfOnlyExistingGroups()
        {
            Name = "ОБЩ_Нет неразмещенных групп";
            Dep = "ОБЩ";
            ResultType = ApplicationViewModel.CheckingResultType.ElementsList;
        }
        public override ApplicationViewModel.CheckingStatus Run(Document doc, BindingList<ElementCheckingResult> oldResults)
        {
            IList<Element> results = new List<Element>();
            //Получение всех типов групп
            IList<Element> groupTypes = new FilteredElementCollector(doc).OfClass(typeof(GroupType)).ToElements();
            //Смотрим, список экземпляров групп какой группы пуст и записываем эту группу в список результатов
            foreach (Element group in groupTypes)
            {
                GroupType groupType = group as GroupType;
                if (groupType.Groups != null)
                {
                    if (groupType.Groups.IsEmpty)
                    {
                        results.Add(group);
                    }
                }
                else
                {
                    results.Add(group);
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
