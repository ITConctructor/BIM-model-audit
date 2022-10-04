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

namespace Audit.Model.Checkings
{
    [Transaction(TransactionMode.Manual)]
    public class IfLevelsAreNotDuplicated : CheckingTemplate
    {
        public IfLevelsAreNotDuplicated()
        {
            Name = "ОБЩ_Уровни не дублируются";
            Dep = "ОБЩ";
        }
        public override void Run(string filePath, BindingList<ElementCheckingResult> oldResults)
        {
            //Открытие документа с отсоединением и закрытием всех рабочих наборов
            UIApplication uiapp = CommandLauncher.uiapp;
            ModelPath path = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath);
            OpenOptions openOptions = new OpenOptions();
            openOptions.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets;
            WorksetConfiguration worksets = new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets);
            openOptions.SetOpenWorksetsConfiguration(worksets);
            Application app = CommandLauncher.app;
            Document doc = app.OpenDocumentFile(path, openOptions);

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
            for (int i = 0; i < marks.Count; i++)
            {
                if (marks.Count(x => x.Equals(marks[i])) > 1)
                {
                    results.Add(levels[i]);
                }
            }

            //Из списка элементов заполняем отчет
            foreach (Element element in results)
            {
                ElementCheckingResult result = new ElementCheckingResult() { Name = element.Name, ID = element.Id.ToString(), Time = System.DateTime.Now.ToString() };
                ApplicationViewModel.AddElementCheckingResult(result, oldResults);
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
        }
    }
}
