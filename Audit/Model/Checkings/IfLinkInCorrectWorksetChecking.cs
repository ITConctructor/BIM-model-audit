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
        public override void Run(string filePath, BindingList<ElementCheckingResult> oldResults)
        {
            //ElementCheckingResult newResult = new ElementCheckingResult() { Name = "elementName", ID = "elementID", Time = System.DateTime.Now.ToString() };
            //ElementCheckingResult newResult2 = new ElementCheckingResult() { Name = "elementName2", ID = "elementID2", Time = System.DateTime.Now.ToString() };
            //ApplicationViewModel.AddElementCheckingResult(newResult2, oldResults);
            //ApplicationViewModel.AddElementCheckingResult(newResult, oldResults);
            //TaskDialog dialog = new TaskDialog("Test");
            //dialog.MainContent = Name;
            //dialog.Show();

            //Открытие документа с отсоединением
            UIApplication uiapp = CommandLauncher.uiapp;
            ModelPath path = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath);
            OpenOptions openOptions = new OpenOptions();
            openOptions.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets;
            WorksetConfiguration worksets = new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets);
            openOptions.SetOpenWorksetsConfiguration(worksets);
            Application app = CommandLauncher.app;
            Document doc = app.OpenDocumentFile(path, openOptions);
            
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
