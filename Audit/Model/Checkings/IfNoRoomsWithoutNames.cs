﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.ComponentModel;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB.Architecture;

namespace Audit.Model.Checkings
{
    [Transaction(TransactionMode.Manual)]
    public class IfNoRoomsWithoutNames : CheckingTemplate
    {
        public IfNoRoomsWithoutNames()
        {
            Name = "АР_Нет помещений без имени";
            Dep = "АР";
            ResultType = ApplicationViewModel.CheckingResultType.ElementsList;
        }
        public override ApplicationViewModel.CheckingStatus Run(Document doc, BindingList<ElementCheckingResult> oldResults)
        {
            IList<Element> results = new List<Element>();

            //Проверяем, что в имени файла содержится "АР"
            if (doc.Title.Split(new string[] { "\\" }, StringSplitOptions.None).Last().Contains(Dep))
            {
                //Получаем все помещения
                IList<Element> rooms = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements();

                //Проверяем, имеет ли помещение имя "Помещение"
                foreach (Element room in rooms)
                {
                    if (room.LookupParameter("Имя").AsString() == "Помещение" || room.LookupParameter("Имя").AsString() == "")
                    {
                        results.Add(room);
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

            //Выводим результат проверки - пройдена или нет
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
