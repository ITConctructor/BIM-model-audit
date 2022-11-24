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
using Autodesk.Revit.DB.Architecture;

namespace Audit.Model.Checkings
{
    [Transaction(TransactionMode.Manual)]
    public class IfARWallsNamesAreCorrect : CheckingTemplate
    {
        public IfARWallsNamesAreCorrect()
        {
            Name = "АР_Корректность наименований типов стен";
            Dep = "АР";
            ResultType = ApplicationViewModel.CheckingResultType.ElementsList;
        }
        public override ApplicationViewModel.CheckingStatus Run(Document doc, BindingList<ElementCheckingResult> oldResults)
        {
            IList<Element> results = new List<Element>();

            ////Проверяем, что в имени файла содержится "АР"
            //if (doc.Title.Split(new string[] { "\\" }, StringSplitOptions.None).Last().Contains(Dep))
            //{
            //    //Получаем все типы стен
            //    IList<Element> walls = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsElementType().ToElements();

            //    //Отсеиваем стены, имена которых состоят менее чем из трех частей, разделенных "_"
            //    IList<Element> goodWalls = new List<Element>();
            //    foreach (Element wall in walls)
            //    {
            //        string[] nameArray = wall.Name.Split('_');
            //        if (nameArray.Length < 3)
            //        {
            //            results.Add(wall);
            //        }
            //        else
            //        {
            //            goodWalls.Add(wall);
            //        }

            //    }

            //    //Массив возможных вариантов функциональной принадлежности стены
            //    string[] func = new string[] { "В", "Н", "ВО", "НО", "КР" };

            //    //Проверяем оставшиеся стены
            //    foreach (Element wall1 in goodWalls)
            //    {
            //        Wall wall = wall1 as Wall;
            //        HostObjAttributes host = (HostObjAttributes)wall1;
            //        string width = Math.Round(wall.Width * 304.8, 0).ToString();
            //        IList<CompoundStructureLayer> layers = host.GetCompoundStructure().GetLayers();
            //        List<string> nameArray = wall.Name.Split('_').ToList();
            //        if (!func.Contains(nameArray[0]))
            //        {
            //            results.Add(wall1);
            //        }
            //        else if (nameArray[1] != width)
            //        {
            //            results.Add(wall1);
            //        }
            //        else if (layers.Count == 1)
            //        {
            //            //string last = string.Join("_", nameArray.GetRange(2, nameArray.Count - 1).ToArray());
            //            string last = nameArray[2];
            //            string lowlast = last.ToLower();
            //            string materialName = doc.GetElement(layers[0].MaterialId).Name;
            //            string lowMaterialName = materialName.ToLower();
            //            if (!materialName.Contains(lowlast)) 
            //            {
            //                results.Add(wall1);
            //            }
                        
            //        }
            //    }

            //    //Получаем все типы составных стен
            //    IList<Element> stackedWalls = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StackedWalls).WhereElementIsElementType().ToElements();

            //    //Отсеиваем стены, имена которых состоят менее чем из шести частей, разделенных "_"
            //    IList<Element> goodStackedWalls = new List<Element>();
            //    foreach (Element wall in walls)
            //    {
            //        string[] nameArray = wall.Name.Split('_');
            //        if (nameArray.Length < 6)
            //        {
            //            results.Add(wall);
            //        }
            //        else
            //        {
            //            goodStackedWalls.Add(wall);
            //        }

            //    }

            //    //Проверяем оставшиеся составные стены
            //    //foreach (Element wall1 in goodStackedWalls)
            //    //{
            //    //    Wall wall = wall1 as Wall;
            //    //    IList<ElementId> membersIds = wall.GetStackedWallMemberIds();
            //    //    List<string> nameArray = wall.Name.Split('_').ToList();
            //    //    if (!func.Contains(nameArray[0]))
            //    //    {
            //    //        results.Add(wall1);
            //    //    }
            //    //    else if (nameArray[1] != width)
            //    //    {
            //    //        results.Add(wall1);
            //    //    }
            //    //    else if (layers.Count == 1)
            //    //    {
            //    //        //string last = string.Join("_", nameArray.GetRange(2, nameArray.Count - 1).ToArray());
            //    //        string last = nameArray[2];
            //    //        string lowlast = last.ToLower();
            //    //        string materialName = doc.GetElement(layers[0].MaterialId).Name;
            //    //        string lowMaterialName = materialName.ToLower();
            //    //        if (!materialName.Contains(lowlast))
            //    //        {
            //    //            results.Add(wall1);
            //    //        }

            //    //    }
            //    //}
            //}

            ////Из списка элементов заполняем отчет
            //foreach (Element element in results)
            //{
            //    ElementCheckingResult result = new ElementCheckingResult() { Name = element.Name, ID = element.Id.ToString(), Time = System.DateTime.Now.ToString() };
            //    ApplicationViewModel.AddElementCheckingResult(result, oldResults);
            //}

            ////Проверяем, есть ли среди прошлого результата проверок какой-либо результат из новой. Если нет, то ставим этому результату статус "Исправленная"
            //foreach (ElementCheckingResult item in oldResults)
            //{
            //    int flag = 0;
            //    foreach (Element level in results)
            //    {
            //        if (item.Name == level.Name)
            //        {
            //            flag = 1;
            //        }
            //        if (flag == 1)
            //        {
            //            break;
            //        }
            //    }
            //    if (flag == 0)
            //    {
            //        item.Status = "Исправленная";
            //    }
            //}

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
