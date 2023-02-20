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
    public class IfBasePointPlacementIsCorrect : CheckingTemplate
    {
        public IfBasePointPlacementIsCorrect()
        {
            Name = "ОБЩ_Положение базовой точки";
            Dep = "ОБЩ";
            ResultType = ApplicationViewModel.CheckingResultType.Message;
        }
        public override ApplicationViewModel.CheckingStatus Run(Document doc, BindingList<ElementCheckingResult> oldResults)
        {
            //Флаг для отображение результата проверки: пройдена или нет
            int flag = 0;
            //Получаем базовую точку проекта
            BasePoint basePoint = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProjectBasePoint).WhereElementIsNotElementType().ToElements()[0] as BasePoint;
            //Координаты базовой точке в системе координат проекта
            XYZ position = basePoint.Position;
            //Ищем оси 1/А
            IList<Element> grids = new FilteredElementCollector(doc).OfClass(typeof(Grid)).WhereElementIsNotElementType().ToElements();
            IList<Element> corgrids = new List<Element>();
            foreach (Grid grid in grids)
            {
                string name = grid.Name;
                if (name == "A" || name == "А" || name.Contains("A") || name.Contains("А") || name.StartsWith("1/") || name == "1")
                {
                    corgrids.Add(grid);
                }
                if (corgrids.Count == 2)
                {
                    break;
                }
            }
            //Ищем точку пересечения осей
            Curve line1 = ((Grid)corgrids[0]).Curve;
            Curve line2 = ((Grid)corgrids[1]).Curve;
            XYZ intersectPoint = new XYZ();
            IntersectionResultArray result = new IntersectionResultArray();
            if (line1.Intersect(line2, out result) != SetComparisonResult.Disjoint)
            {
                intersectPoint = result.get_Item(0).XYZPoint;
            }
            //Округляем координаты точек для предотвращения случайной ошибки
            if (intersectPoint != null)
            {
                double x1 = Math.Round(position.X, 3);
                double y1 = Math.Round(position.Y, 3);
                double x2 = Math.Round(intersectPoint.X, 3);
                double y2 = Math.Round(intersectPoint.Y, 3);
                //Сравниваем координаты и делаем проверку
                if (x1 == x2 && y1 == y2)
                {
                    flag = 1;
                }
            }
            //Возвращаем результат проверки - пройдена или нет
            if (flag == 1)
            {
                Message = "Базовая точка находится на пересечении осей 1/А";
                return ApplicationViewModel.CheckingStatus.CheckingSuccessful;
            }
            else
            {
                Message = "Базовая точка не находится на пересечении осей 1/А";
                return ApplicationViewModel.CheckingStatus.CheckingFailed;
            }
        }
    }
}
