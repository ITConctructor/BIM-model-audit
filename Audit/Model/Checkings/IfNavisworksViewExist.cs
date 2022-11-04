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
    public class IfNavisworksViewExist : CheckingTemplate
    {
        public IfNavisworksViewExist()
        {
            Name = "ОБЩ_Наличие вида Navisworks";
            Dep = "ОБЩ";
            ResultType = ApplicationViewModel.CheckingResultType.Message;
        }
        public override ApplicationViewModel.CheckingStatus Run(Document doc, BindingList<ElementCheckingResult> oldResults)
        {
            //Получение всех 3D видов из файла
            IList<Element> views = new FilteredElementCollector(doc).OfClass(typeof(View3D)).WhereElementIsNotElementType().ToElements();

            //Проверка наличия вида Navisworks
            int flag = 0;
            foreach (Element element in views)
            {
                if (element.Name == "Navisworks")
                {
                    flag = 1;
                    break;
                }
            }

            //Возвращаем результат проверки - пройдена или нет
            if (flag == 1)
            {
                Message = "Вид Navisworks существует";
                return ApplicationViewModel.CheckingStatus.CheckingSuccessful;
            }
            else
            {
                Message = "Вид Navisworks не существует";
                return ApplicationViewModel.CheckingStatus.CheckingFailed;
            }
        }
    }
}
