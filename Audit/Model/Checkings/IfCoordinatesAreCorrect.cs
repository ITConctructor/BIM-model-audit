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
using System.IO;
using static Audit.Model.Utils;

namespace Audit.Model.Checkings
{
    [Transaction(TransactionMode.Manual)]
    public class IfCoordinatesAreCorrect : CheckingTemplate
    {
        public IfCoordinatesAreCorrect()
        {
            Name = "ОБЩ_Корректность координат";
            Dep = "ОБЩ";
            ResultType = ApplicationViewModel.CheckingResultType.Message;
        }
        public override ApplicationViewModel.CheckingStatus Run(Document doc, BindingList<ElementCheckingResult> oldResults)
        {
            //Получение всех связей из файла
            IList<Element> links = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType().ToElements();
            //Флаг для определения результата проверки: пройдена или нет
            int flag = 3;
            
            ////Проверяем наличие связей
            //if (links.Count == 0)
            //{
            //    flag = 0;
            //}
            ////Получаем документы из связей
            ////Чтобы получить документы из связей, берем имена всех связанных файлов и ищем их в том же каталоге,
            ////где лежит базовый файл, углубляясь на 2 папки
            //List<Document> linkdocs = new List<Document>();
            //List<string> names = new List<string>();
            //foreach (Element rvtlink in links)
            //{
            //    RevitLinkInstance link = (RevitLinkInstance)rvtlink;
            //    if (!link.Name.ToLower().Contains("задани") || !link.Name.ToLower().Contains("отверсти"))
            //    {
            //        names.Add(link.Name);
            //    }
            //}
            //foreach (string name in names)
            //{
            //    string Name = name.Split(' ')[0];
            //    string filePath = Properties.Settings.Default.CurrentFile;
            //    string dir = Path.GetDirectoryName(Path.GetDirectoryName(filePath));
            //    string[] files = Directory.GetFiles(dir, "*.*", SearchOption.AllDirectories);
            //    string linkPath = "";
            //    foreach (string file in files)
            //    {
            //        if (file.Contains(Name))
            //        {
            //            linkPath = file;
            //            break;
            //        }
            //    }
            //    if (linkPath != "")
            //    {
            //        Document document = ApplicationViewModel.OpenRvtFile(linkPath);
            //        linkdocs.Add(document);
            //    }
            //}
            ////Получаем систему координат проекта
            //SiteLocation doclocation = doc.ActiveProjectLocation.GetSiteLocation();

            ////Просматриваем все связи и сравниваем их системы координат с системой координат проекта
            //int errorscount = 0;
            //foreach (Document linkdoc in linkdocs)
            //{
            //    SiteLocation linklocation = linkdoc.ActiveProjectLocation.GetSiteLocation();
            //    if (Math.Round(linklocation.Latitude, 3) != Math.Round(doclocation.Latitude, 3) || Math.Round(linklocation.Longitude, 3) != Math.Round(doclocation.Longitude, 3))
            //    {
            //        errorscount++;
            //    }
            //}
            //if (errorscount == linkdocs.Count && linkdocs.Count != 0)
            //{
            //    flag = 1;
            //}

            ////Проверяем, стоят ли связи по общим координатам
            //if (flag == 3)
            //{
            //    foreach (Element link in links)
            //    {
            //        if (link.Name.Contains("Не общедоступное"))
            //        {
            //            flag = 2;
            //            break;
            //        }
            //    }
            //}

            //Возвращаем результат проверки
            if (flag == 0)
            {
                Message = "В файле отсутствуют связи, проверка невозможна";
                return ApplicationViewModel.CheckingStatus.NotLaunched;
            }
            else if (flag == 1)
            {
                Message = "Координаты сломаны: модель и связи имеют разные системы координат";
                return ApplicationViewModel.CheckingStatus.CheckingFailed;
            }
            else if (flag == 2)
            {
                Message = "Координаты корректны, но не все модели стоят по общим координатам";
                return ApplicationViewModel.CheckingStatus.CheckingSuccessful;
            }
            else
            {
                Message = "Координаты корректны, модель имеет общие координаты";
                return ApplicationViewModel.CheckingStatus.CheckingSuccessful;
            }
        }
    }
}
