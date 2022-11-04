using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using System.ComponentModel;
using System.Reflection;
using System.Text.RegularExpressions;

namespace Audit.Model.Checkings
{
    [Transaction(TransactionMode.Manual)]
    public class IfLevelNamesAreCorrectChecking : CheckingTemplate
    {
        public IfLevelNamesAreCorrectChecking()
        {
            Name = "ОБЩ_Корректность наименований уровней";
            Dep = "ОБЩ";
        }
        public override void Run(Document doc, BindingList<ElementCheckingResult> oldResults)
        {
            //Получаем уровни
            IList<Element> levels = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements();
            IList<Element> good_levels = new List<Element>();
            IList<Element> results = new List<Element>();
            //Сортировка списка с уровнями по отметке
            List<double> marks = new List<double>();
            foreach (Element level in levels)
            {
                double num = double.Parse(level.LookupParameter("Фасад").AsValueString());
                marks.Add(num);
            }
            for (int i = 0; i < marks.Count; i++)
            {
                int j = i;
                while (j > 0 && marks[j - 1] > marks[j])
                {
                    double buf = marks[j - 1];
                    Element buf1 = levels[j - 1];
                    marks[j - 1] = marks[j];
                    levels[j - 1] = levels[j];
                    marks[j] = buf;
                    levels[j] = buf1;
                    j--;
                }
            }
            //Отбираем уровни, название которых состоит из 3-х блоков, разделенных "_"
            foreach (Element level in levels)
            {
                if (new Regex("_").Matches(level.Name).Count == 2)
                {
                    good_levels.Add(level);
                }
            }
            //Проверяем, заключена ли в скобки отметка уровня
            foreach (Element level in good_levels)
            {
                string[] name_array = level.Name.Split('_');
                if (!name_array[2].StartsWith("(") || !name_array[2].EndsWith(")"))   
                {
                    results.Add(level);
                    //good_levels.Remove(level);
                }
            }
            ///Проверяем правильность отметки в названии
            foreach (Element level in good_levels)
            {
                string[] name_array = level.Name.Split('_');
                string real_mark = name_array[2].Substring(1, name_array[2].Length - 2);
                string good_mark;
                double mark = level.LookupParameter("Фасад").AsDouble() * 304.8;
                if (Math.Round(mark, 0) == 0)
                {
                    good_mark = "+0,000";
                    if (real_mark != good_mark)
                    {
                        results.Add(level);
                        //good_levels.Remove(level);
                    }
                }
                else if (Math.Round(mark, 0) > 0)
                {
                    good_mark = "";
                    if (Math.Round(mark / 1000, 3).ToString().Contains(","))
                    {
                        good_mark = "+" + Math.Round(mark / 1000, 3).ToString() + "00";
                    }
                    else
                    {
                        good_mark = "+" + Math.Round(mark / 1000, 3).ToString() + ",000";
                    }
                    if (real_mark != good_mark)
                    {
                        results.Add(level);
                        //good_levels.Remove(level);
                    }
                }
                else
                {
                    good_mark = "";
                    if (Math.Round(mark / 1000, 3).ToString().Contains(","))
                    {
                        good_mark = Math.Round(mark / 1000, 3).ToString() + "00";
                    }
                    else
                    {
                        good_mark = Math.Round(mark / 1000, 3).ToString() + ",000";
                    }
                    if (real_mark != good_mark)
                    {
                        results.Add(level);
                        //good_levels.Remove(level);
                    }
                }
            }
            //Распределение уровней по назначению
            IList<Element> floors = new List<Element>();
            IList<Element> pits = new List<Element>();
            IList<Element> roofs = new List<Element>();
            IList<Element> grounds = new List<Element>();
            IList<Element> techs = new List<Element>();
            IList<Element> parapets = new List<Element>();
            IList<Element> supports = new List<Element>();
            foreach (Element level in good_levels)
            {
                string[] name_array = level.Name.Split('_');
                if (name_array[1].StartsWith("Этаж"))
                {
                    floors.Add(level);
                }
                else if (name_array[1] == "Приямки")
                {
                    pits.Add(level);
                }
                else if (name_array[1].StartsWith("Кровля"))
                {
                    roofs.Add(level);
                }
                else if (name_array[1] == "Земляные работы")
                {
                    grounds.Add(level);
                }
                else if (name_array[1].StartsWith("Тех"))
                {
                    techs.Add(level);
                }
                else if (name_array[1] == "Парапет")
                {
                    parapets.Add(level);
                }
                else
                {
                    supports.Add(level);
                }
            }
            //Создание спика для уровней, прошедших проверку
            IList<Element> good_floors = new List<Element>();
            //Анализ этажей
            foreach (Element floor in floors)
            {
                try
                {
                    if (floor.Name.Split('_')[1].Substring(4).StartsWith(" "))
                    {
                        string code = " ";
                        if (double.Parse(floor.LookupParameter("Фасад").AsValueString()) < 0)
                        {
                            if (floor.Name.Split('_')[1].Substring(6).StartsWith("0") && floor.Name.Split('_')[1].Substring(7) != "0")
                            {
                                int num_1 = int.Parse(floor.Name.Split('_')[1].Substring(7)) + 1;
                                string num = num_1.ToString();
                                code = "00" + num;
                            }
                            else if ((int.Parse(floor.Name.Split('_')[1].Substring(6)) + 1).ToString().Length == 2)
                            {
                                code = "0" + (int.Parse(floor.Name.Split('_')[1].Substring(6)) + 1).ToString();
                            }
                            if (floor.Name.Split('_')[0] != code)
                            {
                                results.Add(floor);
                            }
                            else
                            {
                                code = " ";
                            }
                        }
                        else
                        {
                            code = "1" + floor.Name.Split('_')[1].Substring(5);
                            if (floor.Name.Split('_')[0] != code)
                            {
                                results.Add(floor);
                            }
                            else
                            {
                                code = " ";
                            }
                        }
                    }
                    else
                    {
                        results.Add(floor);
                    }
                }
                catch (Exception)
                {
                    results.Add(floor);
                }
            }
            //Анализ технических этажей
            IList<Element> good_techs = new List<Element>();
            foreach (Element tech in techs)
            {
                good_techs.Add(tech);
                if (double.Parse(tech.LookupParameter("Фасад").AsValueString()) < 0)
                {
                    if (tech.Name.Split('_')[0] != "001")
                    {
                        results.Add(tech);
                    }
                }
                else
                {
                    Element underLevel = levels[levels.IndexOf(tech) - 1];
                    if (tech.Name.Split('_')[0] != underLevel.Name.Substring(0, 3))
                    {
                        results.Add(tech);
                    }
                }
            }
            //Анализ кровель
            IList<Element> good_roofs = new List<Element>();
            foreach (Element roof in roofs)
            {
                string code = "";
                good_roofs.Add(roof);
                if (roofs.Count == 1)
                {
                    if (roof.Name.Split('_')[1] != "Кровля" || roof.Name.Split('_')[0] != "201")
                    {
                        results.Add(roof);
                    }
                }
                else
                {
                    //try нужен на случай, если substring вывалится с исключением, что будет означать, что название кровли не содержит 01 или 02
                    try
                    {
                        if (roof.Name.Split('_')[1].Substring(6).Length > 0)
                        {
                            code = "2" + roof.Name.Split('_')[1].Substring(7);
                            if (roof.Name.Split('_')[0] != code)
                            {
                                results.Add(roof);
                            }
                        }
                        else
                        {
                            results.Add(roof);
                        }
                    }
                    catch (Exception)
                    {
                        results.Add(roof);
                    }
                    
                }
            }
            //Анализ парапетов
            IList<Element> good_parapets = new List<Element>();
            foreach (Element par in parapets)
            {
                Element underLevel = levels[levels.IndexOf(par) - 1];
                string code = "";
                good_parapets.Add(par);
                code = "20" + (parapets.IndexOf(par) + 1).ToString();
                if (par.Name.Split('_')[0] != code && roofs.Count >= parapets.Count)
                {
                    results.Add(par);
                }
                else if (!underLevel.Name.Split('_')[1].StartsWith("Кровля"))
                {
                    results.Add(par);
                }
            }
            //Анализ уровня земляных работ
            IList<Element> good_grounds = new List<Element>();
            foreach (Element ground in grounds)
            {
                good_grounds.Add(ground);
                if (ground.Name.Split('_')[0] != "000")
                {
                    results.Add(ground);
                }
            }
            //Анализ уровней приямков
            IList<Element> good_pits = new List<Element>();
            foreach (Element pit in pits)
            {
                Element aboveLevel = levels[levels.IndexOf(pit) + 1];
                if (aboveLevel.Name.Substring(0, 3) != pit.Name.Split('_')[0])
                {
                    results.Add(pit);
                }
            }
            //Анализ вспомогательных уровней
            foreach (Element support in supports)
            {
                if (support.Name.Split('_')[0] != "Вспом")
                {
                    results.Add(support);
                }
            }
            //Проходимся по некорректно названным уровням и для каждого такого уровня создаем результат в отчете о проверке
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
