using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using static Audit.ApplicationViewModel;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace Audit.Model
{
    public class ReportSettings
    {
        public ReportFields Fields { get; set; } = new ReportFields();
        public ReportStatuses Statuses { get; set; } = new ReportStatuses();
        public string Type { get; set; }
        public string Format { get; set; }
        public bool WriteForSelectedFile { get; set; }

        /// <summary>
        /// Создает отчеты о проверках
        /// </summary>
        public void CreateReport(ApplicationViewModel ViewModel)
        {
            if (Format.ToString() == "Excel")
            {
                CreateExcel(ViewModel);
            }
        }
        /// <summary>
        /// Создает результаты проверок в формат excel
        /// </summary>
        private void CreateExcel(ApplicationViewModel ViewModel)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.Description = "Выберите папку для сохранения отчетов";
            if (Properties.Settings.Default.LastReportPath != "")
            {
                folderBrowserDialog.SelectedPath = Properties.Settings.Default.LastReportPath;
            }
            folderBrowserDialog.ShowDialog();
            string path = folderBrowserDialog.SelectedPath;
            Properties.Settings.Default.LastReportPath = path;
            Properties.Settings.Default.Save();
            Excel.Application app = new Excel.Application();
            app.DisplayAlerts = false;
            if (path != "" && path != null)
            {
                if (WriteForSelectedFile == true)
                {
                    RvtFileInfo selectedFile = ViewModel.SelectedFile;
                    Excel.Workbook wb = app.Workbooks.Add();
                    wb.Title = selectedFile.Name;
                    switch (Type)
                    {
                        case "Все тесты":
                            foreach (CheckingTemplate checking in selectedFile.CheckingResults.Last().Checkings)
                            {
                                CreateExcelCheckingWorksheet(wb, checking);
                            }
                            break;
                        case "Выбранные тесты":
                            foreach (CheckingTemplate checking in selectedFile.CheckingResults.Last().Checkings)
                            {
                                foreach (string checkingName in ViewModel.SelectedCheckingsNames)
                                {
                                    if (checking.Name == checkingName)
                                    {
                                        CreateExcelCheckingWorksheet(wb, checking);
                                    }
                                }
                            }
                            break;
                        default:
                            break;
                    }
                    string fullPath = path + "\\" + selectedFile.Name.Replace(".rvt", ".xlsx");
                    while (File.Exists(fullPath))
                    {
                        fullPath = fullPath.Replace(".xlsx", "") + "(1)" + ".xlsx";
                    }
                    wb.SaveAs(fullPath);
                    app.Quit();
                }
                else
                {
                    foreach (RvtFileInfo file in ViewModel.PreanalysFiles)
                    {
                        Excel.Workbook wb = app.Workbooks.Add();
                        wb.Title = file.Name;
                        if (Type.ToString() == "Все тесты")
                        {
                            foreach (CheckingTemplate checking in file.CheckingResults[0].Checkings)
                            {
                                CreateExcelCheckingWorksheet(wb, checking);
                            }
                        }
                        else if (Type.ToString() == "Выбранные тесты")
                        {
                            foreach (CheckingTemplate checking in file.CheckingResults.Last().Checkings)
                            {
                                foreach (string checkingName in ViewModel.SelectedCheckingsNames)
                                {
                                    if (checking.Name == checkingName)
                                    {
                                        CreateExcelCheckingWorksheet(wb, checking);
                                    }
                                }
                            }
                        }
                        string fullPath = path + "\\" + file.Name.Replace(".rvt", ".xlsx");
                        while (File.Exists(fullPath))
                        {
                            fullPath = fullPath.Replace(".xlsx", "") + "(1)" + ".xlsx";
                        }
                        wb.SaveAs(fullPath);
                        app.Quit();
                    }
                }

            }
        }
        /// <summary>
        /// Создает лист с результатом проверки для одной проверки одного файла. Метод используется для предотвращения дублирования кода внутри CreateExcel()
        /// </summary>
        private void CreateExcelCheckingWorksheet(Excel.Workbook wb, CheckingTemplate checking)
        {
            Excel.Worksheet sheet = wb.Worksheets.Add();
            sheet.Name = checking.Name.Substring(0, Math.Min(30, checking.Name.Length));
            if (checking.ResultType == CheckingResultType.ElementsList)
            {
                ///Создаем булевые маски из настроек отчета и по ним фильтруем список с информацией о файлах
                bool[] fieldBoolMask = { Fields.Name,
                            Fields.Status,
                            Fields.Id,
                            Fields.Comment,
                            Fields.Time };
                List<string> rawFields = new List<string>() { "Имя", "Статус", "ID элемента", "Комментарий", "Время обнаружения" };
                List<string> readyFields = rawFields.Where(x => fieldBoolMask[rawFields.IndexOf(x)] == true).ToList();
                bool[] statusesBoolMask = { Statuses.Created,
                        Statuses.Active, Statuses.Checked, Statuses.Corrected};
                List<string> rawStatuses = new List<string>() { "Созданная", "Активная", "Проверенная", "Исправленная" };
                List<string> readyStatuses = rawStatuses.Where(x => statusesBoolMask[rawStatuses.IndexOf(x)] == true).ToList();
                List<List<string>> data = new List<List<string>>();
                CheckingStatusToTextConverter statusConverter = new CheckingStatusToTextConverter();
                string strStatus = (string)statusConverter.Convert(checking.Status, typeof(string), new object(), new CultureInfo(0x040A, false));
                data.Add(new List<string>() { strStatus });
                data.Add(readyFields);
                foreach (ElementCheckingResult elementResult in checking.ElementCheckingResults.Where(x => readyStatuses.Contains(x.Status)))
                {
                    List<string> values = new List<string>();
                    for (int i = 0; i < elementResult.GetType().GetTypeInfo().GetProperties().Length; i++)
                    {
                        if (fieldBoolMask[i] == true)
                        {
                            values.Add((string)elementResult.GetType().GetTypeInfo().GetProperties()[i].GetValue(elementResult));
                        }
                    }
                    data.Add(values);
                }
                for (int i = 0; i < data.Count; i++)
                {
                    for (int j = 0; j < data[i].Count; j++)
                    {
                        sheet.Cells[i + 1, j + 1] = data[i][j];
                    }
                }
            }
            else if (checking.ResultType == CheckingResultType.Message)
            {
                sheet.Cells[1, 1] = checking.Message;
            }
        }

        public class ReportFields
        {
            public bool Name { get; set; }
            public bool Status { get; set; }
            public bool Id { get; set; }
            public bool Comment { get; set; }
            public bool Time { get; set; }
        }

        public class ReportStatuses
        {
            public bool Created { get; set; }
            public bool Active { get; set; }
            public bool Checked { get; set; }
            public bool Corrected { get; set; }
        }
    }
}
